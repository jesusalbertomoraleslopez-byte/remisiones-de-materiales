"""
=============================================================================
SIGRAMA REMISIONES DE MATERIALES - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia y Almacenamiento 100% Google Cloud Storage (GCP)
Optimizado con Cliente Singleton, Caché en Memoria y Sincronización No Bloqueante
=============================================================================
"""

import os
import time
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

BASE_DIR = Path(__file__).resolve().parent

# Estado singleton a nivel de proceso de servidor (compartido entre sesiones en el mismo contenedor)
_CLIENT = None
_BUCKET = None
_BUCKET_LOCK = threading.Lock()
_SYNC_LOCK = threading.Lock()
_LAST_SYNC_TS = 0.0
_SYNC_TTL_SECONDS = 180.0  # Evita re-listar el bucket en cada pestaña de operador

# Índice en memoria de imágenes existentes en GCS: {sku_upper: blob_name}
_KNOWN_IMAGE_BLOBS = {}
_IMAGES_INDEXED = False


def _get_bucket_name() -> str:
    raw = os.environ.get("GCS_BUCKET", "sigrama-remisiones-storage")
    if not raw or not str(raw).strip():
        return "sigrama-remisiones-storage"
    return str(raw).split()[0].strip()


def _gcs_bucket():
    """Retorna una instancia singleton del Bucket de GCS reutilizando el pool de conexiones HTTP."""
    global _CLIENT, _BUCKET
    if _BUCKET is not None:
        return _BUCKET
    with _BUCKET_LOCK:
        if _BUCKET is not None:
            return _BUCKET
        from google.cloud import storage
        bucket_name = _get_bucket_name()
        try:
            import streamlit as st
            if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
                from google.oauth2 import service_account
                creds = service_account.Credentials.from_service_account_info(
                    dict(st.secrets["gcp_service_account"])
                )
                _CLIENT = storage.Client(credentials=creds)
                _BUCKET = _CLIENT.bucket(bucket_name)
                return _BUCKET
        except Exception:
            pass
        _CLIENT = storage.Client()
        _BUCKET = _CLIENT.bucket(bucket_name)
        return _BUCKET


def _extract_sku_from_filename(filename: str) -> str:
    if "(" in filename:
        return filename.split("(")[0].strip()
    return Path(filename).stem.strip()


def sync_from_gcs(force: bool = False) -> bool:
    """
    Sincroniza los archivos Excel de inmediato y descarga las imágenes en segundo plano.
    Si el contenedor ya se sincronizó recientemente, retorna al instante (0ms).
    """
    global _LAST_SYNC_TS, _IMAGES_INDEXED
    now = time.time()
    if not force and _LAST_SYNC_TS > 0 and (now - _LAST_SYNC_TS) < _SYNC_TTL_SECONDS:
        return True

    with _SYNC_LOCK:
        now = time.time()
        if not force and _LAST_SYNC_TS > 0 and (now - _LAST_SYNC_TS) < _SYNC_TTL_SECONDS:
            return True

        bucket_name = _get_bucket_name()
        if not bucket_name:
            return False

        try:
            bucket = _gcs_bucket()
            excel_blobs = []
            image_blobs = []
            valid_exts = ('.png', '.jpg', '.jpeg', '.webp')

            for blob in bucket.list_blobs():
                if blob.name.endswith("/"):
                    continue
                local_path = BASE_DIR / blob.name
                # Indexar imágenes en memoria para evitar búsquedas de red en SKUs sin imagen
                if blob.name.startswith("imagenes_articulos/"):
                    fname = Path(blob.name).name
                    if fname.lower().endswith(valid_exts):
                        sku_key = _extract_sku_from_filename(fname)
                        if sku_key:
                            _KNOWN_IMAGE_BLOBS[sku_key] = blob.name
                            _KNOWN_IMAGE_BLOBS[sku_key.upper()] = blob.name

                if not local_path.exists() or local_path.stat().st_size != blob.size:
                    if blob.name.lower().endswith(".xlsx") or blob.name == "consecutivo_override.txt":
                        excel_blobs.append((blob, local_path))
                    else:
                        image_blobs.append((blob, local_path))

            _IMAGES_INDEXED = True

            def _download(item):
                b, lp = item
                try:
                    lp.parent.mkdir(parents=True, exist_ok=True)
                    b.download_to_filename(str(lp))
                except Exception as e_dl:
                    print(f"[GCS] Error descargando {b.name}: {e_dl}")

            # 1. Descargar Excels de inmediato (son ~10 archivos ligeros, < 0.3s)
            if excel_blobs:
                with ThreadPoolExecutor(max_workers=10) as executor:
                    list(executor.map(_download, excel_blobs))
                print(f"[GCS] {len(excel_blobs)} archivos base (.xlsx) sincronizados al instante.")

            # 2. Descargar imágenes faltantes en hilo secundario para NO bloquear la apertura de la app
            if image_blobs:
                def _bg_download_images(items):
                    with ThreadPoolExecutor(max_workers=16) as bg_exec:
                        list(bg_exec.map(_download, items))
                    print(f"[GCS] {len(items)} imágenes sincronizadas en segundo plano.")

                threading.Thread(target=_bg_download_images, args=(image_blobs,), daemon=True).start()

            _LAST_SYNC_TS = time.time()
            return True
        except Exception as e:
            print(f"[GCS] Error al sincronizar desde gs://{bucket_name}: {e}")
            return False


def push_excel_to_gcs(file_name: str) -> bool:
    """Sube un archivo Excel específico a GCS reutilizando el cliente singleton."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        p = BASE_DIR / file_name if not Path(file_name).is_absolute() else Path(file_name)
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        print(f"[GCS] Archivo Excel subido: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al subir {file_name} a GCS: {e}")
        return False


def push_image_to_gcs(image_path: str) -> bool:
    """Sube una imagen de SKU a GCS y actualiza el índice en memoria."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        p = Path(image_path)
        if not p.is_absolute():
            p = BASE_DIR / image_path
        if not p.exists():
            return False
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        bucket.blob(rel_path).upload_from_filename(str(p))
        sku_key = _extract_sku_from_filename(p.name)
        if sku_key:
            _KNOWN_IMAGE_BLOBS[sku_key] = rel_path
            _KNOWN_IMAGE_BLOBS[sku_key.upper()] = rel_path
        print(f"[GCS] Imagen subida a GCS: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error subiendo imagen a GCS: {e}")
        return False


def delete_image_from_gcs(image_path: str) -> bool:
    """Elimina una imagen de SKU en GCS y limpia el índice en memoria."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        p = Path(image_path)
        if not p.is_absolute():
            p = BASE_DIR / image_path
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        sku_key = _extract_sku_from_filename(p.name)
        if sku_key:
            _KNOWN_IMAGE_BLOBS.pop(sku_key, None)
            _KNOWN_IMAGE_BLOBS.pop(sku_key.upper(), None)
        blob = bucket.blob(rel_path)
        if blob.exists():
            blob.delete()
            print(f"[GCS] Imagen eliminada de GCS: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error al eliminar imagen de GCS: {e}")
        return False


def download_image_from_gcs(image_path: str) -> bool:
    """Descarga una imagen específica desde GCS si no existe en local."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        p = Path(image_path)
        if not p.is_absolute():
            p = BASE_DIR / image_path
        p.parent.mkdir(parents=True, exist_ok=True)
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
        blob = bucket.blob(rel_path)
        if blob.exists():
            blob.download_to_filename(str(p))
            return True
    except Exception as e:
        print(f"[GCS] Error descargando imagen {image_path}: {e}")
    return False


def list_skus_in_gcs() -> set:
    """Retorna el conjunto de SKUs con imagen usando el índice en memoria o GCS."""
    global _IMAGES_INDEXED
    if _IMAGES_INDEXED and _KNOWN_IMAGE_BLOBS:
        return set(_KNOWN_IMAGE_BLOBS.keys())

    skus = set()
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return skus
    try:
        bucket = _gcs_bucket()
        valid_exts = ('.png', '.jpg', '.jpeg', '.webp')
        for blob in bucket.list_blobs(prefix="imagenes_articulos/"):
            name = Path(blob.name).name
            if name.lower().endswith(valid_exts):
                sku = _extract_sku_from_filename(name)
                if sku:
                    skus.add(sku)
                    _KNOWN_IMAGE_BLOBS[sku] = blob.name
                    _KNOWN_IMAGE_BLOBS[sku.upper()] = blob.name
        _IMAGES_INDEXED = True
    except Exception as e:
        print(f"[GCS] Error listando SKUs de GCS: {e}")
    return skus


def find_and_download_sku_image(sku: str) -> str:
    """
    Busca una imagen para el SKU usando el índice en memoria (0ms si no tiene imagen).
    Solo descarga de GCS si el SKU realmente tiene imagen en el bucket y aún no está en disco.
    """
    if not sku:
        return None
    sku_clean = str(sku).strip()

    # Si ya indexamos el bucket y el SKU NO está en el índice, retornar None al instante (sin llamada de red)
    if _IMAGES_INDEXED:
        blob_name = _KNOWN_IMAGE_BLOBS.get(sku_clean) or _KNOWN_IMAGE_BLOBS.get(sku_clean.upper())
        if not blob_name:
            return None
        local_dest = BASE_DIR / blob_name
        if local_dest.exists():
            return str(local_dest)
        try:
            local_dest.parent.mkdir(parents=True, exist_ok=True)
            bucket = _gcs_bucket()
            bucket.blob(blob_name).download_to_filename(str(local_dest))
            return str(local_dest)
        except Exception:
            return None

    # Si aún no se ha indexado, indexar una sola vez
    list_skus_in_gcs()
    blob_name = _KNOWN_IMAGE_BLOBS.get(sku_clean) or _KNOWN_IMAGE_BLOBS.get(sku_clean.upper())
    if not blob_name:
        return None
    local_dest = BASE_DIR / blob_name
    if local_dest.exists():
        return str(local_dest)
    try:
        local_dest.parent.mkdir(parents=True, exist_ok=True)
        bucket = _gcs_bucket()
        bucket.blob(blob_name).download_to_filename(str(local_dest))
        return str(local_dest)
    except Exception as e:
        print(f"[GCS] Error descargando imagen indexada para {sku_clean}: {e}")
    return None


def read_fresh_excel_from_gcs(file_name: str):
    """Descarga y retorna un DataFrame directamente desde GCS de forma inmediata."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return None
    try:
        import io
        import pandas as pd
        bucket = _gcs_bucket()
        blob = bucket.blob(file_name)
        if blob.exists():
            content = blob.download_as_bytes()
            local_path = BASE_DIR / file_name
            with open(local_path, "wb") as f:
                f.write(content)
            try:
                return pd.read_excel(io.BytesIO(content), sheet_name='Datos_Sistema')
            except Exception:
                return pd.read_excel(io.BytesIO(content), sheet_name=0)
    except Exception as e:
        print(f"[GCS] Error al leer {file_name} desde GCS: {e}")
    return None
