"""
=============================================================================
SIGRAMA REMISIONES DE MATERIALES - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia y Almacenamiento 100% Google Cloud Storage (GCP)
Optimizado con Descarga Concurrente Multihilo de Alto Rendimiento
=============================================================================
"""

import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

BASE_DIR = Path(__file__).resolve().parent

def _get_bucket_name() -> str:
    raw = os.environ.get("GCS_BUCKET", "sigrama-remisiones-storage")
    if not raw or not str(raw).strip():
        return "sigrama-remisiones-storage"
    return str(raw).split()[0].strip()

def _gcs_bucket():
    from google.cloud import storage
    bucket_name = _get_bucket_name()
    # Soporte para credenciales explícitas en st.secrets si se ejecuta en Streamlit Cloud
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
            from google.oauth2 import service_account
            creds = service_account.Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]))
            return storage.Client(credentials=creds).bucket(bucket_name)
    except Exception:
        pass
    return storage.Client().bucket(bucket_name)

def sync_from_gcs():
    """Descarga bases de datos Excel e imágenes desde GCS de forma ultrarrápida y concurrente."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        bucket = _gcs_bucket()
        blobs_to_download = []
        for blob in bucket.list_blobs():
            if blob.name.endswith("/"):
                continue
            local_path = BASE_DIR / blob.name
            if not local_path.exists() or local_path.stat().st_size != blob.size:
                blobs_to_download.append((blob, local_path))

        if not blobs_to_download:
            print(f"[GCS] Todos los archivos ya están al día con gs://{bucket_name}")
            return True

        def _download(item):
            b, lp = item
            lp.parent.mkdir(parents=True, exist_ok=True)
            b.download_to_filename(str(lp))

        with ThreadPoolExecutor(max_workers=16) as executor:
            list(executor.map(_download, blobs_to_download))

        print(f"[GCS] Sincronizados exitosamente {len(blobs_to_download)} archivos concurrentemente desde gs://{bucket_name}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{bucket_name}: {e}")
        return False

def push_excel_to_gcs(file_name: str) -> bool:
    """Sube un archivo Excel específico a GCS."""
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
    """Sube una imagen de SKU a GCS."""
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
        print(f"[GCS] Imagen subida a GCS: {rel_path}")
        return True
    except Exception as e:
        print(f"[GCS] Error subiendo imagen a GCS: {e}")
        return False

def delete_image_from_gcs(image_path: str) -> bool:
    """Elimina una imagen de SKU en GCS."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        p = Path(image_path)
        if not p.is_absolute():
            p = BASE_DIR / image_path
        bucket = _gcs_bucket()
        rel_path = p.relative_to(BASE_DIR).as_posix()
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
    """Lista el conjunto de SKUs con imagen en el bucket de GCS."""
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
                sku = name.split("(")[0].strip() if "(" in name else Path(name).stem.strip()
                if sku:
                    skus.add(sku)
    except Exception as e:
        print(f"[GCS] Error listando SKUs de GCS: {e}")
    return skus

def find_and_download_sku_image(sku: str) -> str:
    """Busca una imagen para el SKU en GCS y la descarga a local si la encuentra."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return None
    try:
        bucket = _gcs_bucket()
        valid_exts = ('.png', '.jpg', '.jpeg', '.webp')
        prefix = f"imagenes_articulos/{sku}"
        for blob in bucket.list_blobs(prefix=prefix):
            filename = Path(blob.name).name
            blob_sku = filename.split("(")[0].strip() if "(" in filename else Path(filename).stem.strip()
            if blob_sku == sku and filename.lower().endswith(valid_exts):
                local_dest = BASE_DIR / blob.name
                local_dest.parent.mkdir(parents=True, exist_ok=True)
                blob.download_to_filename(str(local_dest))
                print(f"[GCS] Imagen descargada para {sku}: {local_dest}")
                return str(local_dest)
    except Exception as e:
        print(f"[GCS] Error buscando imagen para {sku} en GCS: {e}")
    return None

def read_fresh_excel_from_gcs(file_name: str):
    """Descarga y retorna un DataFrame directamente desde GCS de forma inmediata."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return None
    try:
        import io, pandas as pd
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
