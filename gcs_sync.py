"""
=============================================================================
SIGRAMA REMISIONES DE MATERIALES - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia Bidireccional con Google Cloud Storage (Cloud Run)
=============================================================================
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

def _get_bucket_name() -> str:
    raw = os.environ.get("GCS_BUCKET", "sigrama-remisiones-storage")
    if not raw or not str(raw).strip():
        return "sigrama-remisiones-storage"
    # Sanitizar en caso de que variables concatenadas se cuelen en la cadena
    return str(raw).split()[0].strip()

def _gcs_bucket():
    from google.cloud import storage
    bucket_name = _get_bucket_name()
    # Soporte para credenciales en st.secrets si se ejecuta en Streamlit Cloud
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
    """Descarga bases de datos Excel e imágenes desde GCS al arrancar el contenedor."""
    bucket_name = _get_bucket_name()
    if not bucket_name:
        return False
    try:
        bucket = _gcs_bucket()
        count = 0
        for blob in bucket.list_blobs():
            if blob.name.endswith("/"):
                continue
            local_path = BASE_DIR / blob.name
            local_path.parent.mkdir(parents=True, exist_ok=True)
            if not local_path.exists() or local_path.stat().st_size != blob.size:
                blob.download_to_filename(str(local_path))
                count += 1
        print(f"[GCS] Sincronizados exitosamente {count} archivos desde gs://{bucket_name}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{bucket_name}: {e}")
        return False

def push_excel_to_gcs(file_name: str):
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

def push_image_to_gcs(image_path: str):
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

def delete_image_from_gcs(image_path: str):
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
