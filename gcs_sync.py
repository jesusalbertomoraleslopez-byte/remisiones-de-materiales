"""
=============================================================================
SIGRAMA REMISIONES DE MATERIALES - INDUSTRIA SIGRAMA S.A. DE C.V.
Módulo de Persistencia Bidireccional con Google Cloud Storage (Cloud Run)
=============================================================================
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
GCS_BUCKET = os.environ.get("GCS_BUCKET", "sigrama-remisiones-storage").strip()

def _gcs_bucket():
    from google.cloud import storage
    return storage.Client().bucket(GCS_BUCKET)

def sync_from_gcs():
    """Descarga bases de datos Excel e imágenes desde GCS al arrancar el contenedor."""
    if not GCS_BUCKET:
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
        print(f"[GCS] Sincronizados exitosamente {count} archivos desde gs://{GCS_BUCKET}")
        return True
    except Exception as e:
        print(f"[GCS] Error al sincronizar desde gs://{GCS_BUCKET}: {e}")
        return False

def push_excel_to_gcs(file_name: str):
    """Sube un archivo Excel específico a GCS."""
    if not GCS_BUCKET:
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
    if not GCS_BUCKET:
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
