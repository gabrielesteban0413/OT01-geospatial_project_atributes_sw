-- ============================================================
-- Auditoría de registros faltantes
-- Compara nce_backup_4_8_26 (backup) contra nce_exporte_ltp (actual)
-- por la columna 'ne' y guarda los faltantes en audit_nce_backup
-- con una sola columna final de error.
-- Excluye los ne que contengan 5328 o 5324 (equipos QUIDWAY).
-- ============================================================

DROP TABLE IF EXISTS public.audit_nce_backup;

CREATE TABLE public.audit_nce_backup AS
SELECT
    b.*,
    'No se encuentra en el archivo actual' AS error
FROM public.nce_backup_4_8_26 b
LEFT JOIN public.nce_exporte_ltp a
    ON a.ne = b.ne
WHERE a.ne IS NULL
  AND b.ne NOT LIKE '%5328%'
  AND b.ne NOT LIKE '%5324%'
  AND b.ne
  ;
  

CREATE INDEX idx_audit_nce_backup_ne
    ON public.audit_nce_backup (ne);

SELECT COUNT(*) AS filas_faltantes
FROM public.audit_nce_backup;