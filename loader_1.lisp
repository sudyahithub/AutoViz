;; Drag & drop this .lsp into AutoCAD.
;; It will NETLOAD the plugin DLL from a fixed path.

(defun C:LOAD_LINEAUDITTOOL ( / dll)
  ;; 🔴 Change this path if you move the DLL:
  (setq dll "C:/Users/admin/Downloads/VIZ-AUTOCAD/IDENTIFIER/LineAuditTool.dll")

  (if (findfile dll)
    (progn
      (command "._NETLOAD" dll)
      (princ "\n[LineAuditTool] Loaded. Commands: BATCH_ARC2LIN, POLY2TRI_VALIDATE, POLY2TRI_CLEAR_ERRORS.")
    )
    (princ (strcat "\n[LineAuditTool] DLL not found at: " dll))
  )
  (princ)
)

;; Auto-run when loaded
(C:LOAD_LINEAUDITTOOL)
(princ)
