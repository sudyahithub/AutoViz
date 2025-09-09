;; Portable loader for LineAuditTool.dll
;; 1) If AutoCAD knows this LISP's full path, load DLL from the same folder.
;; 2) If not, ask you to pick the DLL once via a file dialog.

(defun _dirname (path / pos i ch)
  (setq pos 0 i 1)
  (while (<= i (strlen path))
    (setq ch (substr path i 1))
    (if (or (= ch "\\") (= ch "/")) (setq pos i))
    (setq i (1+ i))
  )
  (if (> pos 0) (substr path 1 pos) "")
)

(defun C:LOAD_LINEAUDITTOOL ( / lspfull dll)
  ;; Try to use the path of the file currently being loaded.
  (if (and (boundp '*LOAD-TRUENAME*) *LOAD-TRUENAME*)
    (setq lspfull *LOAD-TRUENAME*)
    (setq lspfull nil)
  )

  (if lspfull
    (progn
      (setq dll (strcat (_dirname lspfull) "LineAuditTool.dll"))
      (if (findfile dll)
        (progn
          (command "._NETLOAD" dll)
          (princ "\n[LineAuditTool] Loaded. Commands: BATCH_ARC2LIN, POLY2TRI_VALIDATE, POLY2TRI_CLEAR_ERRORS.")
        )
        (princ (strcat "\n[LineAuditTool] DLL not found next to LISP:\n  " dll))
      )
    )
    (progn
      ;; Fallback: ask the user once
      (princ "\n[LineAuditTool] Can't auto-detect LISP folder. Please select LineAuditTool.dll")
      (setq dll (getfiled "Select LineAuditTool.dll" "" "dll" 0))
      (if dll
        (progn
          (command "._NETLOAD" dll)
          (princ "\n[LineAuditTool] Loaded. Commands: BATCH_ARC2LIN, POLY2TRI_VALIDATE, POLY2TRI_CLEAR_ERRORS.")
        )
        (princ "\n[LineAuditTool] Cancelled.")
      )
    )
  )
  (princ)
)

;; Auto-run on load/drag
(C:LOAD_LINEAUDITTOOL)
(princ)
