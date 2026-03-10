;;; 030_config.lsp — Path constants and YAML config (pure AutoLISP)
;;; No COM dependencies — LT compatible

;;; ============================================================
;;; Global path configuration
;;; ============================================================

;; Resolve root directory.
;; Dev mode: autolisp/lisp/config.lsp → autolisp/ (source tree)
(defun config:init ( / lsp-dir)
  ;; Try to find config relative to known search paths
  (cond
    ;; If *ob:root* already set, keep it
    (*ob:root* nil)

    ;; Dev mode: derive from 030_config.lsp → lisp/ → autolisp/
    ((setq lsp-dir (findfile "030_config.lsp"))
     (setq *ob:root*
       (vl-string-right-trim
         "\\/"
         (vl-filename-directory (vl-filename-directory lsp-dir))
       )
     )
    )

    ;; Dev mode fallback: derive from 080_main.lsp
    ((setq lsp-dir (findfile "080_main.lsp"))
     (setq *ob:root*
       (vl-string-right-trim
         "\\/"
         (vl-filename-directory (vl-filename-directory lsp-dir))
       )
     )
    )

    ;; Last resort
    (T
      (setq *ob:root* (getvar "DWGPREFIX"))
      (princ "\n[OB] Warning: Could not find root, using DWGPREFIX")
    )
  )

  ;; Derived paths
  (setq *ob:lisp-dir*   (strcat *ob:root* "/lisp/"))
  (setq *ob:config-dir* (strcat *ob:root* "/config/"))

  ;; Odoo YAML config files (production — same as main Python app)
  (setq *ob:server-yaml*
    (cond
      ((findfile (strcat *ob:config-dir* "server_prod.yaml"))
       (findfile (strcat *ob:config-dir* "server_prod.yaml")))
      ((findfile "C:/odoo/config/server_prod.yaml")
       (findfile "C:/odoo/config/server_prod.yaml"))
      ((findfile (strcat *ob:root* "/server_prod.yaml"))
       (findfile (strcat *ob:root* "/server_prod.yaml")))
      (T nil)
    )
  )
  (setq *ob:token-yaml*
    (cond
      ((findfile (strcat *ob:config-dir* "token.yaml"))
       (findfile (strcat *ob:config-dir* "token.yaml")))
      ((findfile "C:/odoo/config/token.yaml")
       (findfile "C:/odoo/config/token.yaml"))
      ((findfile (strcat *ob:root* "/token.yaml"))
       (findfile (strcat *ob:root* "/token.yaml")))
      (T nil)
    )
  )

  ;; Temp directory for JSON exchange (IPC mode)
  (setq *ob:temp-dir*
    (cond
      ((getenv "TEMP") (strcat (getenv "TEMP") "/odoo_bridge/"))
      ((getenv "TMP")  (strcat (getenv "TMP")  "/odoo_bridge/"))
      (T "C:/Temp/odoo_bridge/")
    )
  )

  ;; Connection state
  (if (not *ob:odoo-connected*)
    (setq *ob:odoo-connected* nil)
  )

  (princ (strcat "\n[OB] Root: " *ob:root*))
  (if *ob:server-yaml*
    (princ (strcat "\n[OB] Server config: " *ob:server-yaml*))
    (princ "\n[OB] Server config: (not found)")
  )
  (if *ob:token-yaml*
    (princ (strcat "\n[OB] Token config: " *ob:token-yaml*))
    (princ "\n[OB] Token config: (not found)")
  )
  (princ)
)

;;; ============================================================
;;; INI file reading (pure AutoLISP — kept for YAML-like parsing)
;;; ============================================================

(defun ini:read-file (ini-path / f line section key val result current-section)
  "Reads an INI file, returns association list of ((section.key . value) ...)"
  (setq result '())
  (setq current-section "")
  (if (and ini-path (setq f (open ini-path "r")))
    (progn
      (while (setq line (read-line f))
        ;; Trim whitespace
        (setq line (vl-string-trim " \t" line))
        (cond
          ;; Empty or comment line
          ((or (= line "") (= (substr line 1 1) ";") (= (substr line 1 1) "#"))
           nil
          )
          ;; Section header [section]
          ((and (= (substr line 1 1) "[")
                (vl-string-search "]" line))
           (setq current-section
             (substr line 2 (- (vl-string-search "]" line) 1))
           )
          )
          ;; Key = value
          ((vl-string-search "=" line)
           (setq key (vl-string-trim " \t"
                       (substr line 1 (vl-string-search "=" line))))
           (setq val (vl-string-trim " \t"
                       (substr line (+ (vl-string-search "=" line) 2))))
           ;; Remove inline comments
           (if (vl-string-search ";" val)
             (setq val (vl-string-trim " \t"
                         (substr val 1 (vl-string-search ";" val))))
           )
           (setq result
             (cons (cons (strcat current-section "." key) val) result)
           )
          )
        )
      )
      (close f)
    )
  )
  (reverse result)
)

(defun ini:get (ini-data section key / lookup)
  "Gets a value from parsed INI data"
  (setq lookup (strcat section "." key))
  (cdr (assoc lookup ini-data))
)

(princ "\n[OB] config.lsp loaded")
(princ)
