;;; config.lsp — Path constants and INI read/write (pure AutoLISP)
;;; No COM dependencies — LT compatible

;;; ============================================================
;;; Global path configuration
;;; ============================================================

;; Resolve root directory.
;; VLX mode: OdooBridge.vlx location → root (exe + config/ alongside)
;; Dev mode: autolisp/lisp/config.lsp → autolisp/ (source tree)
(defun config:init ( / vlx-path lsp-dir)
  ;; Try to find config relative to known search paths
  (cond
    ;; If *ob:root* already set, keep it
    (*ob:root* nil)

    ;; VLX mode: find OdooBridge.vlx on search path
    ((setq vlx-path (findfile "OdooBridge.vlx"))
     (setq *ob:root*
       (vl-string-right-trim "\\/" (vl-filename-directory vlx-path))
     )
     (princ "\n[OB] VLX mode detected")
    )

    ;; Dev mode: derive from config.lsp → lisp/ → autolisp/
    ((setq lsp-dir (findfile "config.lsp"))
     (setq *ob:root*
       (vl-string-right-trim
         "\\/"
         (vl-filename-directory (vl-filename-directory lsp-dir))
       )
     )
    )

    ;; Dev mode fallback: derive from main.lsp
    ((setq lsp-dir (findfile "main.lsp"))
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

  ;; Derived paths — dev mode uses subdirectory structure
  (setq *ob:lisp-dir*   (strcat *ob:root* "/lisp/"))
  (setq *ob:dcl-dir*    (strcat *ob:root* "/dcl/"))
  (setq *ob:bridge-dir* (strcat *ob:root* "/bridge/"))
  (setq *ob:config-dir* (strcat *ob:root* "/config/"))
  (setq *ob:dist-dir*   (strcat *ob:root* "/dist/"))

  ;; Bridge executable — search multiple locations:
  ;; 1. Same dir as VLX (deployment)
  ;; 2. dist/ subdir (dev)
  ;; 3. bridge/ subdir (dev)
  ;; 4. findfile on search path (APPLOAD path)
  (setq *ob:bridge-exe*
    (cond
      ((findfile (strcat *ob:root* "/odoo_bridge.exe"))
       (findfile (strcat *ob:root* "/odoo_bridge.exe")))
      ((findfile (strcat *ob:dist-dir* "odoo_bridge.exe"))
       (findfile (strcat *ob:dist-dir* "odoo_bridge.exe")))
      ((findfile (strcat *ob:bridge-dir* "odoo_bridge.exe"))
       (findfile (strcat *ob:bridge-dir* "odoo_bridge.exe")))
      ((findfile "odoo_bridge.exe")
       (findfile "odoo_bridge.exe"))
      ;; Dev mode: use python directly
      (T nil)
    )
  )

  ;; Odoo YAML config files (production — same as main Python app)
  ;; Search: config/ alongside VLX, then ../../config/ (dev), then C:/odoo/config/
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

  ;; INI config file (legacy fallback)
  (setq *ob:ini-file*
    (cond
      ((findfile (strcat *ob:config-dir* "bridge.ini"))
       (findfile (strcat *ob:config-dir* "bridge.ini")))
      ((findfile (strcat *ob:root* "/bridge.ini"))
       (findfile (strcat *ob:root* "/bridge.ini")))
      (T (strcat *ob:config-dir* "bridge.ini"))
    )
  )

  ;; Temp directory for JSON exchange
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

  ;; Bridge timeout (ms)
  (if (not *ob:bridge-timeout*)
    (setq *ob:bridge-timeout* 30000)
  )

  (princ (strcat "\n[OB] Root: " *ob:root*))
  (princ (strcat "\n[OB] Bridge: " (if *ob:bridge-exe* *ob:bridge-exe* "(not found - dev mode)")))
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
;;; INI file reading (pure AutoLISP — open/read-line)
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

(defun ini:write-file (ini-path data / f prev-section section key val)
  "Writes association list back to INI file"
  (if (setq f (open ini-path "w"))
    (progn
      (setq prev-section "")
      (foreach pair data
        ;; Split "section.key" into section and key
        (setq section (substr (car pair) 1
                        (vl-string-search "." (car pair))))
        (setq key (substr (car pair)
                    (+ (vl-string-search "." (car pair)) 2)))
        (setq val (cdr pair))
        ;; Write section header if changed
        (if (/= section prev-section)
          (progn
            (if (/= prev-section "") (write-line "" f))
            (write-line (strcat "[" section "]") f)
            (setq prev-section section)
          )
        )
        (write-line (strcat key " = " val) f)
      )
      (close f)
      T
    )
  )
)

;;; ============================================================
;;; Load bridge.ini settings
;;; ============================================================

(defun config:load-ini ( / ini-data)
  (setq ini-data (ini:read-file *ob:ini-file*))
  (if ini-data
    (progn
      (setq *ob:odoo-url*    (ini:get ini-data "odoo" "url"))
      (setq *ob:odoo-db*     (ini:get ini-data "odoo" "db"))
      (setq *ob:odoo-user*   (ini:get ini-data "credentials" "username"))
      (setq *ob:odoo-pass*   (ini:get ini-data "credentials" "password"))
      (setq *ob:ini-data*    ini-data)
      T
    )
    (progn
      (princ "\n[OB] Warning: bridge.ini not found or empty")
      nil
    )
  )
)

(defun config:save-ini ()
  "Saves current settings back to bridge.ini"
  (setq *ob:ini-data*
    (list
      (cons "odoo.url"                (if *ob:odoo-url*  *ob:odoo-url*  ""))
      (cons "odoo.db"                 (if *ob:odoo-db*   *ob:odoo-db*   ""))
      (cons "odoo.auth_method"        "basic")
      (cons "credentials.username"    (if *ob:odoo-user* *ob:odoo-user* ""))
      (cons "credentials.password"    (if *ob:odoo-pass* *ob:odoo-pass* ""))
      (cons "paths.temp_dir"          *ob:temp-dir*)
      (cons "paths.log_file"          (strcat *ob:temp-dir* "bridge.log"))
      (cons "options.timeout"         "30")
      (cons "options.debug"           "false")
      (cons "options.verify_ssl"      "true")
    )
  )
  (ini:write-file *ob:ini-file* *ob:ini-data*)
)

(princ "\n[OB] config.lsp loaded")
(princ)
