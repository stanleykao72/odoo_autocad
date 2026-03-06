;;; main.lsp — Entry point: loads all modules and registers user commands
;;; Usage: (load "main.lsp") then OB:MENU to open the integration menu
;;;
;;; AutoCAD LT 2024+ compatible (no vlax-create-object / vlax-get-or-create-object)

(vl-load-com)

;;; ============================================================
;;; Module loading
;;; ============================================================

;; Determine our base directory from this file
(defun ob:get-lisp-dir ( / fp)
  (setq fp (findfile "main.lsp"))
  (if fp
    (vl-filename-directory fp)
    ""
  )
)

(defun ob:vlx-mode-p ()
  "Returns T if running inside a VLX (compiled) package."
  ;; In VLX mode, main.lsp won't be found as an external file
  ;; because it's embedded in the VLX. Check if our modules are
  ;; already loaded by testing for a function defined in the last module.
  (and (not (findfile "main.lsp"))
       (eval '(and menu:show T)))
)

(defun ob:load-module (filename / filepath)
  "Loads a .lsp file from the lisp/ directory. Skips if VLX mode."
  (if (ob:vlx-mode-p)
    (progn
      (princ (strcat "\n[OB] VLX mode: " filename " (embedded)"))
      T
    )
    (progn
      (setq filepath (findfile (strcat (ob:get-lisp-dir) "/" filename)))
      (if (not filepath)
        (setq filepath (findfile filename))
      )
      (if filepath
        (progn
          (load filepath)
          T
        )
        (progn
          (princ (strcat "\n[OB] Warning: Could not find " filename))
          nil
        )
      )
    )
  )
)

;; Load modules in dependency order
(princ "\n[OB] ========================================")
(princ "\n[OB] Odoo-AutoCAD Integration (AutoLISP)")
(princ "\n[OB] Loading modules...")
(princ "\n[OB] ========================================")

(ob:load-module "json_util.lsp")
(ob:load-module "file_util.lsp")
(ob:load-module "config.lsp")
(ob:load-module "strip_mtext.lsp")
(ob:load-module "block_util.lsp")
(ob:load-module "table_util.lsp")
(ob:load-module "odoo_bridge.lsp")
(ob:load-module "param_form.lsp")
(ob:load-module "main_menu.lsp")

;; Initialize configuration
(config:init)
(config:load-ini)

;;; ============================================================
;;; User commands
;;; ============================================================

(defun c:OB:MENU ()
  "Opens the Odoo-AutoCAD Integration main menu."
  (menu:show)
  (princ)
)

(defun c:OB:CONNECT ()
  "Tests the Odoo connection."
  (menu:do-connect)
  (princ)
)

(defun c:OB:CONFIG ()
  "Opens the connection settings dialog."
  (menu:do-config)
  (princ)
)

(defun c:OB:SET-PARAMS ()
  "Opens the parameter selection form (loads data from Odoo)."
  (menu:do-set-params)
  (princ)
)

(defun c:OB:PUSH-BOQ ()
  "Collects TABLE data and pushes to Odoo BOQ."
  (menu:do-push-boq)
  (princ)
)

(defun c:OB:CREATE-PR ()
  "Transfers BOQ to Purchase Requisition."
  (menu:do-create-pr)
  (princ)
)

(defun c:OB:CLEAR-IDS ()
  "Clears TABLE IDs in current layout."
  (menu:do-clear-current)
  (princ)
)

(defun c:OB:CLEAR-ALL-IDS ()
  "Clears TABLE IDs in ALL layouts."
  (menu:do-clear-all)
  (princ)
)

;;; ============================================================
;;; Startup message
;;; ============================================================

(princ "\n[OB] ========================================")
(princ "\n[OB] All modules loaded successfully!")
(princ "\n[OB] Commands:")
(princ "\n[OB]   OB:MENU        - Main menu")
(princ "\n[OB]   OB:CONNECT     - Test Odoo connection")
(princ "\n[OB]   OB:CONFIG      - Connection settings")
(princ "\n[OB]   OB:SET-PARAMS  - Set parameters from Odoo")
(princ "\n[OB]   OB:PUSH-BOQ    - Push to BOQ")
(princ "\n[OB]   OB:CREATE-PR   - Transfer BOQ to PR")
(princ "\n[OB]   OB:CLEAR-IDS   - Clear current layout IDs")
(princ "\n[OB]   OB:CLEAR-ALL-IDS - Clear ALL layout IDs")
(princ "\n[OB] ========================================")
(princ)
