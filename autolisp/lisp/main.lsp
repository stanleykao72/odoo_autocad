;;; main.lsp — Entry point: loads all modules and registers user commands
;;; Usage: (load "main.lsp") then OB:MCP-DISPATCH to execute IPC/MCP commands
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
    ;; Fallback: hardcoded project path
    "C:/odoo/autocad_source/autolisp/lisp"
  )
)

(defun ob:load-module (filename / filepath)
  "Loads a .lsp file from the lisp/ directory."
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
(ob:load-module "ob_mcp_dispatch.lsp")

;; Initialize configuration
(config:init)

;;; ============================================================
;;; User commands
;;; ============================================================

(defun c:OB:MCP-DISPATCH ()
  "Executes IPC/MCP dispatch — reads command JSON, executes, writes result JSON."
  (c:mcp-dispatch)
  (princ)
)

;;; ============================================================
;;; Startup message
;;; ============================================================

(princ "\n[OB] ========================================")
(princ "\n[OB] All modules loaded successfully!")
(princ "\n[OB] Commands:")
(princ "\n[OB]   OB:MCP-DISPATCH - IPC/MCP dispatcher")
(princ "\n[OB] ========================================")
(princ)
