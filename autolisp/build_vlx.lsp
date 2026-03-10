;;; build_vlx.lsp — VLX Build Script for Odoo-AutoCAD Integration
;;;
;;; Usage (in Full AutoCAD command line):
;;;   (load "C:/odoo/autocad_source/autolisp/build_vlx.lsp")
;;;   BUILD-ALL          ; build both VLX files
;;;   BUILD-MCP-DISPATCH ; build McpDispatch.vlx only
;;;   BUILD-ODOO-VLX     ; build OdooAutoCAD.vlx only
;;;
;;; Output:
;;;   autolisp/dist/McpDispatch.vlx
;;;   autolisp/dist/OdooAutoCAD.vlx
;;;
;;; Requires: Full AutoCAD 2024+ (VLIDE compiler)

;;; ============================================================
;;; Configuration — modify these paths if your layout differs
;;; ============================================================

(setq *build:project-root*  "C:/odoo/autocad_source")
(setq *build:lisp-dir*      (strcat *build:project-root* "/autolisp/lisp/"))
(setq *build:mcp-lisp-dir*  (strcat *build:project-root* "/libs/autocad-mcp/lisp-code/"))
(setq *build:dist-dir*      (strcat *build:project-root* "/autolisp/dist/"))

;;; ============================================================
;;; Helper: compile single .lsp -> .fas
;;; ============================================================

(defun build:compile-one (src-path / fas-path)
  "Compile a .lsp file to .fas, returns .fas path"
  (setq fas-path (vl-string-subst ".fas" ".lsp" src-path))
  (princ (strcat "\n  Compiling: " (vl-filename-base src-path) ".lsp"))
  (vlisp-compile 'st src-path fas-path)
  fas-path
)

;;; ============================================================
;;; 1. Build McpDispatch.vlx
;;; ============================================================

(defun c:BUILD-MCP-DISPATCH ( / fas1 fas2 vlx-path)
  (princ "\n[BUILD] ========================================")
  (princ "\n[BUILD] Building McpDispatch.vlx ...")
  (princ "\n[BUILD] ========================================")

  (setq fas1 (build:compile-one (strcat *build:mcp-lisp-dir* "attribute_tools.lsp")))
  (setq fas2 (build:compile-one (strcat *build:mcp-lisp-dir* "mcp_dispatch.lsp")))

  (setq vlx-path (strcat *build:dist-dir* "McpDispatch.vlx"))

  (command "_.VLISP-MAKE-APP"
    "McpDispatch"
    vlx-path
    fas1
    fas2
    ""
  )

  (princ (strcat "\n[BUILD] McpDispatch.vlx -> " vlx-path))
  (princ)
)

;;; ============================================================
;;; 2. Build OdooAutoCAD.vlx
;;; ============================================================

(defun c:BUILD-ODOO-VLX ( / fas-list vlx-path)
  (princ "\n[BUILD] ========================================")
  (princ "\n[BUILD] Building OdooAutoCAD.vlx ...")
  (princ "\n[BUILD] ========================================")

  ;; Compile in dependency order
  (setq fas-list (list
    (build:compile-one (strcat *build:lisp-dir* "010_json_util.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "020_file_util.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "030_config.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "040_strip_mtext.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "050_block_util.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "060_table_util.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "070_ob_mcp_dispatch.lsp"))
    (build:compile-one (strcat *build:lisp-dir* "080_main.lsp"))
  ))

  (setq vlx-path (strcat *build:dist-dir* "OdooAutoCAD.vlx"))

  (command "_.VLISP-MAKE-APP"
    "OdooAutoCAD"
    vlx-path
    (nth 0 fas-list) (nth 1 fas-list)
    (nth 2 fas-list) (nth 3 fas-list)
    (nth 4 fas-list) (nth 5 fas-list)
    (nth 6 fas-list) (nth 7 fas-list)
    ""
  )

  (princ (strcat "\n[BUILD] OdooAutoCAD.vlx -> " vlx-path))
  (princ)
)

;;; ============================================================
;;; 3. Build All
;;; ============================================================

(defun c:BUILD-ALL ()
  (c:BUILD-MCP-DISPATCH)
  (c:BUILD-ODOO-VLX)
  (princ "\n")
  (princ "\n[BUILD] ========================================")
  (princ "\n[BUILD] All VLX files built successfully!")
  (princ "\n[BUILD]   dist/McpDispatch.vlx  (IPC core)")
  (princ "\n[BUILD]   dist/OdooAutoCAD.vlx  (Odoo modules)")
  (princ "\n[BUILD] ========================================")
  (princ)
)

;;; ============================================================
(princ "\n[BUILD] build_vlx.lsp loaded")
(princ "\n[BUILD] Commands: BUILD-ALL, BUILD-MCP-DISPATCH, BUILD-ODOO-VLX")
(princ)
