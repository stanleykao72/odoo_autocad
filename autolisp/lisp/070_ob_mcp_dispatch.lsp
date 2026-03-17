;;; 070_ob_mcp_dispatch.lsp — Odoo IPC extension for autocad-mcp dispatcher
;;; Loads mcp_dispatch.lsp from the autocad-mcp submodule, then extends
;;; the dispatch table with Odoo-specific actions.
;;;
;;; AutoCAD LT 2024+ compatible

;;; ============================================================
;;; Load autocad-mcp base dispatcher
;;; ============================================================

(defun ob:find-mcp-dispatch ( / base-path candidate lisp-dir)
  "Find mcp_dispatch.lsp in the autocad-mcp submodule"
  ;; Try relative to autolisp root: ../libs/autocad-mcp/lisp-code/
  (if *ob:root*
    (progn
      (setq base-path (strcat *ob:root* "/../libs/autocad-mcp/lisp-code/mcp_dispatch.lsp"))
      (setq candidate (findfile base-path))
    )
  )
  ;; Fallback: derive from our own lisp directory
  (if (not candidate)
    (progn
      (setq lisp-dir (vl-filename-directory (findfile "070_ob_mcp_dispatch.lsp")))
      (if lisp-dir
        (progn
          (setq base-path (strcat lisp-dir "/../../libs/autocad-mcp/lisp-code/mcp_dispatch.lsp"))
          (setq candidate (findfile base-path))
        )
      )
    )
  )
  ;; Fallback: hardcoded project path
  (if (not candidate)
    (setq candidate (findfile "C:/odoo/autocad_source/libs/autocad-mcp/lisp-code/mcp_dispatch.lsp"))
  )
  ;; Last fallback: search on ACAD support path
  (if (not candidate)
    (setq candidate (findfile "mcp_dispatch.lsp"))
  )
  candidate
)

;; Check if mcp-dispatch-command is already loaded (e.g. from McpDispatch.vlx)
(if (and (not (eval '(and mcp-dispatch-command T)))
         (not (eval '(and c:mcp-dispatch T))))
  (progn
    ;; Not loaded yet — try to find and load from file
    (setq ob:mcp-dispatch-path (ob:find-mcp-dispatch))
    (if ob:mcp-dispatch-path
      (progn
        (load ob:mcp-dispatch-path)
        (princ "\n[OB] mcp_dispatch.lsp loaded from file")
      )
      (princ "\n[OB] Warning: mcp_dispatch.lsp not found — IPC/MCP mode unavailable")
    )
  )
  (princ "\n[OB] mcp_dispatch already loaded (VLX or prior load)")
)

;;; ============================================================
;;; Save reference to original dispatcher
;;; ============================================================

;; Save the original mcp-dispatch-command before we wrap it
(if (eval '(and mcp-dispatch-command T))
  (progn
    (setq ob:original-mcp-dispatch-command mcp-dispatch-command)
    (princ "\n[OB] Saved reference to original mcp-dispatch-command")
  )
)

;;; ============================================================
;;; Odoo action implementations
;;; ============================================================

(defun ob:action-get-layouts (json-text / doc layouts layout-name result)
  "Return list of non-Model layout names as JSON array"
  (vl-load-com)
  (setq result '())
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-layouts doc))
      (vlax-for layout layouts
        (setq layout-name (vla-get-name layout))
        (if (/= layout-name "Model")
          (setq result (cons layout-name result))
        )
      )
    )
  )
  (setq result (reverse result))
  ;; Build JSON array string manually
  (if result
    (progn
      (setq json-str "[")
      (foreach name result
        (if (> (strlen json-str) 1)
          (setq json-str (strcat json-str ","))
        )
        (setq json-str (strcat json-str "\"" name "\""))
      )
      (setq json-str (strcat json-str "]"))
      (cons T (strcat "{\"layouts\":" json-str "}"))
    )
    (cons T "{\"layouts\": []}")
  )
)

(defun ob:action-extract-tables (json-text / result)
  "Collect TABLE + Block data from all layouts"
  (setq result (table:get-all-layouts-data))
  (if result
    (cons T (dmc:json:list_to_json result))
    (cons T "{\"layouts\": []}")
  )
)

(defun ob:action-extract-table-for-layout (json-text / full-data layout-name result)
  "Collect TABLE + Block data from a single layout by name"
  (setq full-data (dmc:json:json_to_list json-text nil))
  (setq layout-name (cdr (assoc "layout_name"
                     (if (assoc "params" full-data)
                       (cadr (assoc "params" full-data))
                       full-data))))
  (if (and layout-name (/= layout-name "") (/= layout-name "null"))
    (progn
      (setq result (table:get-single-layout-data layout-name))
      (if result
        (cons T (dmc:json:list_to_json result))
        (cons T "{}")
      )
    )
    (cons nil "Missing or empty layout_name parameter")
  )
)

(defun ob:action-get-header-ids (json-text / result json-str)
  "Collect all header_ids from all layouts"
  (setq result (table:get-all-header-ids))
  (if (eval '(and ob:log T))
    (ob:log (strcat "[HEADER-IDS] Found " (itoa (length result)) " IDs"))
  )
  (if result
    (progn
      ;; Build JSON array manually — list_to_json can't handle flat int list
      (setq json-str "[")
      (foreach id result
        (if (> (strlen json-str) 1)
          (setq json-str (strcat json-str ","))
        )
        (setq json-str (strcat json-str (itoa id)))
      )
      (setq json-str (strcat json-str "]"))
      (cons T (strcat "{\"header_ids\":" json-str "}"))
    )
    (cons T "{\"header_ids\": []}")
  )
)

(defun ob:action-write-ids (json-text / full-data response all-list layout-count)
  "Write header_id + detail_id back to TABLEs"
  ;; json-text contains the full IPC JSON — extract params sub-object
  (setq full-data (dmc:json:json_to_list json-text nil))
  (if (eval '(and ob:log T))
    (ob:log (strcat "[WRITE-IDS] full-data keys: "
      (if full-data
        (vl-princ-to-string (mapcar 'car (vl-remove-if-not 'listp full-data)))
        "(nil)")))
  )
  (setq response (cadr (assoc "params" full-data)))
  (if (not response)
    (setq response full-data)  ;; fallback
  )
  (if (eval '(and ob:log T))
    (ob:log (strcat "[WRITE-IDS] response type: " (vl-princ-to-string (type response))
      ", keys: " (if (and response (listp response))
        (vl-princ-to-string (mapcar 'car (vl-remove-if-not 'listp response)))
        "(not alist)")))
  )
  ;; Check for "all" key
  (setq all-list (cdr (assoc "all" response)))
  (if (eval '(and ob:log T))
    (ob:log (strcat "[WRITE-IDS] all-list: "
      (if all-list
        (strcat "found, length=" (itoa (length all-list))
          ", first=" (vl-princ-to-string (car all-list)))
        "(nil - not found)")))
  )
  (if response
    (progn
      (table:update-ids-from-response response)
      (cons T "{\"message\": \"IDs written successfully\"}")
    )
    (cons nil "Invalid response data for ID writeback")
  )
)

(defun ob:action-get-block-attrs (json-text / block attrs layout-name pr-block pr-text)
  "Read attribute block values + pr_no block text from current layout"
  (setq layout-name
    (if (vla-get-activedocument (vlax-get-acad-object))
      (vla-get-name (vla-get-activelayout (vla-get-activedocument (vlax-get-acad-object))))
      "(no doc)"
    )
  )
  (if (eval '(and ob:log T))
    (ob:log (strcat "[BLOCK] Searching for attribute block in layout: " layout-name))
  )
  (setq block (block:find-attribute-block))
  (if block
    (progn
      (setq attrs (block:get-all-attributes block))
      ;; Add layout_name to response
      (setq attrs (cons (cons "layout_name" layout-name) attrs))
      ;; Also look for the separate "pr_no" block (contains text, not attributes)
      (setq pr-block (block:find-block-by-name "pr_no"))
      (if pr-block
        (progn
          (setq pr-text (block:get-block-text pr-block))
          (if pr-text
            (progn
              (if (eval '(and ob:log T))
                (ob:log (strcat "[BLOCK] Found pr_no block text: " pr-text))
              )
              ;; Prepend pr_no to attrs list
              (setq attrs (cons (cons "pr_no" pr-text) attrs))
            )
          )
        )
      )
      (if (eval '(and ob:log T))
        (ob:log (strcat "[BLOCK] Found block with " (itoa (length attrs)) " attributes"))
      )
      (if attrs
        (cons T (dmc:json:list_to_json attrs))
        (cons T "{}")
      )
    )
    (progn
      (if (eval '(and ob:log T))
        (ob:log (strcat "[BLOCK] No attribute block (pr_no/project_name/job_working_plan_name) in layout: " layout-name))
      )
      (cons nil (strcat "No attribute block found in layout: " layout-name))
    )
  )
)

(defun ob:action-set-block-attrs (json-text / full-data attr-data layout-name layout-obj doc)
  "Write attributes to Block in specified layout (or all layouts)"
  (setq full-data (dmc:json:json_to_list json-text nil))
  ;; IPC wraps user data inside "params" key — extract it
  ;; json_parser stores nested objects as ("key" (alist...)) so use cadr
  (setq attr-data (cadr (assoc "params" full-data)))
  (if (not attr-data)
    (setq attr-data full-data)  ;; fallback: use top-level if no params wrapper
  )
  (if attr-data
    (progn
      ;; Check if a specific layout_name is specified in params
      (setq layout-name (cdr (assoc "layout_name" attr-data)))
      (if (eval '(and ob:log T))
        (ob:log (strcat "[SET] layout_name param: " (if layout-name layout-name "(nil)")))
      )
      (if (and layout-name (/= layout-name "null") (/= layout-name ""))
        ;; Write to specific layout — look up VLA layout object by name
        (progn
          (setq doc (vla-get-activedocument (vlax-get-acad-object)))
          (setq layout-obj
            (vl-catch-all-apply 'vla-item
              (list (vla-get-layouts doc) layout-name)))
          (if (vl-catch-all-error-p layout-obj)
            (cons nil (strcat "Layout not found: " layout-name))
            (progn
              (setq ob:tmp-block (block:find-attribute-block-in-layout layout-obj))
              (if ob:tmp-block
                (progn
                  (block:set-attributes ob:tmp-block attr-data)
                  (if (eval '(and ob:log T))
                    (ob:log (strcat "[SET] Attributes set in layout: " layout-name))
                  )
                  (cons T (strcat "{\"message\": \"Attributes set in layout: " layout-name "\"}"))
                )
                (cons nil (strcat "No attribute block found in layout: " layout-name))
              )
            )
          )
        )
        ;; Write to all layouts
        (progn
          (block:set-attributes-all-layouts attr-data)
          (cons T "{\"message\": \"Attributes set in all layouts\"}")
        )
      )
    )
    (cons nil "Invalid attribute data")
  )
)

(defun ob:action-clear-ids (json-text /)
  "Clear all TABLE IDs in all layouts"
  (table:clear-all-layouts-ids)
  (cons T "{\"message\": \"All TABLE IDs cleared\"}")
)

;;; ============================================================
;;; Extended dispatcher — wraps original with Odoo actions
;;; ============================================================

(defun ob:mcp-dispatch-command (cmd-name json-text / result err-obj)
  "Extended dispatcher: try Odoo actions first, then fall back to original"
  (if (eval '(and ob:log T))
    (ob:log (strcat "[DISPATCH] " cmd-name))
  )
  (setq result
    (cond
      ;; Odoo-specific actions
      ((= cmd-name "odoo_get_layouts")      (ob:action-get-layouts json-text))
      ((= cmd-name "odoo_extract_tables")  (ob:action-extract-tables json-text))
      ((= cmd-name "odoo_extract_table_for_layout") (ob:action-extract-table-for-layout json-text))
      ((= cmd-name "odoo_get_header_ids")  (ob:action-get-header-ids json-text))
      ((= cmd-name "odoo_write_ids")       (ob:action-write-ids json-text))
      ((= cmd-name "odoo_get_block_attrs") (ob:action-get-block-attrs json-text))
      ((= cmd-name "odoo_set_block_attrs") (ob:action-set-block-attrs json-text))
      ((= cmd-name "odoo_clear_ids")       (ob:action-clear-ids json-text))
      ;; Fall through to original autocad-mcp dispatcher
      ;; Note: cannot use apply/vl-catch-all-apply on VLX SUBR — use lambda wrapper
      (ob:original-mcp-dispatch-command
        (progn
          (setq err-obj
            (vl-catch-all-apply
              '(lambda (c j) (ob:original-mcp-dispatch-command c j))
              (list cmd-name json-text)
            )
          )
          (if (vl-catch-all-error-p err-obj)
            (progn
              (if (eval '(and ob:log T))
                (ob:log (strcat "[DISPATCH] ERROR in original dispatcher: " (vl-catch-all-error-message err-obj)))
              )
              (cons nil (strcat "Internal error: " (vl-catch-all-error-message err-obj)))
            )
            err-obj
          )
        )
      )
      ;; No original dispatcher available
      (T
        (cons nil (strcat "Unknown command: " cmd-name))
      )
    )
  )
  ;; Log result status
  (if (eval '(and ob:log T))
    (if result
      (ob:log (strcat "[DISPATCH] " cmd-name " -> " (if (car result) "OK" (strcat "FAIL: " (if (cdr result) (cdr result) "")))))
      (ob:log (strcat "[DISPATCH] " cmd-name " -> NIL (no result)"))
    )
  )
  result
)

;; Override the global dispatcher — always, regardless of how mcp_dispatch.lsp was loaded
(if (eval '(and mcp-dispatch-command T))
  (progn
    (setq mcp-dispatch-command ob:mcp-dispatch-command)
    (princ "\n[OB] Dispatcher extended with 8 Odoo actions")
    (if (eval '(and ob:log T))
      (ob:log "[OB] Dispatcher override: mcp-dispatch-command -> ob:mcp-dispatch-command (8 Odoo actions)")
    )
  )
  (progn
    (princ "\n[OB] Warning: mcp-dispatch-command not found, cannot extend dispatcher")
    (if (eval '(and ob:log T))
      (ob:log "[OB] ERROR: mcp-dispatch-command not found — Odoo IPC commands unavailable")
    )
  )
)

(princ "\n[OB] ob_mcp_dispatch.lsp loaded")
(princ)
