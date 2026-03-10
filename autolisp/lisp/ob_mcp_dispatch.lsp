;;; ob_mcp_dispatch.lsp — Odoo IPC extension for autocad-mcp dispatcher
;;; Loads mcp_dispatch.lsp from the autocad-mcp submodule, then extends
;;; the dispatch table with Odoo-specific actions.
;;;
;;; AutoCAD LT 2024+ compatible

;;; ============================================================
;;; Load autocad-mcp base dispatcher
;;; ============================================================

(defun ob:find-mcp-dispatch ( / base-path candidate)
  "Find mcp_dispatch.lsp in the autocad-mcp submodule"
  ;; Try relative to autolisp root: ../libs/autocad-mcp/lisp-code/
  (setq base-path (strcat *ob:root* "/../libs/autocad-mcp/lisp-code/mcp_dispatch.lsp"))
  (setq candidate (findfile base-path))
  (if candidate
    candidate
    ;; Fallback: search on ACAD support path
    (findfile "mcp_dispatch.lsp")
  )
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

(defun ob:action-extract-tables (json-text / result)
  "Collect TABLE + Block data from all layouts"
  (setq result (table:get-all-layouts-data))
  (if result
    (cons T (dmc:json:list_to_json result))
    (cons T "{\"layouts\": []}")
  )
)

(defun ob:action-get-header-ids (json-text / result)
  "Collect all header_ids from all layouts"
  (setq result (table:get-all-header-ids))
  (if result
    (cons T (dmc:json:list_to_json result))
    (cons T "{\"header_ids\": []}")
  )
)

(defun ob:action-write-ids (json-text / response)
  "Write header_id + detail_id back to TABLEs"
  ;; json-text contains the full response data to write back
  (setq response (dmc:json:json_to_list json-text nil))
  (if response
    (progn
      (table:update-ids-from-response response)
      (cons T "{\"message\": \"IDs written successfully\"}")
    )
    (cons nil "Invalid response data for ID writeback")
  )
)

(defun ob:action-get-block-attrs (json-text / block attrs)
  "Read attribute block values from current layout"
  (setq block (block:find-attribute-block))
  (if block
    (progn
      (setq attrs (block:get-all-attributes block))
      (if attrs
        (cons T (dmc:json:list_to_json attrs))
        (cons T "{}")
      )
    )
    (cons nil "No attribute block found in current layout")
  )
)

(defun ob:action-set-block-attrs (json-text / attr-data layout-name)
  "Write attributes to Block in all layouts (or specified layout)"
  (setq attr-data (dmc:json:json_to_list json-text nil))
  (if attr-data
    (progn
      ;; Check if a specific layout_name is specified in params
      (setq layout-name (cdr (assoc "layout_name" attr-data)))
      (if (and layout-name (/= layout-name "null") (/= layout-name ""))
        ;; Write to specific layout
        (let ((block (block:find-attribute-block-in-layout layout-name)))
          (if block
            (progn
              (block:set-attributes block attr-data)
              (cons T (strcat "{\"message\": \"Attributes set in layout: " layout-name "\"}"))
            )
            (cons nil (strcat "No attribute block found in layout: " layout-name))
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

(defun ob:mcp-dispatch-command (cmd-name json-text)
  "Extended dispatcher: try Odoo actions first, then fall back to original"
  (cond
    ;; Odoo-specific actions
    ((= cmd-name "odoo_extract_tables")  (ob:action-extract-tables json-text))
    ((= cmd-name "odoo_get_header_ids")  (ob:action-get-header-ids json-text))
    ((= cmd-name "odoo_write_ids")       (ob:action-write-ids json-text))
    ((= cmd-name "odoo_get_block_attrs") (ob:action-get-block-attrs json-text))
    ((= cmd-name "odoo_set_block_attrs") (ob:action-set-block-attrs json-text))
    ((= cmd-name "odoo_clear_ids")       (ob:action-clear-ids json-text))
    ;; Fall through to original autocad-mcp dispatcher
    (ob:original-mcp-dispatch-command
      (apply ob:original-mcp-dispatch-command (list cmd-name json-text))
    )
    ;; No original dispatcher available
    (T
      (cons nil (strcat "Unknown command: " cmd-name))
    )
  )
)

;; Override the global dispatcher if mcp_dispatch.lsp was loaded
(if ob:mcp-dispatch-path
  (progn
    (setq mcp-dispatch-command ob:mcp-dispatch-command)
    (princ "\n[OB] Dispatcher extended with 6 Odoo actions")
  )
)

(princ "\n[OB] ob_mcp_dispatch.lsp loaded")
(princ)
