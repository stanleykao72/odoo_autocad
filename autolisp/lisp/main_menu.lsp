;;; main_menu.lsp — Main menu while-loop + action dispatch
;;; Implements the DCL close → work → reopen pattern to solve
;;; the DCL blocking limitation (no polling during start_dialog).

;;; ============================================================
;;; Result display helper
;;; ============================================================

(defun menu:show-result (title msg / dch lines i)
  "Shows a result dialog with title and multi-line message."
  (setq dch (dcl:load "result.dcl"))
  (if dch
    (progn
      (if (new_dialog "ob_result" dch)
        (progn
          (set_tile "result_title" title)
          ;; Split message into lines (by \n)
          (setq lines (param:split-string msg "\n"))
          (setq i 1)
          (foreach line lines
            (if (<= i 5)
              (set_tile (strcat "result_line" (itoa i)) line)
            )
            (setq i (1+ i))
          )
          (action_tile "accept" "(done_dialog 1)")
          (start_dialog)
        )
      )
      (unload_dialog dch)
    )
    ;; Fallback to alert if DCL not found
    (alert (strcat title "\n\n" msg))
  )
)

;;; ============================================================
;;; String split helper
;;; ============================================================

(defun param:split-string (str delim / pos result)
  "Splits a string by delimiter. Returns list of substrings."
  (setq result '())
  (while (setq pos (vl-string-search delim str))
    (setq result (cons (substr str 1 pos) result))
    (setq str (substr str (+ pos (strlen delim) 1)))
  )
  (setq result (cons str result))
  (reverse result)
)

;;; ============================================================
;;; Action handlers
;;; ============================================================

(defun menu:do-connect (/ response)
  "Tests Odoo connection via bridge."
  (princ "\n[OB] Testing Odoo connection...")
  (setq response (bridge:test-connection))
  (if (bridge:success-p response)
    (progn
      (setq *ob:odoo-connected* T)
      (menu:show-result "Connection Test"
        (strcat "Success!\n" (if (bridge:get-message response)
                               (bridge:get-message response) "")))
    )
    (progn
      (setq *ob:odoo-connected* nil)
      (menu:show-result "Connection Test"
        (strcat "Failed!\n"
                (if (bridge:get-error response)
                  (strcat "Error: " (bridge:get-error response) "\n") "")
                (if (bridge:get-message response)
                  (bridge:get-message response) "")))
    )
  )
)

(defun menu:do-config (/ dch action)
  "Shows connection settings dialog."
  ;; Load current settings
  (config:load-ini)

  (setq dch (dcl:load "config.dcl"))
  (if dch
    (progn
      (if (new_dialog "ob_config" dch)
        (progn
          ;; Populate current values
          (if *ob:odoo-url*  (set_tile "cfg_url"  *ob:odoo-url*))
          (if *ob:odoo-db*   (set_tile "cfg_db"   *ob:odoo-db*))
          (if *ob:odoo-user* (set_tile "cfg_user" *ob:odoo-user*))
          (if *ob:odoo-pass* (set_tile "cfg_pass" *ob:odoo-pass*))

          (action_tile "accept"
            (strcat
              "(progn "
                "(setq *ob:odoo-url*  (get_tile \"cfg_url\"))"
                "(setq *ob:odoo-db*   (get_tile \"cfg_db\"))"
                "(setq *ob:odoo-user* (get_tile \"cfg_user\"))"
                "(setq *ob:odoo-pass* (get_tile \"cfg_pass\"))"
                "(done_dialog 1)"
              ")")
          )
          (action_tile "cancel" "(done_dialog 0)")

          (setq action (start_dialog))
        )
      )
      (unload_dialog dch)

      ;; Save if OK was pressed
      (if (= action 1)
        (progn
          (config:save-ini)
          (princ "\n[OB] Settings saved to bridge.ini")
          (menu:show-result "Settings" "Configuration saved successfully.")
        )
      )
    )
    (alert "[OB] Error: config.dcl not found")
  )
)

(defun menu:do-set-params (/ lov-list block selections)
  "Loads Odoo data → shows parameter form → writes to block attributes."
  (if (not *ob:odoo-connected*)
    (progn
      (alert "Not connected to Odoo.\nPlease test connection first.")
    )
    (progn
      ;; Load LOV data from Odoo
      (setq lov-list (param:load-lov-data))
      (if lov-list
        (progn
          ;; Find the attribute block in current layout
          (setq block (block:find-attribute-block))
          (if (not block)
            (alert "No attribute block found in current layout.\nPlease make sure you have a block with project_name attribute.")
            (progn
              ;; Show form and get user selection
              (setq selections (param:show-form lov-list))
              (if selections
                (progn
                  (param:apply-to-block block selections)
                  (menu:show-result "Set Parameters" "Parameters applied successfully!")
                )
              )
            )
          )
        )
      )
    )
  )
)

(defun menu:do-push-boq (/ all-data import-json response)
  "Collects TABLE data from all layouts → pushes to Odoo BOQ → writes back IDs."
  (if (not *ob:odoo-connected*)
    (progn
      (alert "Not connected to Odoo.\nPlease test connection first.")
    )
    (progn
      (princ "\n[OB] Collecting TABLE data from all layouts...")
      (setq all-data (table:get-all-layouts-data))
      (if (not all-data)
        (alert "No valid TABLE data found in any layout.")
        (progn
          ;; Build JSON and send to bridge
          (princ "\n[OB] Pushing to Odoo BOQ...")
          (setq response (bridge:import-to-boq all-data))
          (if (bridge:success-p response)
            (progn
              ;; Write back IDs
              (princ "\n[OB] Writing back IDs to tables...")
              (table:update-ids-from-response response)
              (menu:show-result "Push to BOQ" "Successfully imported to BOQ!\nIDs written back to tables.")
            )
            (menu:show-result "Push to BOQ"
              (strcat "Failed!\n" (bridge:get-message response)))
          )
        )
      )
    )
  )
)

(defun menu:do-create-pr (/ header-ids response)
  "Collects header_ids → sends BOQ-to-PR request via bridge."
  (if (not *ob:odoo-connected*)
    (progn
      (alert "Not connected to Odoo.\nPlease test connection first.")
    )
    (progn
      (princ "\n[OB] Collecting header IDs from tables...")
      (setq header-ids (table:get-all-header-ids))
      (if (not header-ids)
        (alert "No header IDs found in tables.\nPlease push to BOQ first.")
        (progn
          (princ (strcat "\n[OB] Found " (itoa (length header-ids)) " headers. Creating PR..."))
          (setq response (bridge:boq-to-pr header-ids))
          (if (bridge:success-p response)
            (menu:show-result "Transfer BOQ to PR"
              (strcat "Successfully transferred to PR!\n"
                      (if (bridge:get-message response)
                        (bridge:get-message response) "")))
            (menu:show-result "Transfer BOQ to PR"
              (strcat "Failed!\n" (bridge:get-message response)))
          )
        )
      )
    )
  )
)

(defun menu:do-clear-current (/ count)
  "Clears TABLE IDs in current layout."
  (setq count (table:clear-current-layout-ids))
  (menu:show-result "Clear IDs"
    (strcat "Cleared IDs from " (itoa count) " table(s) in current layout."))
)

(defun menu:do-clear-all (/ count)
  "Clears TABLE IDs in ALL layouts."
  (setq count (table:clear-all-layouts-ids))
  (menu:show-result "Clear IDs"
    (strcat "Cleared IDs from " (itoa count) " table(s) across all layouts."))
)

;;; ============================================================
;;; Main menu loop
;;; ============================================================

(defun menu:show (/ keep-open dch action)
  "Shows the main menu in a while loop.
   Each button does done_dialog(N) → work outside dialog → re-open."
  (setq keep-open T)

  (while keep-open
    ;; Load and show dialog
    (setq dch (dcl:load "main_menu.dcl"))

    (if (not dch)
      (progn
        (alert "[OB] Error: main_menu.dcl not found!")
        (setq keep-open nil)
      )
      (progn
        (if (not (new_dialog "ob_main_menu" dch))
          (progn
            (unload_dialog dch)
            (alert "[OB] Error: Could not load main menu dialog")
            (setq keep-open nil)
          )
          (progn
            ;; Update status
            (set_tile "status_text"
              (if *ob:odoo-connected*
                "Odoo: Connected"
                "Odoo: Not Connected"
              )
            )

            ;; Button actions → done_dialog with unique codes
            (action_tile "btn_connect"       "(done_dialog 1)")
            (action_tile "btn_config"        "(done_dialog 2)")
            (action_tile "btn_set_params"    "(done_dialog 3)")
            (action_tile "btn_push_boq"      "(done_dialog 4)")
            (action_tile "btn_create_pr"     "(done_dialog 5)")
            (action_tile "btn_clear_current" "(done_dialog 6)")
            (action_tile "btn_clear_all"     "(done_dialog 7)")
            (action_tile "cancel"            "(done_dialog 0)")

            ;; Block until user clicks a button
            (setq action (start_dialog))
            (unload_dialog dch)

            ;; === Dialog closed — dispatch action ===
            (cond
              ((= action 0) (setq keep-open nil))     ;; Close
              ((= action 1) (menu:do-connect))         ;; Test Connection
              ((= action 2) (menu:do-config))          ;; Settings
              ((= action 3) (menu:do-set-params))      ;; Set Parameters
              ((= action 4) (menu:do-push-boq))        ;; Push to BOQ
              ((= action 5) (menu:do-create-pr))       ;; Create PR
              ((= action 6) (menu:do-clear-current))   ;; Clear current
              ((= action 7) (menu:do-clear-all))       ;; Clear all
            )
            ;; while loop re-opens the menu (unless action=0)
          )
        )
      )
    )
  )
  (princ)
)

(princ "\n[OB] main_menu.lsp loaded")
(princ)
