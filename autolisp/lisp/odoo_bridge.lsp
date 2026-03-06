;;; odoo_bridge.lsp — Bridge communication layer
;;; Writes request JSON → startapp bridge exe → polls for response → reads result
;;; LT compatible (no COM, uses startapp + file polling)

;;; ============================================================
;;; Core bridge call
;;; ============================================================

(defun bridge:call (action payload / req-file resp-file req-json result)
  "Calls the Python bridge with an action and payload.
   payload: association list to be JSON-serialized into params.
   Returns: parsed response association list."

  ;; Ensure temp dir exists
  (file:ensure-dir *ob:temp-dir*)

  ;; Generate unique filenames
  (setq req-file  (file:unique-name *ob:temp-dir* "req" "json"))
  (setq resp-file (file:unique-name *ob:temp-dir* "resp" "json"))

  ;; Build request JSON
  (setq req-json
    (dmc:json:list_to_json
      (list
        (cons "action" action)
        (cons "params" (if payload (list payload) '()))
      )
    )
  )

  ;; Write request file
  (if (not (file:write-string req-file req-json))
    (progn
      (princ (strcat "\n[OB] Error: Could not write request file: " req-file))
      (setq result (list (cons "success" 'false)
                         (cons "error_code" "FILE_WRITE_ERROR")
                         (cons "message" "Could not write request JSON")))
    )
    (progn
      ;; Launch bridge
      (bridge:launch action req-file resp-file)

      ;; Wait for response
      (setq result (bridge:wait-for-response resp-file *ob:bridge-timeout*))
    )
  )

  ;; Cleanup temp files
  (file:delete req-file)
  (file:delete resp-file)

  result
)

(defun bridge:build-config-args (/ args)
  "Builds config CLI arguments for bridge. Prefers YAML over INI."
  (setq args "")
  (if (and *ob:server-yaml* *ob:token-yaml*)
    ;; YAML mode (production)
    (setq args (strcat " --server-config \"" *ob:server-yaml* "\""
                       " --token-config \"" *ob:token-yaml* "\""))
    ;; INI fallback
    (if *ob:ini-file*
      (setq args (strcat " --config \"" *ob:ini-file* "\""))
    )
  )
  args
)

(defun bridge:launch (action req-file resp-file / cmd cfg-args)
  "Launches the bridge executable or python script."
  (setq cfg-args (bridge:build-config-args))
  (if *ob:bridge-exe*
    ;; Production: use compiled exe
    (startapp *ob:bridge-exe*
      (strcat action " \"" req-file "\" \"" resp-file "\"" cfg-args))
    ;; Dev mode: use python
    (progn
      (setq cmd (strcat "python \""
                         *ob:bridge-dir* "odoo_bridge.py\" "
                         action " \"" req-file "\" \"" resp-file "\""
                         cfg-args))
      (startapp "cmd" (strcat "/c " cmd))
    )
  )
)

(defun bridge:wait-for-response (resp-file timeout / elapsed json-str)
  "Polls for response file. Returns parsed result or timeout error.
   NOTE: This runs OUTSIDE any DCL dialog — (command) is safe here."
  (setq elapsed 0)
  (princ "\n[OB] Waiting for bridge response")
  (while (and (not (findfile resp-file)) (< elapsed timeout))
    (command "_.delay" 500)
    (setq elapsed (+ elapsed 500))
    (princ ".")  ;; progress dots on command line
  )
  (princ (if (findfile resp-file) " OK" " TIMEOUT"))
  (if (findfile resp-file)
    (progn
      (setq json-str (file:read-string resp-file))
      (if (and json-str (/= json-str ""))
        (dmc:json:json_to_list json-str nil)
        (list (cons "success" 'false)
              (cons "error_code" "EMPTY_RESPONSE")
              (cons "message" "Bridge returned empty response"))
      )
    )
    (list (cons "success" 'false)
          (cons "error_code" "TIMEOUT")
          (cons "message" (strcat "Bridge response timeout (" (itoa (/ timeout 1000)) "s)")))
  )
)

;;; ============================================================
;;; Response helpers
;;; ============================================================

(defun bridge:success-p (response)
  "Returns T if response indicates success."
  (cond
    ((assoc "success" response)
     (eq (cdr (assoc "success" response)) 'true))
    ;; If no success key but has data, assume success
    ((assoc "data" response) T)
    (T nil)
  )
)

(defun bridge:get-data (response)
  "Gets the data field from response."
  (cdr (assoc "data" response))
)

(defun bridge:get-message (response)
  "Gets the message field from response."
  (cdr (assoc "message" response))
)

(defun bridge:get-error (response)
  "Gets the error_code field from response."
  (cdr (assoc "error_code" response))
)

;;; ============================================================
;;; API wrappers
;;; ============================================================

(defun bridge:test-connection ()
  "Tests Odoo connection. Returns response list."
  (bridge:call "test_connection" nil)
)

(defun bridge:get-project (pr-no)
  "Gets project info by PR number."
  (bridge:call "get_project"
    (list (cons "pr_no" pr-no))
  )
)

(defun bridge:get-products ()
  "Gets product list from Odoo."
  (bridge:call "get_products" nil)
)

(defun bridge:get-setup ()
  "Gets setup values (spec, catalog, etc.) from Odoo."
  (bridge:call "get_setup" nil)
)

(defun bridge:get-colors (project-id)
  "Gets color list from Odoo for a project."
  (bridge:call "get_colors"
    (list (cons "project_id" project-id))
  )
)

(defun bridge:import-to-boq (data)
  "Pushes BOQ data to Odoo. data is the full layout+detail JSON list."
  (bridge:call "import_to_boq" data)
)

(defun bridge:boq-to-pr (header-ids)
  "Converts BOQ to Purchase Requisition. header-ids is a list of IDs."
  (bridge:call "boq_to_pr"
    (list (cons "header_ids" header-ids))
  )
)

(princ "\n[OB] odoo_bridge.lsp loaded")
(princ)
