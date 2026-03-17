;;; file_util.lsp — File I/O utilities (pure AutoLISP, LT compatible)
;;; No COM dependencies

;;; ============================================================
;;; File writing
;;; ============================================================

(defun file:write-string (filepath content / f)
  "Writes a string to a file. Returns filepath on success, nil on failure."
  (if (setq f (open filepath "w"))
    (progn
      (write-line content f)
      (close f)
      filepath
    )
  )
)

(defun file:write-lines (filepath lines / f)
  "Writes a list of strings to a file. Returns filepath on success."
  (if (setq f (open filepath "w"))
    (progn
      (foreach line lines
        (write-line line f)
      )
      (close f)
      filepath
    )
  )
)

;;; ============================================================
;;; File reading
;;; ============================================================

(defun file:read-string (filepath / f line result)
  "Reads entire file into a single string. Returns nil if file not found."
  (if (and filepath (setq f (open filepath "r")))
    (progn
      (setq result "")
      (while (setq line (read-line f))
        (if (= result "")
          (setq result line)
          (setq result (strcat result "\n" line))
        )
      )
      (close f)
      result
    )
  )
)

(defun file:read-lines (filepath / f line result)
  "Reads file into a list of strings. Returns nil if file not found."
  (if (and filepath (setq f (open filepath "r")))
    (progn
      (setq result '())
      (while (setq line (read-line f))
        (setq result (cons line result))
      )
      (close f)
      (reverse result)
    )
  )
)

;;; ============================================================
;;; File utilities
;;; ============================================================

(defun file:unique-name (dir prefix ext / name counter)
  "Generates a unique filename in dir: prefix_TIMESTAMP_N.ext"
  (setq counter 0)
  (setq name
    (strcat dir prefix "_"
      (menucmd "M=$(edtime,$(getvar,date),YYYYMODDHHMMSS)")
    )
  )
  ;; Ensure uniqueness
  (while (findfile (strcat name (if (> counter 0) (strcat "_" (itoa counter)) "") "." ext))
    (setq counter (1+ counter))
  )
  (strcat name
    (if (> counter 0) (strcat "_" (itoa counter)) "")
    "." ext
  )
)

(defun file:delete (filepath)
  "Deletes a file. Returns T on success."
  (if (and filepath (findfile filepath))
    (progn
      (vl-file-delete filepath)
      T
    )
  )
)

(defun file:ensure-dir (dir / parts current)
  "Ensures a directory exists by creating it if needed."
  (if (not (vl-file-directory-p dir))
    (vl-mkdir dir)
  )
)

(defun file:exists-p (filepath)
  "Returns T if file exists."
  (if (and filepath (findfile filepath)) T nil)
)

;;; ============================================================
;;; JSON string escaping
;;; ============================================================

(defun json:escape (str / bs dq)
  "Escapes special characters in a string for safe JSON serialization.
   Handles: \" newline carriage-return tab
   Must be called BEFORE wrapping string in double quotes.
   NOTE: Backslash escaping DISABLED — vl-string-search is byte-level,
   not MBCS-aware. Big5/cp950 characters like 蓋 (0xBB5C) contain 0x5C
   as a trail byte, which gets falsely matched as backslash. The Python
   side handles any real backslashes via _fix_json_backslashes()."
  (if (not str) (setq str ""))
  (setq bs (chr 92))   ;; backslash
  (setq dq (chr 34))   ;; double quote
  ;; Backslash escaping removed — MBCS-unsafe (corrupts Big5 蓋功等)
  ;; (setq str (dmc:json:str_replace str bs (strcat bs bs)))
  (setq str (dmc:json:str_replace str dq (strcat bs dq)))           ;; " → \"
  (setq str (dmc:json:str_replace str (chr 10) (strcat bs "n")))    ;; LF → \n
  (setq str (dmc:json:str_replace str (chr 13) (strcat bs "r")))    ;; CR → \r
  (setq str (dmc:json:str_replace str (chr 9)  (strcat bs "t")))    ;; TAB → \t
  str
)

(defun json:unescape (str / bs)
  "Reverses JSON escape sequences in a string.
   Call after extracting string values from parsed JSON."
  (if (not str) (setq str ""))
  (setq bs (chr 92))  ;; backslash
  ;; Order matters: control chars first, backslash last
  (setq str (dmc:json:str_replace str (strcat bs "n")  (chr 10)))   ;; \n → LF
  (setq str (dmc:json:str_replace str (strcat bs "r")  (chr 13)))   ;; \r → CR
  (setq str (dmc:json:str_replace str (strcat bs "t")  (chr 9)))    ;; \t → TAB
  (setq str (dmc:json:str_replace str (strcat bs (chr 34)) (chr 34))) ;; \" → "
  (setq str (dmc:json:str_replace str (strcat bs bs)   bs))          ;; \\ → \
  str
)

;;; ============================================================
;;; DCL loading (VLX / LT dual-mode)
;;; ============================================================

(defun dcl:load (filename / dcl-path dch)
  "Loads a DCL file. VLX mode: embedded resource. LT mode: external file.
   Returns dialog handle (dch) or nil on failure."
  ;; Try VLX embedded resource first (just filename, no path)
  (setq dch (load_dialog filename))
  (if (and dch (> dch 0))
    dch
    (progn
      ;; Fallback: external file with path
      (setq dcl-path (findfile (strcat *ob:dcl-dir* filename)))
      (if (not dcl-path)
        (setq dcl-path (findfile filename))
      )
      (if dcl-path
        (progn
          (setq dch (load_dialog dcl-path))
          (if (and dch (> dch 0)) dch nil)
        )
        nil
      )
    )
  )
)

(princ "\n[OB] file_util.lsp loaded")
(princ)
