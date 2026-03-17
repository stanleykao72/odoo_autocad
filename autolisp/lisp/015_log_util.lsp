;;; 015_log_util.lsp - File logging for AutoLISP
;;; Writes to logs/ directory alongside Python logs
;;; No COM dependencies - LT compatible

;;; ============================================================
;;; Log file setup
;;; ============================================================

(setq *ob:log-version* "6.0.0.13")

(defun log:get-log-dir ( / root)
  "Returns the logs/ directory path (creates if needed)"
  (setq root
    (cond
      (*ob:root* (strcat (vl-string-right-trim "\\/" *ob:root*) "/../logs/"))
      (T "C:/odoo/autocad_source/logs/")
    )
  )
  ;; Normalize slashes
  (while (vl-string-search "\\" root)
    (setq root (vl-string-subst "/" "\\" root))
  )
  (vl-mkdir root)
  root
)

(defun log:get-timestamp ( / d yr mo dy hr mi se)
  "Returns current time as HH:MM:SS string"
  (setq d (getvar "CDATE"))
  ;; CDATE = YYYYMMDD.HHMMSS as real number
  (setq yr (fix (/ d 10000.0)))
  (setq d (- d (* yr 10000.0)))
  (setq mo (fix (/ d 100.0)))
  (setq dy (fix (rem d 100.0)))
  (setq d (rem (getvar "CDATE") 1.0))
  (setq d (* d 1000000.0))
  (setq hr (fix (/ d 10000.0)))
  (setq d (- d (* hr 10000.0)))
  (setq mi (fix (/ d 100.0)))
  (setq se (fix (rem d 100.0)))
  (strcat
    (if (< hr 10) (strcat "0" (itoa hr)) (itoa hr))
    ":"
    (if (< mi 10) (strcat "0" (itoa mi)) (itoa mi))
    ":"
    (if (< se 10) (strcat "0" (itoa se)) (itoa se))
  )
)

(defun log:get-date-str ( / d yr mo dy)
  "Returns current date as YYYY-MM-DD string"
  (setq d (fix (getvar "CDATE")))
  (setq yr (fix (/ d 10000)))
  (setq d (- d (* yr 10000)))
  (setq mo (fix (/ d 100)))
  (setq dy (rem d 100))
  (strcat
    (itoa yr) "-"
    (if (< mo 10) (strcat "0" (itoa mo)) (itoa mo)) "-"
    (if (< dy 10) (strcat "0" (itoa dy)) (itoa dy))
  )
)

(defun log:get-log-file ()
  "Returns today's log file path"
  (strcat (log:get-log-dir) "autolisp_" (log:get-date-str) ".log")
)

;;; ============================================================
;;; Logging functions
;;; ============================================================

(defun ob:log (msg / f logfile)
  "Write a log message to file and command line.
   Usage: (ob:log \"[OB] Something happened\")"
  (setq logfile (log:get-log-file))
  (if (setq f (open logfile "a"))
    (progn
      (write-line (strcat (log:get-timestamp) " " msg) f)
      (close f)
    )
  )
  ;; Also print to command line
  (princ (strcat "\n" msg))
)

(defun ob:log-session-start ()
  "Write session start marker to log"
  (ob:log (strcat "=== AutoLISP session started v" *ob:log-version* " ==="))
  (ob:log (strcat "[OB] Log file: " (log:get-log-file)))
)

(princ "\n[OB] log_util.lsp loaded")
(princ)
