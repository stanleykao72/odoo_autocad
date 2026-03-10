;;; strip_mtext.lsp — MText formatting removal
;;; LT-compatible version with fallback
;;;
;;; Primary: LM:UnFormat (uses VBScript.RegExp — Full AutoCAD only)
;;; Fallback: strip:unformat-simple (pure string substitution — LT compatible)

;;; ============================================================
;;; LT Fallback: Pure string substitution chain
;;; ============================================================

(defun strip:unformat-simple (str / patterns pair old new)
  "Removes common MText formatting codes using simple string substitution.
   Works on AutoCAD LT (no VBScript.RegExp needed)."
  (if (not str) (setq str ""))
  (if (not (= (type str) 'STR)) (setq str ""))

  ;; Remove \\P (paragraph/newline) → space
  (setq str (dmc:json:str_replace str "\\P" " "))
  ;; Remove \\p (paragraph formatting)
  ;; Remove \\pxi... patterns - find \p and remove to next ;
  (while (vl-string-search "\\p" str)
    (setq pos (vl-string-search "\\p" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\f (font) codes: \fArial|...; → remove to next ;
  (while (vl-string-search "\\f" str)
    (setq pos (vl-string-search "\\f" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\F (font) codes
  (while (vl-string-search "\\F" str)
    (setq pos (vl-string-search "\\F" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\C (color) codes
  (while (vl-string-search "\\C" str)
    (setq pos (vl-string-search "\\C" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\c (color) codes
  (while (vl-string-search "\\c" str)
    (setq pos (vl-string-search "\\c" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\H (height) codes
  (while (vl-string-search "\\H" str)
    (setq pos (vl-string-search "\\H" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\W (width) codes
  (while (vl-string-search "\\W" str)
    (setq pos (vl-string-search "\\W" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\T (tracking) codes
  (while (vl-string-search "\\T" str)
    (setq pos (vl-string-search "\\T" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\Q (obliquing) codes
  (while (vl-string-search "\\Q" str)
    (setq pos (vl-string-search "\\Q" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\A (alignment) codes
  (while (vl-string-search "\\A" str)
    (setq pos (vl-string-search "\\A" str))
    (setq end (vl-string-search ";" str pos))
    (if end
      (setq str (strcat (substr str 1 pos) (substr str (+ end 2))))
      (setq str (substr str 1 pos))
    )
  )

  ;; Remove \\L (underline on) and \\l (underline off)
  (setq str (dmc:json:str_replace str "\\L" ""))
  (setq str (dmc:json:str_replace str "\\l" ""))

  ;; Remove \\O (overline on) and \\o (overline off)
  (setq str (dmc:json:str_replace str "\\O" ""))
  (setq str (dmc:json:str_replace str "\\o" ""))

  ;; Remove braces
  (setq str (dmc:json:str_replace str "{" ""))
  (setq str (dmc:json:str_replace str "}" ""))

  ;; Remove \\~ (non-breaking space) → regular space
  (setq str (dmc:json:str_replace str "\\~" " "))

  ;; Clean up multiple spaces
  (while (vl-string-search "  " str)
    (setq str (dmc:json:str_replace str "  " " "))
  )

  ;; Trim
  (vl-string-trim " " str)
)

;;; ============================================================
;;; LM:UnFormat (Full AutoCAD — uses VBScript.RegExp)
;;; ============================================================

(defun LM:UnFormat ( str mtx / _replace rx )
  "Returns a string with all MText formatting codes removed.
   Author: Lee Mac, Copyright c 2011 - www.lee-mac.com"

  (defun _replace ( new old str )
    (vlax-put-property rx 'pattern old)
    (vlax-invoke rx 'replace str new)
  )

  (if (setq rx (vl-catch-all-apply 'vlax-create-object '("VBScript.RegExp")))
    (progn
      (if (vl-catch-all-error-p rx)
        ;; VBScript.RegExp not available (LT) — use fallback
        (strip:unformat-simple str)
        (progn
          (setq str
            (vl-catch-all-apply
              (function
                (lambda ( )
                  (vlax-put-property rx 'global     :vlax-true)
                  (vlax-put-property rx 'multiline  :vlax-true)
                  (vlax-put-property rx 'ignorecase :vlax-false)
                  (foreach pair
                    '(
                      ("\032"    . "\\\\\\\\")
                      (" "       . "\\\\P|\\n|\\t")
                      ("$1"      . "\\\\(\\\\[ACcFfHLlOopQTW])|\\\\[ACcFfHLlOopQTW][^\\\\;]*;|\\\\[ACcFfHLlOopQTW]")
                      ("$1$2/$3" . "([^\\\\])\\\\S([^;]*)[/#\\^]([^;]*);")
                      ("$1$2"    . "\\\\(\\\\S)|[\\\\](})|}")
                      ("$1"      . "[\\\\]({)|{")
                    )
                    (setq str (_replace (car pair) (cdr pair) str))
                  )
                  (if mtx
                    (_replace "\\\\" "\032" (_replace "\\$1$2$3" "(\\\\[ACcFfHLlOoPpQSTW])|({)|(})" str))
                    (_replace "\\"   "\032" str)
                  )
                )
              )
            )
          )
          (vlax-release-object rx)
          (if (null (vl-catch-all-error-p str))
            str
            (strip:unformat-simple str) ;; Fallback on error
          )
        )
      )
    )
    ;; vlax-create-object failed entirely — use fallback
    (strip:unformat-simple str)
  )
)

;;; ============================================================
;;; Unified entry point
;;; ============================================================

(defun strip:unformat (str)
  "Removes MText formatting. Uses RegExp on full AutoCAD, fallback on LT."
  (if (not str) (setq str ""))
  (if (not (= (type str) 'STR)) (setq str (vl-princ-to-string str)))
  (if (= str "") str
    (LM:UnFormat str nil)
  )
)

(princ "\n[OB] strip_mtext.lsp loaded")
(princ)
