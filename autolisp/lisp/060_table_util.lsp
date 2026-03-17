;;; table_util.lsp — TABLE entity operations (LT compatible)
;;; Refactored from legacy/transfer_to_odoo.lsp
;;; Uses DXF entity data (entget/entmod) instead of VLA COM objects
;;;
;;; TABLE structure (9 columns):
;;;   Row 0: title row — col 7 = "HEADER_ID", col 8 = header_id value
;;;   Row 1: column headers
;;;   Row 2+: data rows — col 0=position, 1=product_no, 2=width, 3=height,
;;;           4=length, 5=thickness, 6=qty, 7=desc, 8=detail_id

;;; ============================================================
;;; TABLE cell read/write via DXF entity data
;;; ============================================================

(defun table:get-cell-text (ename row col / obj)
  "Reads text from a TABLE cell using entget.
   Returns the text string or empty string."
  ;; Use (vlax-invoke) if vla available, otherwise ActiveX is needed
  ;; For LT we'll use the ssget + entget approach with TABLE entity
  ;; Note: TABLE cells in DXF are complex. We use vla if available.
  (if (and ename (not (null ename)))
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj
        (progn
          (setq txt (vl-catch-all-apply 'vla-gettext (list obj row col)))
          (if (vl-catch-all-error-p txt)
            ""
            (strip:unformat txt)
          )
        )
        ""
      )
    )
    ""
  )
)

(defun table:set-cell-text (ename row col value / obj)
  "Writes text to a TABLE cell."
  (if (and ename (not (null ename)))
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj
        (vl-catch-all-apply 'vla-settext (list obj row col value))
      )
    )
  )
)

(defun table:get-rows (ename / obj)
  "Gets number of rows in a TABLE."
  (if ename
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj (vla-get-rows obj) 0)
    )
    0
  )
)

(defun table:get-columns (ename / obj)
  "Gets number of columns in a TABLE."
  (if ename
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (if obj (vla-get-columns obj) 0)
    )
    0
  )
)

;;; ============================================================
;;; TABLE identification
;;; ============================================================

(defun table:legal-p (ename)
  "Returns T if ename is a TABLE with HEADER_ID at col 7 row 0 and 9 columns."
  (and ename
       (>= (table:get-columns ename) 9)
       (= (table:get-cell-text ename 0 7) "HEADER_ID"))
)

;;; ============================================================
;;; Layout iteration — collect all tables
;;; ============================================================

(defun table:get-layout-tables (/ ss i ename tables)
  "Gets all TABLE entities in current layout (paperspace).
   Returns list of enames."
  (setq tables '())
  (if (setq ss (ssget "X" '((0 . "ACAD_TABLE"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i))
        (if (table:legal-p ename)
          (setq tables (cons ename tables))
        )
        (setq i (1+ i))
      )
    )
  )
  (reverse tables)
)

;;; ============================================================
;;; Data extraction from TABLE
;;; ============================================================

(defun table:get-detail-rows (ename / rows row detail-lst qty position product-no
                                     width height len thickness desc detail-id)
  "Extracts all data rows from a TABLE. Returns list of assoc lists.
   Skips rows where qty is not a positive integer."
  (setq detail-lst '())
  (setq rows (table:get-rows ename))
  (setq row 2) ;; Skip title (0) and header (1) rows
  (while (< row rows)
    (setq qty (table:get-cell-text ename row 6))
    (setq qty-num (if (distof qty) (atoi qty) nil))
    (if (and qty-num (> qty-num 0))
      (progn
        (setq position   (table:get-cell-text ename row 0))
        (setq product-no (table:get-cell-text ename row 1))
        (setq width      (table:get-cell-text ename row 2))
        (setq height     (table:get-cell-text ename row 3))
        (setq len        (table:get-cell-text ename row 4))
        (setq thickness  (table:get-cell-text ename row 5))
        (setq desc       (table:get-cell-text ename row 7))
        (setq detail-id  (table:get-cell-text ename row 8))

        (setq detail-lst
          (cons
            (list
              (cons "header_id" (table:get-cell-text ename 0 8))
              (cons "detail_id" detail-id)
              (cons "position" position)
              (cons "product_no" product-no)
              (cons "width" width)
              (cons "height" height)
              (cons "len" len)
              (cons "thickness" thickness)
              (cons "qty" qty-num)
              (cons "desc" desc)
            )
            detail-lst
          )
        )
      )
    )
    (setq row (1+ row))
  )
  (reverse detail-lst)
)

(defun table:get-header-id (ename)
  "Gets the header_id value from TABLE cell (0, 8)."
  (table:get-cell-text ename 0 8)
)

;;; ============================================================
;;; Layout-level data collection
;;; ============================================================

(defun table:get-current-layout-data (/ tables all-details header-id)
  "Collects all TABLE data from current layout.
   Returns: ((header_id . \"123\") (details . (list-of-detail-assocs)))"
  (setq tables (table:get-layout-tables))
  (setq all-details '())
  (setq header-id "")
  (foreach tbl tables
    (setq header-id (table:get-header-id tbl))
    (setq all-details (append all-details (table:get-detail-rows tbl)))
  )
  (if all-details
    (list (cons "header_id" header-id)
          (cons "details" all-details))
    nil
  )
)

(defun table:get-single-layout-data (layout-name / doc layout-obj blocks
                                      header-data all-details header-id
                                      ename layout-data)
  "Collects TABLE + Block data from a SINGLE layout by name.
   Returns layout data assoc list, or nil if no data found."
  (vl-load-com)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (if (not doc) (progn (princ "\n[OB] No active document") (exit)))
  (setq layout-obj (table:find-layout-by-name doc layout-name))
  (if (not layout-obj)
    (progn
      (princ (strcat "\n[OB] Layout not found: " layout-name))
      nil
    )
    (progn
      (princ (strcat "\n[OB] Processing layout: " layout-name))
      (setq blocks (vla-get-block layout-obj))

      ;; Collect block attributes for header + pr_no text block
      (setq header-data nil)
      (setq ob:tmp-pr-no nil)
      (vlax-for block blocks
        (if (= (vla-get-objectname block) "AcDbBlockReference")
          (progn
            ;; Attribute block (project_name, job_working_plan_name, etc.)
            (if (not header-data)
              (setq header-data (block:get-header-attrs block layout-name))
            )
            ;; Separate "pr_no" text block (Name = "pr_no")
            (if (and (not ob:tmp-pr-no)
                     (= (strcase (vla-get-name block)) "PR_NO"))
              (progn
                (setq ob:tmp-pr-no (block:get-block-text block))
                (if (and ob:tmp-pr-no (/= ob:tmp-pr-no ""))
                  (princ (strcat "\n[OB] Found pr_no block: " ob:tmp-pr-no))
                )
              )
            )
          )
        )
      )
      ;; Merge pr_no into header-data if found and not already present
      (if (and ob:tmp-pr-no (/= ob:tmp-pr-no "") header-data
               (not (assoc "pr_no" header-data)))
        (setq header-data (cons (cons "pr_no" ob:tmp-pr-no) header-data))
      )

      ;; Collect TABLE data
      (setq all-details '())
      (setq header-id "")
      (vlax-for block blocks
        (if (= (vla-get-objectname block) "AcDbTable")
          (progn
            (setq ename (vlax-vla-object->ename block))
            (if (table:legal-p ename)
              (progn
                (setq header-id (table:get-header-id ename))
                (setq all-details
                  (append all-details (table:get-detail-rows ename)))
              )
            )
          )
        )
      )

      ;; Build layout entry
      ;; Note: use (list "detail" ...) not (cons "detail" ...) so
      ;; the <ARRAY> list is nested — list_to_json needs (car sublist) = <ARRAY>
      (if (and header-data all-details)
        (append
          (list (cons "header_id" header-id))
          header-data
          (list (list "detail"
            (cons (quote <ARRAY>)
              (append all-details (list (quote </ARRAY>)))))))
        nil
      )
    )
  )
)

(defun table:get-all-layouts-data (/ doc layouts layout layout-name
                                     all-data layout-data)
  "Collects TABLE data from ALL layouts (excluding Model).
   Returns list of layout data assocs for bridge:import-to-boq."
  (vl-load-com)
  (setq all-data '())
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-layouts doc))
      (vlax-for layout layouts
        (setq layout-name (vla-get-name layout))
        (if (/= layout-name "Model")
          (progn
            (setq layout-data (table:get-single-layout-data layout-name))
            (if layout-data
              (setq all-data (cons layout-data all-data))
            )
          )
        )
      )
    )
  )
  (reverse all-data)
)

;;; ============================================================
;;; ID writeback
;;; ============================================================

(defun table:id-to-string (val)
  "Converts header_id/detail_id to string. Handles INT, REAL, and STR."
  (cond
    ((= (type val) 'INT) (itoa val))
    ((= (type val) 'REAL) (itoa (fix val)))  ;; JSON parser may return float
    ((= (type val) 'STR) val)
    (T (vl-princ-to-string val))
  )
)

(defun table:write-header-id (ename header-id)
  "Writes header_id to TABLE cell (0, 8)."
  (table:set-cell-text ename 0 8 (table:id-to-string header-id))
)

(defun table:write-detail-id (ename product-no detail-id / rows row cell-product)
  "Writes detail_id to the row matching product_no in column 1.
   detail-id is written to column 8."
  (setq rows (table:get-rows ename))
  (setq row 2)
  (while (< row rows)
    (setq cell-product (table:get-cell-text ename row 1))
    (if (= cell-product product-no)
      (table:set-cell-text ename row 8 (table:id-to-string detail-id))
    )
    (setq row (1+ row))
  )
)

(defun table:update-ids-from-response (response / doc all-list layout-list
                                        header-id layout-name detail-list
                                        layout layout-obj blocks ename)
  "Updates TABLE IDs from bridge response.
   Response format: ((\"all\" . ((layout-data) (layout-data) ...)))"
  (vl-load-com)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (if doc
    (progn
      (setq all-list (cdr (assoc "all" response)))
      (if (not all-list)
        ;; Try direct data structure
        (setq all-list (bridge:get-data response))
      )
      ;; json_to_list wraps JSON array as extra list level: ((obj1 obj2 ...))
      ;; Unwrap if all-list has 1 element that is itself a list of alists
      (if (and all-list
               (= (length all-list) 1)
               (listp (car all-list))
               (listp (caar all-list))
               (assoc "header_id" (caar all-list)))
        (setq all-list (car all-list))
      )
      (princ (strcat "\n[OB] update-ids: all-list has "
        (if all-list (itoa (length all-list)) "0") " items"))
      (if (eval '(and ob:log T))
        (ob:log (strcat "[WRITE-IDS] Iterating " (if all-list (itoa (length all-list)) "0") " layouts"))
      )
      (foreach layout-list all-list
        (setq header-id   (cdr (assoc "header_id" layout-list)))
        (setq layout-name (cdr (assoc "layout_name" layout-list)))
        (setq detail-list (cdr (assoc "detail" layout-list)))
        ;; Unwrap detail array nesting from json_to_list
        (if (and detail-list
                 (= (length detail-list) 1)
                 (listp (car detail-list))
                 (listp (caar detail-list))
                 (assoc "product_no" (caar detail-list)))
          (setq detail-list (car detail-list))
        )
        (if (eval '(and ob:log T))
          (ob:log (strcat "[WRITE-IDS] Layout: " (if layout-name layout-name "(nil)")
            " header_id: " (vl-princ-to-string header-id)
            " detail count: " (if detail-list (itoa (length detail-list)) "0")))
        )

        (if layout-name
          (progn
            (setq layout-obj (table:find-layout-by-name doc layout-name))
            (if layout-obj
              (progn
                (setq blocks (vla-get-block layout-obj))
                (vlax-for block blocks
                  (if (= (vla-get-objectname block) "AcDbTable")
                    (progn
                      (setq ename (vlax-vla-object->ename block))
                      (if (table:legal-p ename)
                        (progn
                          ;; Write header ID
                          (table:write-header-id ename header-id)
                          ;; Write detail IDs
                          (if detail-list
                            (foreach d detail-list
                              (table:write-detail-id ename
                                (cdr (assoc "product_no" d))
                                (cdr (assoc "detail_id" d)))
                            )
                          )
                        )
                      )
                    )
                  )
                )
              )
            )
          )
        )
      )
    )
  )
)

;;; ============================================================
;;; Clear IDs
;;; ============================================================

(defun table:clear-ids (ename / rows row)
  "Clears header_id and all detail_id values from a TABLE."
  (if (table:legal-p ename)
    (progn
      ;; Clear header_id
      (table:set-cell-text ename 0 8 "")
      ;; Clear all detail_ids
      (setq rows (table:get-rows ename))
      (setq row 2)
      (while (< row rows)
        (table:set-cell-text ename row 8 "")
        (setq row (1+ row))
      )
      T
    )
  )
)

(defun table:clear-current-layout-ids (/ tables count)
  "Clears all TABLE IDs in the current layout."
  (setq tables (table:get-layout-tables))
  (setq count 0)
  (foreach tbl tables
    (if (table:clear-ids tbl)
      (setq count (1+ count))
    )
  )
  count
)

(defun table:clear-all-layouts-ids (/ doc layouts layout layout-name blocks ename count)
  "Clears all TABLE IDs in ALL layouts (excluding Model)."
  (vl-load-com)
  (setq count 0)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-layouts doc))
      (vlax-for layout layouts
        (setq layout-name (vla-get-name layout))
        (if (/= layout-name "Model")
          (progn
            (setq blocks (vla-get-block layout))
            (vlax-for block blocks
              (if (= (vla-get-objectname block) "AcDbTable")
                (progn
                  (setq ename (vlax-vla-object->ename block))
                  (if (table:clear-ids ename)
                    (setq count (1+ count))
                  )
                )
              )
            )
          )
        )
      )
    )
  )
  count
)

;;; ============================================================
;;; Collect header_ids for PR transfer
;;; ============================================================

(defun table:get-all-header-ids (/ doc layouts layout layout-name blocks
                                    ename header-id ids)
  "Collects all header_ids from all layouts for BOQ-to-PR transfer."
  (vl-load-com)
  (setq ids '())
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-layouts doc))
      (vlax-for layout layouts
        (setq layout-name (vla-get-name layout))
        (if (/= layout-name "Model")
          (progn
            (setq blocks (vla-get-block layout))
            (vlax-for block blocks
              (if (= (vla-get-objectname block) "AcDbTable")
                (progn
                  (setq ename (vlax-vla-object->ename block))
                  (if (table:legal-p ename)
                    (progn
                      (setq header-id (table:get-header-id ename))
                      (if (and header-id
                               (/= header-id "")
                               (distof header-id))
                        (setq ids (cons (atoi header-id) ids))
                      )
                    )
                  )
                )
              )
            )
          )
        )
      )
    )
  )
  (reverse ids)
)

;;; ============================================================
;;; Helper: find layout by name
;;; ============================================================

(defun table:find-layout-by-name (doc layout-name / layouts result)
  "Finds a layout object by name."
  (setq result nil)
  (setq layouts (vla-get-layouts doc))
  (vlax-for layout layouts
    (if (= (strcase layout-name) (strcase (vla-get-name layout)))
      (setq result layout)
    )
  )
  result
)

(princ "\n[OB] table_util.lsp loaded")
(princ)
