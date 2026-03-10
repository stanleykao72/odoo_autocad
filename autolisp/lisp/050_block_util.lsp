;;; block_util.lsp — Block attribute read/write operations
;;; Refactored from legacy/contract_product.lsp
;;; Uses VLA for block attribute access (requires vl-load-com on full AutoCAD)
;;;
;;; Block identification: AcDbBlockReference with "project_name" or
;;; "job_working_plan_name" attribute tag.

(vl-load-com)

;;; ============================================================
;;; Attribute read/write (by VLA object)
;;; ============================================================

(defun block:get-attribute (obj tag / value)
  "Gets attribute value by tag from a block reference VLA object.
   Returns value string or nil."
  (foreach a (vlax-invoke obj 'getattributes)
    (if (= (strcase (vlax-get a 'TagString)) (strcase tag))
      (setq value (vlax-get a 'textString))
    )
  )
  value
)

(defun block:set-attribute (obj tag value)
  "Sets attribute value by tag on a block reference VLA object."
  (foreach a (vlax-invoke obj 'getattributes)
    (if (= (strcase (vlax-get a 'TagString)) (strcase tag))
      (vlax-put a 'textstring value)
    )
  )
)

(defun block:get-all-attributes (obj / result)
  "Gets all attribute values as ((tag . value) ...) from a block reference."
  (setq result '())
  (foreach a (vlax-invoke obj 'getattributes)
    (setq result
      (cons
        (cons (vlax-get a 'TagString) (vlax-get a 'textString))
        result
      )
    )
  )
  (reverse result)
)

(defun block:set-attributes (obj attr-list)
  "Sets multiple attributes. attr-list is ((tag . value) ...)."
  (foreach pair attr-list
    (block:set-attribute obj (car pair) (cdr pair))
  )
)

;;; ============================================================
;;; Find the attribute block in a layout
;;; ============================================================

(defun block:find-attribute-block (/ doc layout blocks block result)
  "Finds the AcDbBlockReference with attributes in the current layout.
   Identifies by presence of 'project_name' or 'job_working_plan_name' tag.
   Returns VLA block reference object or nil."
  (setq result nil)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layout (vla-get-activelayout doc))
      (setq blocks (vla-get-block layout))
      (vlax-for block blocks
        (if (= (vla-get-objectname block) "AcDbBlockReference")
          (progn
            (if (or (block:get-attribute block "project_name")
                    (block:get-attribute block "job_working_plan_name"))
              (setq result block)
            )
          )
        )
      )
    )
  )
  result
)

(defun block:find-attribute-block-in-layout (layout / blocks block result)
  "Finds the attribute block in a specific layout VLA object."
  (setq result nil)
  (if layout
    (progn
      (setq blocks (vla-get-block layout))
      (vlax-for block blocks
        (if (= (vla-get-objectname block) "AcDbBlockReference")
          (progn
            (if (or (block:get-attribute block "project_name")
                    (block:get-attribute block "job_working_plan_name"))
              (setq result block)
            )
          )
        )
      )
    )
  )
  result
)

;;; ============================================================
;;; Header attributes collection (for BOQ export)
;;; ============================================================

;; The 7 attribute tags written to blocks
(setq *ob:attr-tags*
  '("product_name" "spec" "product_catelog" "operation_flow"
    "surface_treatment" "color_name" "color_no"))

;; Additional header tags
(setq *ob:header-tags*
  '("job_working_plan_name" "project_name"))

(defun block:get-header-attrs (block layout-name / attrs result tag val)
  "Gets header attributes from a block for BOQ JSON.
   Returns assoc list with layout_name and all attribute values."
  (setq result nil)
  (if (and block
           (= (vla-get-objectname block) "AcDbBlockReference")
           (equal :vlax-true (vla-get-hasattributes block)))
    (progn
      (setq attrs (block:get-all-attributes block))
      ;; Only return if this is a recognized attribute block
      (if (or (assoc "project_name" attrs)
              (assoc "job_working_plan_name" attrs))
        (progn
          (setq result (list (cons "layout_name" layout-name)))
          ;; Add all known tags
          (foreach tag (append *ob:header-tags* *ob:attr-tags*)
            (setq val (cdr (assoc tag attrs)))
            (if val
              (setq result (cons (cons tag (strip:unformat val)) result))
            )
          )
          (setq result (reverse result))
        )
      )
    )
  )
  result
)

;;; ============================================================
;;; Write attributes to ALL layouts
;;; ============================================================

(defun block:set-attributes-all-layouts (attr-list / doc layouts layout block)
  "Writes attributes to the attribute block in ALL layouts (excluding Model).
   attr-list is ((tag . value) ...)."
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-layouts doc))
      (vlax-for layout layouts
        (if (/= (vla-get-name layout) "Model")
          (progn
            (setq block (block:find-attribute-block-in-layout layout))
            (if block
              (block:set-attributes block attr-list)
            )
          )
        )
      )
    )
  )
)

(princ "\n[OB] block_util.lsp loaded")
(princ)
