;;; param_form.lsp — Parameter selection form logic
;;; Loads Odoo data via bridge → shows DCL form → writes block attributes
;;; Based on legacy/contract_product.lsp + legacy/transfer_to_odoo.lsp (param_form)

;;; ============================================================
;;; Convert bridge response to list-of-values for DCL
;;; ============================================================

(defun param:parse-products (data / result ids names uoms)
  "Parses product data from bridge response into 3 parallel lists.
   Returns: (id-list name-list uom-list)"
  (setq ids '() names '() uoms '())
  (foreach product data
    (setq ids   (cons (cdr (assoc "id" product)) ids))
    (setq names (cons (cdr (assoc "name" product)) names))
    (setq uoms  (cons (cdr (assoc "uom" product)) uoms))
  )
  (list (reverse ids) (reverse names) (reverse uoms))
)

(defun param:parse-setup (data / catalogs specs ops surfs)
  "Parses setup data into 4 category lists.
   Returns: (catalog-list spec-list operation-list surface-list)"
  (setq catalogs '() specs '() ops '() surfs '())
  (foreach item data
    (setq key (cdr (assoc "key" item)))
    (setq val (cdr (assoc "value" item)))
    (cond
      ((= key "product_catelog")    (setq catalogs (cons val catalogs)))
      ((= key "spec")              (setq specs    (cons val specs)))
      ((= key "operation_flow")    (setq ops      (cons val ops)))
      ((= key "surface_treatment") (setq surfs    (cons val surfs)))
    )
  )
  (list (reverse catalogs) (reverse specs) (reverse ops) (reverse surfs))
)

(defun param:parse-colors (data / names nos)
  "Parses color data into 2 parallel lists.
   Returns: (name-list no-list)"
  (setq names '() nos '())
  (foreach color data
    (setq names (cons (cdr (assoc "name" color)) names))
    (setq nos   (cons (cdr (assoc "no" color)) nos))
  )
  ;; Add "No Color" option at start
  (list (cons "No Color" (reverse names))
        (cons "NO_COLOR" (reverse nos)))
)

;;; ============================================================
;;; Load all LOV data from Odoo via bridge
;;; ============================================================

(defun param:load-lov-data (/ resp-prod resp-setup resp-colors
                              product-data setup-data color-data
                              prod-lists setup-lists color-lists)
  "Loads all list-of-values data from Odoo.
   Returns combined LOV list or nil on failure."

  (princ "\n[OB] Loading products from Odoo...")
  (setq resp-prod (bridge:get-products))
  (if (not (bridge:success-p resp-prod))
    (progn
      (princ (strcat "\n[OB] Error: " (bridge:get-message resp-prod)))
      (alert (strcat "Failed to load products:\n" (bridge:get-message resp-prod)))
      nil
    )
    (progn
      (princ "\n[OB] Loading setup values...")
      (setq resp-setup (bridge:get-setup))
      (if (not (bridge:success-p resp-setup))
        (progn
          (alert (strcat "Failed to load setup:\n" (bridge:get-message resp-setup)))
          nil
        )
        (progn
          (princ "\n[OB] Loading colors...")
          (setq resp-colors (bridge:get-colors nil))
          (if (not (bridge:success-p resp-colors))
            (progn
              (alert (strcat "Failed to load colors:\n" (bridge:get-message resp-colors)))
              nil
            )
            (progn
              ;; Parse all data
              (setq prod-lists  (param:parse-products (bridge:get-data resp-prod)))
              (setq setup-lists (param:parse-setup (bridge:get-data resp-setup)))
              (setq color-lists (param:parse-colors (bridge:get-data resp-colors)))

              ;; Return combined: (ids names uoms catalogs specs ops surfs color-names color-nos)
              (list
                (nth 0 prod-lists)   ;; product ids
                (nth 1 prod-lists)   ;; product names
                (nth 2 prod-lists)   ;; product uoms
                (nth 0 setup-lists)  ;; catalogs
                (nth 1 setup-lists)  ;; specs
                (nth 2 setup-lists)  ;; operations
                (nth 3 setup-lists)  ;; surface treatments
                (nth 0 color-lists)  ;; color names
                (nth 1 color-lists)  ;; color nos
              )
            )
          )
        )
      )
    )
  )
)

;;; ============================================================
;;; Product keyword search helpers
;;; ============================================================

(defun param:filter-products (keyword / filtered-idx i name lower-kw)
  "Filters product list by keyword (case-insensitive substring match).
   Sets *param:filtered-idx* to list of original indices.
   Updates the list_box in the active dialog."
  (setq lower-kw (strcase keyword T))
  (setq filtered-idx '())
  (setq i 0)
  (foreach name *param:all-product-names*
    (if (or (= lower-kw "")
            (wcmatch (strcase name T) (strcat "*" lower-kw "*")))
      (setq filtered-idx (cons i filtered-idx))
    )
    (setq i (1+ i))
  )
  (setq *param:filtered-idx* (reverse filtered-idx))

  ;; Update list_box
  (start_list "product_list")
  (foreach idx *param:filtered-idx*
    (add_list (nth idx *param:all-product-names*))
  )
  (end_list)

  ;; Auto-select first if results exist
  (if *param:filtered-idx*
    (progn
      (set_tile "product_list" "0")
      (set_tile "product_uom"
        (nth (car *param:filtered-idx*) *param:all-product-uoms*))
    )
    (set_tile "product_uom" "")
  )
)

(defun param:get-selected-product-index (list-sel-str)
  "Maps list_box selection index back to original product index.
   list-sel-str: $value from list_box (string index in filtered list).
   Returns: original index in product-name-lst, or nil."
  (if (and list-sel-str (/= list-sel-str ""))
    (nth (atoi list-sel-str) *param:filtered-idx*)
    nil
  )
)

;;; ============================================================
;;; Show parameter form DCL
;;; ============================================================

(defun param:show-form (lov-list / product-id-lst product-name-lst product-uom-lst
                          product-catelog-lst spec-lst operation-flow-lst
                          surface-treatment-lst color-name-lst color-no-lst
                          dch userclick
                          SIZ2 SIZ3 SIZ4 SIZ5 SIZ6 SIZ7)
  "Shows the parameter selection DCL form.
   lov-list: output from param:load-lov-data.
   Returns: association list of selected values, or nil if cancelled."

  ;; Unpack LOV lists
  (setq product-id-lst        (nth 0 lov-list))
  (setq product-name-lst      (nth 1 lov-list))
  (setq product-uom-lst       (nth 2 lov-list))
  (setq product-catelog-lst   (nth 3 lov-list))
  (setq spec-lst              (nth 4 lov-list))
  (setq operation-flow-lst    (nth 5 lov-list))
  (setq surface-treatment-lst (nth 6 lov-list))
  (setq color-name-lst        (nth 7 lov-list))
  (setq color-no-lst          (nth 8 lov-list))

  ;; Store full lists in globals for filter callback access
  (setq *param:all-product-names* product-name-lst)
  (setq *param:all-product-uoms*  product-uom-lst)
  (setq *param:filtered-idx*      nil)

  ;; Find and load DCL
  (setq dch (dcl:load "param_form.dcl"))

  (cond
    ((not dch)
     (princ "\n[OB] Error: param_form.dcl not found")
     nil
    )
    ((not (new_dialog "ob_param_form" dch))
     (unload_dialog dch)
     (princ "\n[OB] Error: Could not load param_form dialog")
     nil
    )
    (T
     ;; Populate product list_box (show all initially)
     (param:filter-products "")

     ;; Populate other dropdowns
     (start_list "product_catelog")(mapcar 'add_list product-catelog-lst)(end_list)
     (start_list "spec")(mapcar 'add_list spec-lst)(end_list)
     (start_list "operation_flow")(mapcar 'add_list operation-flow-lst)(end_list)
     (start_list "surface_treatment")(mapcar 'add_list surface-treatment-lst)(end_list)
     (start_list "color_name")(mapcar 'add_list color-name-lst)(end_list)
     (start_list "color_no")(mapcar 'add_list color-no-lst)(end_list)

     ;; Read-only fields
     (mode_tile "color_no" 1)
     (mode_tile "product_uom" 1)

     ;; Product search: filter on Enter / focus-out
     (action_tile "product_search"
       "(param:filter-products $value)"
     )

     ;; Product list selection → auto-fill UOM
     (action_tile "product_list"
       (strcat
         "(progn "
           "(setq SIZ (param:get-selected-product-index $value))"
           "(if SIZ (set_tile \"product_uom\" (nth SIZ *param:all-product-uoms*)))"
         ")")
     )

     ;; Color name → auto-fill color_no
     (action_tile "color_name"
       (strcat
         "(progn "
           "(setq SIZ (atoi $value))"
           "(set_tile \"color_no\" (nth SIZ color-no-lst))"
         ")")
     )

     ;; OK button
     (action_tile "accept"
       (strcat
         "(progn "
           "(setq SIZ2 (param:get-selected-product-index (get_tile \"product_list\")))"
           "(setq SIZ3 (fix (atof (get_tile \"product_catelog\"))))"
           "(setq SIZ4 (fix (atof (get_tile \"spec\"))))"
           "(setq SIZ5 (fix (atof (get_tile \"operation_flow\"))))"
           "(setq SIZ6 (fix (atof (get_tile \"surface_treatment\"))))"
           "(setq SIZ7 (fix (atof (get_tile \"color_name\"))))"
           "(if (not SIZ2)"
             "(alert \"Please select a product from the list.\")"
             "(progn (done_dialog 1) (setq userclick T))"
           ")"
         ")")
     )

     ;; Cancel
     (action_tile "cancel" "(done_dialog 0) (setq userclick nil)")

     (start_dialog)
     (unload_dialog dch)

     ;; Build result
     (if userclick
       (list
         (cons "product_name"       (nth SIZ2 product-name-lst))
         (cons "product_catelog"    (nth SIZ3 product-catelog-lst))
         (cons "spec"              (nth SIZ4 spec-lst))
         (cons "operation_flow"    (nth SIZ5 operation-flow-lst))
         (cons "surface_treatment" (nth SIZ6 surface-treatment-lst))
         (cons "color_name"        (nth SIZ7 color-name-lst))
         (cons "color_no"          (nth SIZ7 color-no-lst))
       )
       nil
     )
    )
  )
)

;;; ============================================================
;;; Apply selected parameters to block attributes
;;; ============================================================

(defun param:apply-to-block (block selections)
  "Applies user selections to block attributes.
   block: VLA block reference object
   selections: ((tag . value) ...) from param:show-form"
  (if (and block selections)
    (progn
      (block:set-attributes block selections)
      (princ "\n[OB] Parameters written to block attributes")
      T
    )
  )
)

(princ "\n[OB] param_form.lsp loaded")
(princ)
