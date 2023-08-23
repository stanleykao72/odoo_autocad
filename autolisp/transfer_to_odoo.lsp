;; call pythoncom

(defun c:L1() 
  (if (vlax-get-or-create-object "Python.ComServer")
    (progn
      (setq python_com (vlax-get-or-create-object "Python.ComServer"))
      (setq odoo (vlax-invoke-method python_com 'odoo_connection ))
      (setq pr_no_list (GetPrNo))
      (setq pr_no (nth 0 pr_no_list))
      (setq block (nth 1 pr_no_list))
      (princ (strcat "pr_no:" pr_no "\n"))
      (setq project_json (vlax-invoke-method python_com 'get_project pr_no))
      ;; (princ project_json)
      (setq project_id (set_header_id python_com project_json))
    )
    (progn
      (princ "請先執行 L0，再執行後續指令(L1/L3/L4).....")
    )
  )
)

(defun c:L3()
  (if (vlax-get-or-create-object "Python.ComServer")
    (progn
      (setq python_com (vlax-get-or-create-object "Python.ComServer"))
      (setq odoo (vlax-invoke-method python_com 'odoo_connection ))
      (setq import_json (GetAllLayoutsInfo))
      ;; (princ (strcat "import_json:" (vl-princ-to-string import_json) "\n"))
      (setq response_json (vlax-invoke-method python_com 'import2boq import_json))
      ;; (princ (strcat "response_json:" (vl-princ-to-string response_json) "\n"))
      (setq response_list (dmc:json:json_to_list response_json nil))
      ;; (princ (strcat "response_list:" (vl-princ-to-string response_list) "\n"))
      
      (if (= (nth 0 (nth 0 response_list)) "error_code")
        (progn
          (setq return_message (cdr (nth 1 response_list)))
        )
        (progn
          (setq update_header_detail (update-header_id-detail_id response_list))
          (setq return_message "Successfully Import to BOQ")
        )
      )
      
    )
    (progn
      (setq return_message "請先執行 L0，再執行後續指令(L1/L3/L4).....")
    )
  )
  return_message
)


(defun c:L4()
  (if (vlax-get-or-create-object "Python.ComServer")
    (progn
      (setq python_com (vlax-get-or-create-object "Python.ComServer"))
      (setq odoo (vlax-invoke-method python_com 'odoo_connection ))
      (setq import_json (TransferToPr))
      ;; (princ (strcat "import_json" (vl-princ-to-string import_json) "\n"))
      (setq response_list (dmc:json:json_to_list import_json nil))
      ;; (princ (strcat "response_list:" (vl-princ-to-string response_list) "\n"))
      
      (if (= (nth 0 (nth 0 response_list)) "error_code")
        (progn
          (setq return_message (cdr (nth 1 response_list)))
        )
        (progn
          (setq response_json (vlax-invoke-method python_com 'boq2pr import_json))
          ;; (princ (strcat "response_json:" (vl-princ-to-string response_json) "\n"))
          (setq return_message "Successfully Transfer to PR")
        )
      )
    )
    (progn
      (setq return_message "請先執行 L0，再執行後續指令(L1/L3/L4).....")
    )
  )
  return_message
)

(defun c:L7()
  (setq clr_header_detail (ClearAllLayoutsTable))
)

(defun c:L9()
  (if (vlax-get-or-create-object "Python.ComServer")
    (progn
      (setq python_com (vlax-get-or-create-object "Python.ComServer"))
      (setq odoo (vlax-invoke-method python_com 'odoo_connection ))
      (setq import_json (GetCurrentLayoutInfo))
      (princ import_json)(princ "\n")
      ;; (setq response_json (vlax-invoke-method python_com 'import2boq import_json))
      ;; (princ response_json)
    )
    (progn
      (princ "請先執行 L0，再執行後續指令(L1/L3/L4).....")
    )
  )
)

;;-------------------=={ UnFormat String }==------------------;;
;;                                                            ;;
;;  Returns a string with all MText formatting codes removed. ;;
;;------------------------------------------------------------;;
;;  Author: Lee Mac, Copyright c 2011 - www.lee-mac.com       ;;
;;------------------------------------------------------------;;
;;  Arguments:                                                ;;
;;  str - String to Process                                   ;;
;;  mtx - MText Flag (T if string is for use in MText)        ;;
;;------------------------------------------------------------;;
;;  Returns:  String with formatting codes removed            ;;
;;------------------------------------------------------------;;

(defun LM:UnFormat ( str mtx / _replace rx )

    (defun _replace ( new old str )
        (vlax-put-property rx 'pattern old)
        (vlax-invoke rx 'replace str new)
    )
    (if (setq rx (vlax-get-or-create-object "VBScript.RegExp"))
        (progn
            (setq str
                (vl-catch-all-apply
                    (function
                        (lambda ( )
                            (vlax-put-property rx 'global     actrue)
                            (vlax-put-property rx 'multiline  actrue)
                            (vlax-put-property rx 'ignorecase acfalse) 
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
            )
        )
    )
)
(vl-load-com)

(defun atoi2(str)
  (if (distof str)(atoi str))
)

(defun table2list ( obj / col lst row tmp )
  (setq lst '())  
  (repeat (setq row (vla-get-rows obj))
      (repeat (setq row (1- row) col (vla-get-columns obj))
          (setq tmp (cons (LM:UnFormat (vla-gettext obj row (setq col (1- col))) nil) tmp))
      )
      (setq lst (cons tmp lst) tmp nil)
  )
  ;(princ (strcat (vl-princ-to-string lst) "\n"))
  lst
)

(defun find-layout-by-name (doc layoutName)
  (setq layoutObj nil)
  (vlax-for layout (vla-get-layouts doc)
    (if (equal (strcase layoutName) (strcase (vla-get-name layout)))
      (setq layoutObj layout)
    )
  )
  layoutObj
)

(defun ClearHeaderDetail (obj)
  (vl-load-com)
  (setq header_label (LM:UnFormat (vla-gettext obj 0 7) nil))
  
  (if (= header_label "HEADER_ID")
    (progn
      ;; (princ (strcat "header_id:" (vla-gettext obj 0 8) ":\n"))
      (vla-SetText obj 0 8 "")
      (setq row -1)
      (setq rows (vla-get-rows obj))
      ;; (setq columns (vla-get-columns obj))
      ;; (princ (strcat "rows:" (itoa rows) ":\n"))
      ;; (princ (strcat "columns:" (itoa columns) ":\n"))
      (repeat rows
        (setq row (1+ row))
        ;; (princ (strcat "row:" (itoa row) ":\n"))
        (if (> row 1)
          (progn
            ;; (princ (strcat "detail_id:" (vla-gettext obj row 8) ":\n"))
            ;; (princ (strcat "product_no:" (vla-gettext obj row 1) ":\n"))
            (vla-SetText obj row 8 "")
          )
        )
      )
      ;; (vla-Update obj)
      (vlax-release-object obj)
      (setq return_str "Updated header_id and detail_id to nil")
    )
    (progn
      (setq return_str "Table is not target, Not Update")
    )
  )
  return_str
)


(defun c:GetCurrentLayoutName ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layout (vla-get-activelayout doc))
      (princ (strcat "\nCurrent Layout: " (vla-get-name layout)))
    )
    (princ "\nNo active document found.")
  )
  (princ)
)

;; Set Attribute Value  -  Lee Mac
;; Sets the value of the first attribute with the given tag found within the block, if present.
;; blk - [vla] VLA Block Reference Object
;; tag - [str] Attribute TagString
;; val - [str] Attribute Value
;; Returns: [str] Attribute value if successful, else nil.
(defun LM:vl-setattributevalue ( blk tag val )
    (setq tag (strcase tag))
    (vl-some
       '(lambda ( att )
            (if (= tag (strcase (vla-get-tagstring att)))
                (progn (vla-put-textstring att val) val)
            )
        )
        (vlax-invoke blk 'getattributes)
    )
)

;; Set Attribute Values  -  Lee Mac
;; Sets attributes with tags found in the association list to their associated values.
;; blk - [vla] VLA Block Reference Object
;; lst - [lst] Association list of ((<tag> . <value>) ... )
;; Returns: nil
(defun LM:vl-setattributevalues ( blk lst / itm )
    (foreach att (vlax-invoke blk 'getattributes)
        (if (setq itm (assoc (vla-get-tagstring att) lst))
            (vla-put-textstring att (cdr itm))
        )
    )
)

(defun get_attribute_block ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layout (vla-get-activelayout doc))
      (setq layout_name (vla-get-name layout))
      (princ (strcat "\nCurrent Layout: " layout_name "\n"))
      
      (setq blocks (vla-get-block layout))

      (vlax-for block blocks        
        ;; (setq header_lst '())
        (if (= (vla-get-ObjectName block) "AcDbBlockReference")
          (progn
            (setq return_block block)
          );progn
        );if
      );vlax
    );progn
  );if
  return_block
)

(defun form_lov (product_list setup_list color_list)
  ;; convert prduct
  (setq product_id_lst nil)
  (setq product_name_lst nil)
  (setq product_uom_lst nil)
  (foreach product product_list
    (setq product_id (cdr (nth 0 product)))
    ;; (princ (strcat "product_id:" (vl-princ-to-string product_id) "\n"))
    (setq product_name (cdr (nth 1 product)))
    (setq product_uom (cdr (nth 2 product)))
    (setq product_id_lst (cons product_id product_id_lst))
    ;; (princ (strcat "product_id_lst:" (vl-princ-to-string product_id_lst) "\n"))
    (setq product_name_lst (cons product_name product_name_lst))
    ;; (princ (strcat "product_name_lst:" (vl-princ-to-string product_name_lst) "\n"))
    (setq product_uom_lst(cons product_uom product_uom_lst))
    ;; (princ (strcat "product_uom_lst:" (vl-princ-to-string product_uom_lst) "\n"))
  )

  ;; convert setup
  (setq product_catelog_lst nil)
  (setq spec_lst nil)
  (setq operation_flow_lst nil)
  (setq surface_treatment_lst nil)

  (foreach setup setup_list
    (setq key (cdr (nth 0 setup)))
    (setq value (cdr (nth 1 setup)))
    ;; (princ (strcat "key:" (vl-princ-to-string key) "\n"))
    ;; (princ (strcat "value:" (vl-princ-to-string value) "\n"))
    (cond 
      ((= key "product_catelog")
       (progn
         (setq product_catelog_lst (cons value product_catelog_lst))
       )
      )
      ((= key "spec")
       (progn
         (setq spec_lst (cons value spec_lst))
       )
      )
      ((= key "operation_flow")
       (progn
         (setq operation_flow_lst (cons value operation_flow_lst))
        ;;  (princ (strcat "operation_flow_lst:" (vl-princ-to-string operation_flow_lst) "\n"))
       )
      )
      ((= key "surface_treatment")
       (progn
         (setq surface_treatment_lst (cons value surface_treatment_lst))
        ;;  (princ (strcat "surface_treatment_lst:" (vl-princ-to-string surface_treatment_lst) "\n"))
       )
      )
    );cond
  );foreach
  ;; (princ (strcat "setup_list:" (vl-princ-to-string setup_list) "\n"))
  ;; (princ (strcat "product_catelog_lst:" (vl-princ-to-string product_catelog_lst) "\n"))
  ;; (princ (strcat "spec_lst:" (vl-princ-to-string spec_lst) "\n"))
  ;; (princ (strcat "operation_flow_lst:" (vl-princ-to-string operation_flow_lst) "\n"))
  ;; (princ (strcat "surface_treatment_lst:" (vl-princ-to-string surface_treatment_lst) "\n"))

  ;; convert color
  (setq color_name_lst nil)
  (setq color_no_lst nil)
  (foreach color color_list
    (setq color_name (cdr (nth 0 color)))
    (setq color_no (cdr (nth 1 color)))
    (setq color_name_lst (cons color_name color_name_lst))
    (setq color_no_lst (cons color_no color_no_lst))
  )
  (setq color_name_lst (cons "No Color" color_name_lst))
  (setq color_no_lst (cons "NO_COLOR" color_no_lst))
  ;; (princ (strcat "color_name_lst:" (vl-princ-to-string color_name_lst) "\n"))
  ;; (princ (strcat "color_no_lst:" (vl-princ-to-string color_no_lst) "\n"))

  (setq return_list (list product_id_lst product_name_lst product_uom_lst product_catelog_lst spec_lst operation_flow_lst surface_treatment_lst color_name_lst color_no_lst))
  return_list
)

(defun param_form (block lov_list)

  (setq product_id_lst (nth 0 lov_list))
  (setq product_name_lst (nth 1 lov_list))
  (setq product_uom_lst (nth 2 lov_list))
  (setq product_catelog_lst (nth 3 lov_list))
  (setq spec_lst (nth 4 lov_list))
  (setq operation_flow_lst (nth 5 lov_list))
  (setq surface_treatment_lst (nth 6 lov_list))
  (setq color_name_lst (nth 7 lov_list))
  (setq color_no_lst (nth 8 lov_list))

  (cond
    (
      (not
        (and
          (setq dcl (findfile "param_form.dcl"))    ;; Check for DCL file
          (< 0 (setq dch (load_dialog dcl)))  ;; Attempt to load it if found
        )
      )

      ;; Else dialog is either not found or couldn't be loaded:

      (princ "\n** DCL File not found **")
    )
    (
      (not (new_dialog "param_form" dch "" (cond ( *screenpoint* ) ( '(-1 -1) ))))

      ;; If our global variable *screenpoint* has a value it will be
      ;; used to position the dialog, else the default (-1 -1) will be
      ;; used to center the dialog on screen.

      ;; Should the dialog definition not exist, we unload the dialog
      ;; file from memory and inform the user:
                                             
      (setq dch (unload_dialog dch))
      (princ "\n** Dialog could not be Loaded **")
    )
    (t

      ;; Dialog loaded successfully, now we define the action_tile
      ;; statements:
      (start_list "product")(mapcar 'add_list product_name_lst)(end_list)
      (start_list "product_uom")(mapcar 'add_list product_uom_lst)(end_list)
      (start_list "product_catelog")(mapcar 'add_list product_catelog_lst)(end_list)
      (start_list "spec")(mapcar 'add_list spec_lst)(end_list)
      (start_list "operation_flow")(mapcar 'add_list operation_flow_lst)(end_list)
      (start_list "surface_treatment")(mapcar 'add_list surface_treatment_lst)(end_list)
      (start_list "color_name")(mapcar 'add_list color_name_lst)(end_list)
      (start_list "color_no")(mapcar 'add_list color_no_lst)(end_list)

      (mode_tile "color_no" 1)
      (action_tile
          "color_name"

          (strcat
              "(progn "
                  "(setq color_name $value) "
                  "(setq SIZ (atoi color_name))"
                  "(setq color_no (nth SIZ color_no_lst)) "
                  "(set_tile \"color_no\" color_no)"
              ")"
          )
      )

      (mode_tile "product_uom" 1)
      (action_tile
          "product"

          (strcat
              "(progn "
                  "(setq product $value) "
                  "(setq SIZ (atoi product))"
                  "(setq product_uom (nth SIZ product_uom_lst)) "
                  "(set_tile \"product_uom\" product_uom)"
              ")"
          )
      )

      (action_tile
          "accept"
          ;if O.K. pressed

          (strcat
              "(progn "
                  ;; "(setq SIZ1 (atof (get_tile \"contract\"))) "
                  "(setq SIZ2 (atof (get_tile \"product\"))) "
                  "(setq SIZ3 (atof (get_tile \"product_catelog\"))) "
                  "(setq SIZ4 (atof (get_tile \"spec\"))) "
                  "(setq SIZ5 (atof (get_tile \"operation_flow\"))) "
                  "(setq SIZ6 (atof (get_tile \"surface_treatment\"))) "
                  "(setq SIZ7 (atof (get_tile \"color_name\"))) "
            
                  "(setq *screenpoint* (done_dialog)) (setq userclick T))"
              ;close dialog

          );strcat

      );action tile

      (action_tile
          "cancel"
          ;if cancel button pressed

          "(done_dialog) (setq userclick nil)"
          ;close dialog

      );action_tile

      ;; The dialog screen position is returned by the done_dialog
      ;; function, so we store this in our global variable *screenpoint*
      ;; for next time.

      (start_dialog)

      ;; Display the dialog - we can use the return of start_dialog
      ;; to determine whether the user pressed OK or Cancel.

      (setq dch (unload_dialog dch))

      ;; We're done, unload the Dialog from memory.
      (if userclick
      ;check O.K. was selected

        (progn
        ;if it was do the following

          (setq SIZ2 (fix SIZ2))
          (setq product_name (nth SIZ2 product_name_lst))
          (LM:vl-setattributevalue block "product_name" product_name)
          (setq SIZ3 (fix SIZ3))
          (setq product_catelog (nth SIZ3 product_catelog_lst))
          (LM:vl-setattributevalue block "product_catelog" product_catelog)
          (setq SIZ4 (fix SIZ4))
          (setq spec (nth SIZ4 spec_lst))
          (LM:vl-setattributevalue block "spec" spec)
          (setq SIZ5 (fix SIZ5))
          (setq operation_flow (nth SIZ5 operation_flow_lst))
          (LM:vl-setattributevalue block "operation_flow" operation_flow)
          (setq SIZ6 (fix SIZ6))
          (setq surface_treatment (nth SIZ6 surface_treatment_lst))
          (LM:vl-setattributevalue block "surface_treatment" surface_treatment)

          ;; (princ "SIZ7:")(princ SIZ7)(princ "\n")
          (setq SIZ7 (fix SIZ7))
          (setq color_name (nth SIZ7 color_name_lst))
          (setq color_no (nth SIZ7 color_no_lst))
          (LM:vl-setattributevalue block "color_name" color_name)
          (LM:vl-setattributevalue block "color_no" color_no)

          ;; (alert (strcat "?z?????: \n" product_name "\n" product_catelog "\n" spec "\n" operation_flow "\n" surface_treatment "\n" color_name "\n ?w??s?I?I?I?I"))
          ;display the Day

        );progn

      );if userclick

    );t
  );cond

(princ)
)

(defun set_header_id (python_com project_json)
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq block (get_attribute_block))
      (setq blockName (vla-get-name block))
      (princ (strcat "Block: " blockName "\n"))
      ;; (princ (strcat "project_json: " project_json "\n"))
      (setq project_list (dmc:json:json_to_list project_json nil))
      ;; (princ (strcat "project_list" (vl-princ-to-string project_list) "\n"))
      (setq project_id (cdr (nth 0 project_list)))
      (princ (strcat "project_id: " (itoa project_id) "\n"))
      (setq project_name (cdr (nth 1 project_list)))
      (princ (strcat "project_name: " project_name "\n"))
      (setq job_working_plan_name (cdr (nth 3 project_list)))
      (princ (strcat "job_working_plan_name: " job_working_plan_name "\n"))

      (setq blockname (vla-get-effectivename block))
      ;; (princ (strcat "blockname: " blockname "\n"))

      (LM:vl-setattributevalue block "job_working_plan_name" job_working_plan_name)
      (LM:vl-setattributevalue block "project_name" project_name)                

      (setq product_json (vlax-invoke-method python_com 'get_product))
      ;; (princ product_json)
      (setq product_list (car (cdr (car (dmc:json:json_to_list product_json nil)))))
      ;; (princ (strcat "product_list" (vl-princ-to-string product_list) "\n"))
      (setq setup_json (vlax-invoke-method python_com 'get_setup))
      (setq setup_list (car (cdr (car (dmc:json:json_to_list setup_json nil)))))
      ;; (princ (strcat "setup_list" (vl-princ-to-string setup_list) "\n"))
      (setq color_json (vlax-invoke-method python_com 'get_color project_id))
      (setq color_list (car (cdr (car(dmc:json:json_to_list color_json nil)))))
      ;; (princ (strcat "color_list" (vl-princ-to-string color_list) "\n"))
      
      (setq lov_list (form_lov product_list setup_list color_list))
      ;; (princ (strcat "lov_list:" (vl-princ-to-string lov_list) "\n"))
      (param_form block lov_list)
      
    );progn
  );if
  ;; project_id
)

(defun get_header_lst (block layout_name pr_no)
  (setq blockName (vla-get-name block))
  ;; (princ (strcat "Block: " blockName "\n"))
  (setq header_lst '())
  (if (equal :vlax-true (vla-get-hasattributes block))
    (progn
      (setq blockname (vla-get-effectivename block))
      ;; (princ (strcat "blockname: " blockname "\n"))
      (setq blockdef  (vla-item
                  (vla-get-blocks
                    (vla-get-activedocument (vlax-get-acad-object))
                  );vla-get-blocks
                  blockname
                );vla-item
      )
      (setq blockdata nil)

      (foreach attrib  (vlax-invoke block 'GetAttributes)
      ;; (princ "attrib:")(princ attrib)(princ "\n")
        (setq tagstring (vla-get-tagstring attrib)
              textstring (vla-get-textstring attrib)
        );setq

        (if (= tagstring "job_working_plan_name")
          (setq job_working_plan_name (LM:UnFormat textstring nil))
        ); if tagstring = job_working_plan_name
        ;; (princ (strcat "job_working_plan_name:" job_working_plan_name "\n"))

        (if (= tagstring "color_name")
          (setq color_name (LM:UnFormat textstring nil))
        ); if tagstring = color_name
        ;; (princ (strcat "color_name" color_name "\n"))

        (if (= tagstring "color_no")
          (setq color_no (LM:UnFormat textstring nil))
        ); if tagstring = color_no
        ;; (princ (strcat "color_no" color_no "\n"))

        (if (= tagstring "surface_treatment")
          (setq surface_treatment (LM:UnFormat textstring nil))
        ); if tagstring = surface_treatment
        ;; (princ (strcat "surface_treatment" surface_treatment "\n"))

        (if (= tagstring "product_name")
          (setq product_name (LM:UnFormat textstring nil))
        ); if tagstring = product_name
        ;; (princ (strcat "product_name" product_name "\n"))

        (if (= tagstring "operation_flow")
          (setq operation_flow (LM:UnFormat textstring nil))
        ); if tagstring = operation_flow
        ;; (princ (strcat "operation_flow" operation_flow "\n"))

        (if (= tagstring "product_catelog")
          (setq product_catelog (LM:UnFormat textstring nil))
        ); if tagstring = product_catelog
        ;; (princ (strcat "product_catelog" product_catelog "\n"))

        (if (= tagstring "spec")
          (setq spec (LM:UnFormat textstring nil))
        ); if tagstring = spec
        ;; (princ (strcat "spec" spec "\n"))

        (if (= tagstring "project_name")
          (setq project_name (LM:UnFormat textstring nil))
        ); if tagstring = project_name
        ;; (princ (strcat "project_name" project_name "\n"))

        (setq header_lst (list (cons "layout_name" layout_name) (cons "pr_no" pr_no)
                               (cons "job_working_plan_name" job_working_plan_name) (cons "color_name" color_name) (cons "color_no" color_no)
                               (cons "surface_treatment" surface_treatment) (cons "product_name" product_name) (cons "operation_flow" operation_flow)
                               (cons "product_catelog" product_catelog) (cons "spec" spec) (cons "project_name" project_name)
                         )
        )

      );foreach
    );progn
      ;; (setq blockdata(reverse blockdata))
      ;; (princ (strcat "header_lst" (vl-princ-to-string header_lst) "\n"))      
  );if
  header_lst  
)

(defun chk_legal_table (block)
  (vl-load-com)
  (setq obj block)
  (setq detail_lst '())
  (setq all_dtl_lst '())

  (if (vla-gettext obj 0 7)
    (progn
      (setq header_label (LM:UnFormat (vla-gettext obj 0 7) nil))
      ;; (princ (strcat "header_label:" header_label ":\n"))
      (setq header_id (vla-gettext obj 0 8))
      
      (if (= header_label "HEADER_ID")
        (progn
          (setq return_str "Y")
        )
        (progn
          (setq return_str "N")
        )
      );if
    );progn
    (progn
      (setq return_str "N")
    )
  );if
  return_str
)

(defun get_detail_lst (block table_count to_append_lst)
  (vl-load-com)
  (setq obj block)
  (setq detail_lst '())
  (setq all_dtl_lst '())

  ;; (if (vla-gettext obj 0 7)
  ;;   (progn
  ;;     (setq header_label (LM:UnFormat (vla-gettext obj 0 7) nil))
  ;;     ;; (princ (strcat "header_label:" header_label ":\n"))
      
  ;;     (if (= header_label "HEADER_ID")
  ;;       (progn
  (setq header_id (vla-gettext obj 0 8))
  (setq row -1)
  (setq rows (vla-get-rows obj))
  ;; (setq columns (vla-get-columns obj))
  ;; (princ (strcat "rows:" (itoa rows) ":\n"))
  ;; (princ (strcat "columns:" (itoa columns) ":\n"))
  (repeat rows
    (setq row (1+ row))
    ;; (princ (strcat "row:" (itoa row) ":\n"))
    (if (> row 1)
      (progn
        (setq qty (vla-gettext obj row 6))
        ;; (princ (strcat "qty" qty ":\n"))
        (setq qty (atoi2 qty))
        (if (and  qty
                  (= (type qty) 'INT)
                  (>  qty 0 )
            )
          (progn
            (setq position (LM:UnFormat (vla-gettext obj row 0) nil))
            (setq product_no (LM:UnFormat (vla-gettext obj row 1) nil))
            ;; (princ (strcat "product_no:" product_no ":\n"))
            (setq width (LM:UnFormat (vla-gettext obj row 2) nil))
            (setq height (LM:UnFormat (vla-gettext obj row 3) nil))
            (setq len (LM:UnFormat (vla-gettext obj row 4) nil))
            (setq thickness (LM:UnFormat (vla-gettext obj row 5) nil))
            ;; (setq qty (vla-gettext obj row 6))
            (setq desc (LM:UnFormat (vla-gettext obj row 7) nil))
            (setq detail_id (LM:UnFormat (vla-gettext obj row 8) nil))
            ;; (princ (strcat "detail_id:" detail_id ":\n"))

            (setq detail_lst (list (cons "header_id" header_id) (cons "detail_id" detail_id) 
                                  (cons "position" position) (cons "product_no" product_no)
                                  (cons "width" width) (cons "height" height) (cons "len" len) (cons "thickness" thickness)
                                  (cons "qty" qty) (cons "desc" desc)
                            )
            )
            ;; (princ (strcat "detail_lst" (vl-princ-to-string detail_lst) "\n"))
            (setq detail_lst (list detail_lst))
            ;; (princ (strcat "detail_lst 2" (vl-princ-to-string detail_lst) "\n"))
            ;; (princ (strcat "all_dtl_lst..." (vl-princ-to-string all_dtl_lst) "\n"))
            (setq all_dtl_lst (append all_dtl_lst detail_lst))
            ;; (princ (strcat "before 1 all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
          );progn
        );if
      );progn
    );if
  );repeat

  (setq all_dtl_lst (append to_append_lst all_dtl_lst))
  ;; (princ (strcat "all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
  ;; (if (= table_count 1)
  ;;   (progn
  ;;     (setq all_dtl_lst (cons (quote <ARRAY>) all_dtl_lst))
  ;;     ;; (princ (strcat "before 2 all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
  ;;     (setq all_dtl_lst (append all_dtl_lst (list (quote </ARRAY>))))
  ;;     ;; (princ (strcat "before 3 all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
  ;;     (setq all_dtl_lst (list (list "detail" all_dtl_lst)))
  ;;     ;; (princ (strcat "before 4 all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
  ;;     ;; (setq all_dtl_lst (list all_dtl_lst))
  ;;     ;; (princ (strcat "after all_dtl_lst:" (vl-princ-to-string all_dtl_lst) "\n"))
  ;;     ;; (setq import_json (dmc:json:list_to_json all_dtl_lst))
  ;;     ;; (princ (strcat "import_json:" import_json "\n"))
  ;;   )
  ;;   (progn
  ;;     (princ (strcat "before 4 to_append_lst:" (vl-princ-to-string to_append_lst) "\n"))    
  ;;   )
  ;; )

        ;; ); progn
  ;;     );if

  ;;   );progn
  ;; );if
  (list all_dtl_lst header_id)
)

(defun add_arrary_tag (be_added_list dict_key)
  (setq be_added_list (cons (quote <ARRAY>) be_added_list))
  ;; (princ (strcat "before 2 be_added_list:" (vl-princ-to-string be_added_list) "\n"))
  (setq be_added_list (append be_added_list (list (quote </ARRAY>))))
  ;; (princ (strcat "before 3 be_added_list:" (vl-princ-to-string be_added_list) "\n"))
  (setq be_added_list (list (list dict_key be_added_list)))
  
  be_added_list
)

(defun get_block_text (block)
  (setq blockname (vla-get-effectivename block))
  (setq blockdef  (vla-item
              (vla-get-blocks
                (vla-get-activedocument (vlax-get-acad-object))
              );vla-get-blocks
              blockname
            );vla-item
  )
  (setq blockdata nil)

  (vlax-for item  blockdef
    ;; (princ (strcat "item:" (vla-get-objectname item) "\n"))
    (if (equal (vla-get-objectname item) "AcDbText")
      (progn
        (setq pr_no (vla-get-textstring item))
        ;; (princ (strcat "text:" pr_no "\n"))
      )
    )
  )                  
  pr_no
)

(defun get_block_table (block)
  (setq blockname (vla-get-effectivename block))
  (setq blockdef  
    (vla-item
      (vla-get-blocks
        (vla-get-activedocument (vlax-get-acad-object))
      );vla-get-blocks
      blockname
    );vla-item
  )
  (setq blockdata nil)

  (vlax-for item  blockdef
    (princ (strcat "item:" (vla-get-objectname item) "\n"))
    (if (equal (vla-get-objectname item) "AcDbTable")
      (progn
        (setq v_block block)
        ;; (princ (strcat "text:" pr_no "\n"))
      )
      (progn
        (setq v_block nil)
      )
    )
  )                  
  v_block
)


(defun GetCurrentLayoutInfo ( / v_header_list v_detail_list)
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layout (vla-get-activelayout doc))
      (setq layout_name (vla-get-name layout))
      (princ (strcat "\nCurrent Layout: " layout_name "\n"))
      
      (setq blocks (vla-get-block layout))

      (setq header_list '())
      (setq v_header_list '())
      (setq detail_list '())
      (setq v_detail_list '())
      (vlax-for block blocks
        ;; (setq v_header_list '())
        ;; (setq v_detail_list '())
        ;; 1. get AcDbBlockReference
        ;; get block
        (if (= (vla-get-ObjectName block) "AcDbBlockReference")
          (progn

            (setq block_name (vla-get-name block))          
            (if (= block_name "pr_no")
              (progn
                (setq pr_no (get_block_text block))
              );progn
            ); if

            (setq header_list (get_header_lst block layout_name pr_no))
            (if header_list
              (progn
                (setq v_header_list header_list)
              )
            )
            ;; (princ (strcat "after call get_heaer_lst function v_header_list:" (vl-princ-to-string v_header_list) "\n"))
          );; progn
        ) ;; end if block=AcDbBlockReference
      ) ;;end (vla-for block blocks)
      
      ;; 2. get AcDbTable
      (setq count 0)
      (princ count)(princ "........\n")
      (vlax-for block blocks
        ;; get table
        (if (= (vla-get-ObjectName block) "AcDbTable")
          (progn
            (if (= (chk_legal_table block) "Y")
              (progn
                ;; (setq v_detail_list '())
                (setq count (1+ count))
                (princ count)(princ "..........\n")
                ;; (setq detail_list '())
                ;; (setq v_detail_list '())
                ;; (setq result_list '())
                (setq result_list (get_detail_lst block count))
                ;; (princ (strcat "after call get_detail_lst function result_list:" (vl-princ-to-string result_list) "\n"))
                
                (setq detail_list (nth 0 result_list))
                (setq header_id (nth 1 result_list))
                (if detail_list
                  (progn
                    (setq v_detail_list detail_list)
                    ;; (princ (strcat "after call get_detail_lst function result_list of v_detail_list" (vl-princ-to-string v_detail_list) "\n"))
                  )
                )
                (if v_header_list
                  (progn
                    (setq header_id_dict (cons "header_id" header_id))
                    ;; (princ (strcat "after call get_detail_lst function result_list of header_id_dict" (vl-princ-to-string header_id_dict) "\n"))
                    (setq v_header_list (cons header_id_dict v_header_list))
                    ;; (princ (strcat "after call get_detail_lst function result_list of v_header_list" (vl-princ-to-string v_header_list) "\n"))
                  )
                  (progn
                    (princ "no v_header_list\n")
                  )
                )
                ;; (princ (strcat "after call get_detail_lst function v_detail_list" (vl-princ-to-string v_detail_list) "\n"))              
              )
            )
          );progn          
        ) ;; if block=AcDbTable
      ) ;;end (vla-for block blocks)
      
      ;; (princ (strcat "v_header_list:" (vl-princ-to-string v_header_list) "\n"))
      ;; (princ (strcat "v_detail_list:" (vl-princ-to-string v_detail_list) "\n"))

      (if (and v_header_list v_detail_list)
       (progn
        (setq v_header_list (append v_header_list v_detail_list))
        ;; (princ (strcat "after header_lst" (vl-princ-to-string v_header_list) "\n"))
        (setq import_json (dmc:json:list_to_json v_header_list))
        ;; (princ (strcat "import_json:" import_json "\n"))
       )
      (progn
        (setq import_json nil)
        (princ "import_jason is nil \n")
      )
      )
    );progn
    (princ "\nNo active document found.")
  )
  import_json
)

(defun GetAllLayoutsInfo ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-Layouts doc))
      (setq all_header_lst '())

      (vlax-for layout layouts
        (setq layout_name (vla-get-name layout))
        (if (/= layout_name "Model")
          (progn
            (princ (strcat "\Layout Name: " layout_name "\n"))

            (setq blocks (vla-get-block layout))

            (setq header_list '())
            (setq v_header_list '())
            (setq detail_list '())
            (setq v_detail_list '())
            ;; 1. get AcDbBlockReference
            (vlax-for block blocks
              ;; (princ block)(princ (vla-get-ObjectName block))(princ "\n")
              ;; get block
              (if (= (vla-get-ObjectName block) "AcDbBlockReference")
                (progn
                  (setq block_name (vla-get-name block))          
                  (if (= block_name "pr_no")
                    (progn
                      (setq pr_no (get_block_text block))
                    );progn
                  ); if

                  (setq header_list (get_header_lst block layout_name pr_no))
                  (if header_list
                    (progn
                      (setq v_header_list (append v_header_list header_list))
                    )
                  )
                  ;; (princ (strcat "after call get_heaer_lst function v_header_list:" (vl-princ-to-string v_header_list) "\n"))
                );; progn
              ) ;; end if block=AcDbBlockReference
            ) ;;end (vla-for block blocks)
            
            ;; 2. get AcDbTable
            (setq table_count 0)
            (vlax-for block blocks
              ;; (setq v_detail_list '())
              ;; get table
              (if (= (vla-get-ObjectName block) "AcDbTable")
                (progn
                  (if (= (chk_legal_table block) "Y")
                    (progn
                      (setq table_count (1+ table_count))
                      (if (= table_count 1)
                        (setq to_append_list '())
                      )
                      (setq result_list (get_detail_lst block table_count to_append_list))
                      ;; (princ (strcat "after call get_detail_lst function result_list:" (vl-princ-to-string result_list) "\n"))
                      ;; (princ "=================================================\n")
                      
                      (setq detail_list (nth 0 result_list))
                      (setq to_append_list detail_list)
                      (setq header_id (nth 1 result_list))
                      (if detail_list
                        (progn
                          (setq v_detail_list detail_list)
                          
                        )
                      )

                    );progn
                  );if


                );progn          
              ) ;; if block=AcDbTable
            ) ;;end (vla-for block blocks)
            
            (if v_header_list
              (progn
                (if (= table_count 1)
                  (progn
                    (setq header_id_dict (cons "header_id" header_id))
                    ;; (princ (strcat "after call get_detail_lst function result_list of header_id_dict" (vl-princ-to-string header_id_dict) "\n"))
                    (setq v_header_list (cons header_id_dict v_header_list))
                    ;; (princ (strcat "after call get_detail_lst function result_list of v_header_list" (vl-princ-to-string v_header_list) "\n"))
                    ;; (princ "******************************************************\n")
                  )
                )
              )
            )
            ;; (princ (strcat "after call get_detail_lst function v_detail_list" (vl-princ-to-string v_detail_list) "\n"))

            ;; (princ (strcat "header_lst" (vl-princ-to-string header_lst) "\n"))
            ;; (princ (strcat "v_detail_list" (vl-princ-to-string v_detail_list) "\n"))
            (if v_detail_list
              (progn
                (if v_header_list
                  (progn
                    (setq v_detail_list (add_arrary_tag v_detail_list "detail"))
                    (setq v_header_list (list (append v_header_list v_detail_list)))
                    ;; (princ (strcat "after header_list:" (vl-princ-to-string v_header_list) "\n"))
                    ;; (princ "******************************************************\n")
                  )
                )
              )
              (progn
                (setq v_header_list nil)
              )
            )
                
          );progn
        );if

        (if v_header_list
          (progn
            (setq all_header_lst (append all_header_lst v_header_list))
          )
        )
        ;; (princ (strcat "after all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
      );for layouts
      
      (if all_header_lst
        (progn
          (setq all_header_lst (cons (quote <ARRAY>) all_header_lst))
          ;; (princ (strcat "before 2 all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
          (setq all_header_lst (append all_header_lst (list (quote </ARRAY>))))
          ;; (princ (strcat "before 3 all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
          (setq all_header_lst (list (list "all" all_header_lst)))
          ;; (princ (strcat "before 4 all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
          (setq all_import_json (dmc:json:list_to_json all_header_lst))
          ;; (princ (strcat "all_import_json" all_import_json "\n"))            
        )
      )
    );progn
  );if
  all_import_json
)

(defun ClearAllLayoutsTable ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-Layouts doc))
      (setq all_header_lst '())

      (vlax-for layout layouts
        (setq layout_name (vla-get-name layout))
        (princ (strcat "layout_name:" layout_name "\n"))
        (if (/= layout_name "Model")
          (progn
            (setq blocks (vla-get-block layout))
            (vlax-for block blocks
              ;; get table
              (if (= (vla-get-ObjectName block) "AcDbTable")
                (progn
                  (setq result (ClearHeaderDetail block))
                  ;; (vla-update block)
                  ;; (vlax-release-object block)
                  (princ (strcat "result:" result "\n"))
                )
              );if
            );vlax-for
          );progn
        );if
      );vlax-for
    );progn
  );if
)

(defun update-detail_id (obj targetValue newValue)
  (setq colIdx_Field1 1) ; ??@?????? (product_no)
  (setq colIdx_Field8 8) ; ??K?????? (detail_id)

  (setq row -1)
  (setq rows (vla-get-rows obj))
  ;; (princ (strcat "rows:" (itoa rows) ":\n"))
  
  (repeat rows
    (setq row (1+ row))
    ;; (princ (strcat "row:" (itoa row) ":\n"))
    (setq fieldValue_Field1  (LM:UnFormat (vla-gettext obj row colIdx_Field1) nil))
    (princ (strcat "fieldValue_Field1..........:" (vla-gettext obj row colIdx_Field1) ":\n"))
    (princ (strcat "fieldValue_Field1..........(unformate):" (LM:UnFormat (vla-gettext obj row colIdx_Field1) nil) ":\n"))
    (if (= fieldValue_Field1 targetValue)
      (vla-settext obj row colIdx_Field8 newValue)
      (setq return_row row)
    )
  )
  ;; (vlax-release-object obj)
  return_row
)

(defun update-table (obj header_id detail_list)
  
  (if (vla-gettext obj 0 7)
    (progn
      (setq header_label (LM:UnFormat (vla-gettext obj 0 7) nil))
      ;; (princ (strcat "header_label:" header_label ":\n"))
      
      (if (= header_label "HEADER_ID")
        (progn
          ;; update header_id
          (vla-settext obj 0 8 header_id)
          
          (foreach d_lst detail_list
            ;; (princ (strcat "d_lst" (vl-princ-to-string d_lst) "\n"))
            (setq product_no (cdr (nth 0 d_lst)))
            ;; (princ (strcat "product_no:" product_no "\n"))
            (setq detail_id (cdr (nth 1 d_lst)))
            ;; (princ (strcat "detail_id:" (itoa detail_id) "\n"))

            (setq return_row (update-detail_id obj product_no detail_id))
            ;; (princ (strcat "return_row:" (itoa return_row) "\n"))
          )
        );progn
      );if
    );progn
  
  );if
  (vlax-release-object obj)
  return_row
)

(defun update-header_id-detail_id (response_list)
  (vl-load-com)
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  (if doc
    (progn
      ;; (setq response_list (dmc:json:json_to_list response_json nil))
      ;; (princ (strcat "response_list:" (vl-princ-to-string response_list) "\n"))
      ;; (princ (strcat "response_list:" "=======================================" "\n"))
      (setq all_list (nth 1 (nth 0 response_list)))
      (foreach layout_list all_list
        ;; (princ (strcat "layout_list:" (vl-princ-to-string layout_list) "\n"))
        ;; (princ (strcat "layout_list:" "=======================================" "\n"))
        (setq header_id (cdr (nth 0 layout_list)))
        ;; (princ (strcat "header_id:" (itoa header_id) "\n"))
        (setq layout_name (cdr (nth 1 layout_list)))
        ;; (princ (strcat "layout_name:" layout_name "\n"))
        (setq detail_list (car (cdr (nth 2 layout_list))))
        ;; (princ (strcat "detail_list:" (vl-princ-to-string detail_list) "\n"))


        (setq layout (find-layout-by-name doc layout_name))
        ;; (princ layout)
        (setq blocks (vla-get-block layout))
        (vlax-for block blocks
          ;; get table
          (if (= (vla-get-ObjectName block) "AcDbTable")
            (progn
              (setq return_row (update-table block header_id detail_list))
              ;; (princ (strcat "return_row:" (itoa return_row) "\n"))
            );progn
          );if
        );vlax

      )

    );progn
  );if
)

(defun TransferToPr ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layouts (vla-get-Layouts doc))
      (setq all_header_lst '())

      (vlax-for layout layouts
        (setq layout_name (vla-get-name layout))
        (if (/= layout_name "Model")
          (progn
            (princ (strcat "\Layout Name: " layout_name "\n"))

            (setq blocks (vla-get-block layout))
            ;; 2. get AcDbTable
            (setq table_count 0)
            (vlax-for block blocks
              ;; get table
              (if (= (vla-get-ObjectName block) "AcDbTable")
                (progn
                  (if (= (chk_legal_table block) "Y")
                    (progn
                      (setq table_count (1+ table_count))
                      (if (= table_count 1)
                        (progn
                          (setq to_append_list '())
                          (setq header_id nil)
                          (setq result_list '())
                          (setq result_list (get_detail_lst block table_count to_append_list))

                          ;; (princ (strcat "result_list:" (vl-princ-to-string result_list) "\n"))
                          (setq detail_list (nth 0 result_list))
                          ;; (princ (strcat "detail_list:" (vl-princ-to-string detail_list) "\n"))
                          (if detail_list
                            (progn
                              (setq header_id (nth 1 result_list))
                              (princ (strcat "header_id:" header_id ":\n"))
                              (if (and header_id (atoi2 header_id))
                                (progn
                                  ;; (princ (strcat "header_id 2:" header_id ":\n"))
                                  (setq all_header_lst (cons header_id all_header_lst))
                                )
                              )
                            );progn
                          );if                        
                        )
                      );if                      
                    );progn
                  );if
                );progn
              );if
            );vlax-for block
          
          );progn
        );if
      );vlax-for layout
    );progn
  );if
  (setq all_header_lst (reverse all_header_lst))
  ;; (princ (strcat "all_header_lst:" (vl-princ-to-string all_header_lst) "\n"))

  (setq all_header_lst (cons (quote <ARRAY>) all_header_lst))
  ;; (princ (strcat "before 2 all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
  (setq all_header_lst (append all_header_lst (list (quote </ARRAY>))))
  ;; (princ (strcat "before 3 all_header_lst" (vl-princ-to-string all_header_lst) "\n"))
  (setq all_header_lst (list (list "all" all_header_lst)))

  (setq all_header_json (dmc:json:list_to_json all_header_lst))
  ;; (princ (strcat "all_header_json:" all_header_json "\n"))
  
  all_header_json
)

(defun GetPrNo ()
  (vl-load-com)
  (if (setq doc (vla-get-activedocument (vlax-get-acad-object)))
    (progn
      (setq layout (vla-get-activelayout doc))
      (setq layout_name (vla-get-name layout))
      (princ (strcat "\nCurrent Layout: " layout_name "\n"))
      
      (setq blocks (vla-get-block layout))

      (vlax-for block blocks
        ;; 1. get AcDbBlockReference
        ;; get block
        (if (= (vla-get-ObjectName block) "AcDbBlockReference")
          (progn

            (setq block_name (vla-get-name block))          
            (if (= block_name "pr_no")
              (progn
                (setq pr_no (get_block_text block))
                (setq return_block block)
              );progn
            ); if
          );; progn
        ) ;; end if block=AcDbBlockReference
      ) ;;end (vla-for block blocks)
    );progn
  );if
  (list pr_no return_block)
)
