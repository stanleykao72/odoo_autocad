(vl-load-com)

(defun ReadList (lst /)
 (cond
   ((= (type lst) 'Str) (ReadList (read (strcat "(" (vl-string-translate ",;" "  " lst) ")"))))
   ((= (type lst) 'Sym) (eval lst))
   ((= (type lst) 'List) (mapcar (function ReadList) lst))
   (t lst)
 )
)

(defun SelectObjects (pt filter / ss n lst)
 (vl-load-com)
 (vl-princ-to-string pt)
 (if pt
   (progn
     (setq pt (ReadList pt))
     (if (= (type pt) 'List)
       (cond
         ((and (>= (length pt) 2) (vl-every (function (lambda (i) (member (type i) '(Int Real)))) pt))
          (setq n  (list (/ (getvar 'ViewSize) (* (getvar 'PickBox) 80.0)) 0.0)
                n  (cons (car n) n)
                ss (ssget "C" (mapcar '- pt n) (mapcar '+ pt n) filter)
          )
         )
         ((vl-every (function
                      (lambda (i)
                        (and (= (type i) 'List) (vl-every (function (lambda (j) (member (type j) '(Int Real)))) i))
                      )
                    )
                    pt
          );vl-every
          (cond
            ((= (length pt) 1) (setq lst (SelectObjects (car pt) filter)))
            ((= (length pt) 2) (setq ss (ssget "C" (car pt) (cadr pt) filter)))
	    ;((= (length pt) 3) (setq ss (ssget "C" (car pt) (cadr pt) filter)))
            (t (setq ss (ssget "CP" pt filter)))
          );cond
         );cond
       );if
     );progn
   );if
   (setq ss (ssget "X" filter))
 )
 (if ss
   (progn
     (StripMtext ss "CFH")
     (setq n (sslength ss)) (while (>= (setq n (1- n)) 0) (setq lst (cons (ssname ss n) lst)))
   )
 )
 lst
)

(defun ChangeTableTextColor (pt / lst eo f row col str)
 (vl-load-com)
 (if (setq lst (SelectObjects pt '((0 . "ACAD_TABLE")) ) )
   (progn
     
     (foreach en lst
       (setq tbl (vlax-ename->vla-object en)
             row -1
       )
        (setq acCol (vla-get-truecolor tbl));get current AcCmColor object, might be released at the end
        ;(princ "acCol:")(princ acCol)(princ "\n")

        (while (< (setq row (1+ row)) (vla-get-Rows tbl))
          (setq col -1)
          (while (< (setq col (1+ col)) (vla-get-Columns tbl))
            ;(vla-put-colormethod acCol acColorMethodByACI)
            (vla-put-colormethod acCol acColorMethodByRGB)
            ;(vla-setrgb acCol 0 0 0)
            ;(vla-put-colorindex acCol 0);set colorindex for AcCmColor object
            ;(vla-setcellbackgroundcolornone tbl row col :vlax-false);allow to background cell
            ;(vla-setcellbackgroundcolor tbl row col accol);set background using AcCmColor object settings

            ;(vla-put-colormethod acCol acColorMethodByRGB)
            (vla-setrgb acCol 255 255 255);set color by RGB for AcCmColor object
            ;(vla-put-colorindex acCol 5);you may want to set colorindex for AcCmColor object
            ;(vla-settext tbl 0 0 "Blah");add any text
            (vla-setcellcontentcolor tbl row col acCol);set cell text color using AcCmColor object settings
            
          );while
        );while
        (vlax-release-object acCol);release AcCmColor object
     );foreach
   );progn
 );if
(princ)
);defun
     
;(ChangeColor '((67 169 0) (229 231 0)))

(defun c:ChangeTableTextColor()
  (vl-load-com)
  (setq 1pt (getpoint "\n 請選擇基點: "))
  (princ (strcat (vl-princ-to-string 1pt) "\n"))
  (setq 2pt (getcorner 1pt "\n 請選擇對角點: "))
  (princ (strcat (vl-princ-to-string 2pt) "\n"))
  (setq lst (list 1pt 2pt))

  (princ (strcat (vl-princ-to-string lst) "\n"))

  (ChangeTableTextColor lst)
 )

(defun get-attribute (obj tag / value)
  ;; returns the value of the textstring property of the attribute reference
  (foreach a (vlax-invoke obj 'getattributes)
    (if (= (vlax-get a 'TagString) tag)
        (setq value (vlax-get a 'textString))))
  value)

(defun set-attribute (obj tag value)
  ;; sets the value of the textstring with given tag
  (foreach a (vlax-invoke obj 'getattributes)
    (if (= (vlax-get a 'TagString) tag)
        (vlax-put a 'textstring value))))

(defun get-attribute-values (obj / value)
  ;; returns a list of textstring values as (tag . value) pairs
  (foreach a (vlax-invoke obj 'getattributes)
    (setq value (cons (cons (vlax-get a 'TagString)
                            (vlax-get a 'textString))
                      value)))
  (reverse value))

(defun show-attrs (obj)
  (foreach o (get-attribute-values obj)
    (print o)))

(defun GetBlockAttribute (pt)
  (vl-load-com)
  (if (setq lst (SelectObjects pt '((0 . "INSERT"))))
    (progn
      (foreach en lst
        (princ "en:")(princ en)(princ "\n")
        (setq blockref (vlax-ename->vla-object en)
              blockname (vla-get-effectivename blockref)
              blockdef  (vla-item
                          (vla-get-blocks
                            (vla-get-activedocument (vlax-get-acad-object))
                          );vla-get-blocks
                          blockname
                        );vla-item
              blockdata nil
	      )
        (if (equal :vlax-true (vla-get-hasattributes blockref))
          (progn
            (foreach attrib  (vlax-invoke blockref 'GetAttributes)
              (princ "attrib:")(princ attrib)(princ "\n")
              (vlax-for item  blockdef
                (if (equal (vla-get-objectname item) "AcDbAttributeDefinition")
                  (progn
                    (if (equal (vla-get-tagstring attrib) (vla-get-tagstring item))
                      (progn
                        (princ "item:")(princ item)(princ "\n")
                        (setq promptstring (vla-get-promptstring item)
                              tagstring (vla-get-tagstring attrib)
                              textstring (vla-get-textstring attrib)
                        );setq
                        (princ "promptstring:")(princ promptstring)(princ "\n")
                        (princ "tagstring:")(princ tagstring)(princ "\n")
                        (princ "textstring:")(princ textstring)(princ "\n")
                        ;(if (equal promptstring "aa")
                        ;  (vlax-put attrib 'textstring "replace aa")
                         
                        ;);if
                        (setq tmp	(list
				                            (vla-get-promptstring item)
				                            (vla-get-tagstring attrib)
				                            (vla-get-textstring attrib)
                                   );list
                        );setq
		                    (setq blockdata (cons tmp blockdata))
                      );progn
                    );if
                  );progn
                );if
              );vlax-for
            );foreach
            (setq blockdata(reverse blockdata))
            ;; (princ blockdata)
            ;; (foreach lst blockdata
            ;;   (princ (strcat "\n Prompt: "
            ;;           (car lst)
            ;;           " *** Tag: "
            ;;           (cadr lst)
            ;;           " *** Value: "
            ;;           (last lst)
            ;;          );strcat
            ;;   );princ         
            ;; );foreach
          );progn
          (princ "\n  >>  Nothing selected. Try again...")
        );if
        
      );foreach
    );progn
  ;(princ)

);if
)

(defun SetBlockAttribute (pt tag value)
  (vl-load-com)
  (if (setq lst (SelectObjects pt '((0 . "INSERT"))))
    (progn
      (foreach en lst
        ;(princ "en:")(princ en)(princ "\n")
        (setq blockref (vlax-ename->vla-object en)
              blockname (vla-get-effectivename blockref)
              blockdata nil
	      )
        (if (equal :vlax-true (vla-get-hasattributes blockref))
          (progn
            (foreach attrib  (vlax-invoke blockref 'GetAttributes)
              ;(princ "attrib:")(princ attrib)(princ "\n")
                    (if (equal (vla-get-tagstring attrib) tag)
                      (progn
                        (setq tagstring (vla-get-tagstring attrib)
                              textstring (vla-get-textstring attrib)
                        );setq
                        ;(princ "tagstring:")(princ tagstring)(princ "\n")
                        ;(princ "textstring:")(princ textstring)(princ "\n")
                        (vlax-put attrib 'textstring value)                         
                      );progn
                    );if
            );foreach
          );progn
          (princ "\n  >>  Nothing selected. Try again...")
        );if        
      );foreach
    );progn
  ;(princ)
 );if
)

(defun contract_product (product_file setup_file color_file)


  (setq point1 (getpoint '(0 0) "請選擇基點: "))
  (princ "point1")(princ point1)(princ "\n")
  (setq point2 (getcorner point1 "請選擇對角點: " ))
  (princ "point2")(princ point2)(princ "\n")
  
  (defun update_block_text (block_name new_text)
    (setq blks (vla-get-blocks
                  (vla-get-ActiveDocument (vlax-get-acad-object))
                )
    )
    (setq blk (vla-item blks block_name))
    (vlax-for obj blk
      (if (eq "AcDbText" (vla-get-objectname obj))
        (progn
          ;(setq text (vla-get-TextString obj))
          ;(princ text)
          (vla-put-TextString obj new_text)
        );progn
      );if
      (vla-update obj)
    );valx-for  
  );defun update_block_text

  ;; Error Handler so that we may unload the dialog
  ;; from memory should the user hit Esc.

    (defun *error* ( msg )
      (if dch (unload_dialog dch))
      (or (wcmatch (strcase msg) "*BREAK,*CANCEL*,*EXIT*")
          (princ (strcat "\n** Error: " msg " **")))
      (princ)
    )
  
    (setq product_list (LM:readcsv product_file))
    (setq product_id_lst nil)
    (setq product_name_lst nil)
    (setq product_uom_lst nil)
    (foreach product product_list
        (setq product_id (nth 0 product))
        (setq product_name (nth 1 product))）
        (setq product_uom (nth 2 product))）
        (setq product_list (cons product_id product_id_lst))
        (setq product_name_lst (cons product_name product_name_lst))
        (setq product_uom_lst(cons product_uom product_uom_lst))
    )
    ;(princ product_name_lst)
    ;(setq product_lst (nth 1 product_list))

    ;; (setq contract_list (LM:readcsv contract_file))
    ;; (setq contract_id_lst nil)
    ;; (setq reference_lst nil)
    ;; (setq contract_item_lst nil)
    ;; (foreach contract contract_list
    ;;     (setq contract_id (nth 0 contract))
    ;;     (setq contract_reference (nth 1 contract))）
    ;;     (setq contract_item (nth 2 contract))）
    ;;     (setq contract_id_lst (cons contract_id contract_id_lst))
    ;;     (setq reference_lst (cons contract_reference reference_lst))
    ;;     (setq contract_item_lst(cons contract_item contract_item_lst))
    ;; )
    ;(princ reference_lst)
  ;; Plentiful Error trapping to make sure
  ;; dialog is loaded successfully.

    (setq setup_list (LM:readcsv setup_file))
    (setq product_catelog_lst nil)
    (setq spec_lst nil)
    (setq operation_flow_lst nil)
    (setq surface_treatment_lst nil)

    (foreach setup setup_list
      (setq key (nth 0 setup))
      (setq value (nth 1 setup))
      (cond 
        ((= key "product_catelog") (setq product_catelog_lst (cons value product_catelog_lst)))
        ((= key "spec") (setq spec_lst (cons value spec_lst)))
        ((= key "operation_flow") (setq operation_flow_lst (cons value operation_flow_lst)))
        ((= key "surface_treatment") (setq surface_treatment_lst (cons value surface_treatment_lst)))
      );cond
    )
    ;(princ product_catelog_lst)

    (setq color_list (LM:readcsv color_file))
    (setq color_name_lst nil)
    (setq color_no_lst nil)

    (foreach color color_list
      (setq color_name (nth 0 color))
      (setq color_no (nth 1 color))
        (setq color_name_lst (cons color_name color_name_lst))
        (setq color_no_lst (cons color_no color_no_lst))
    )
    (setq color_name_lst (cons "No Color" color_name_lst))
    (setq color_no_lst (cons "NO_COLOR" color_no_lst))
  
    (princ "color_name_list")(princ color_name_lst)(princ "\n")


  (cond
    (
      (not
        (and
          (setq dcl (findfile "contract_product.dcl"))    ;; Check for DCL file
          (< 0 (setq dch (load_dialog dcl)))  ;; Attempt to load it if found
        )
      )

      ;; Else dialog is either not found or couldn't be loaded:

      (princ "\n** DCL File not found **")
    )
    (
      (not (new_dialog "contract_product" dch "" (cond ( *screenpoint* ) ( '(-1 -1) ))))

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
      ;(set_tile "contract" "合約工項編號:")
      ;; (start_list "contract")(mapcar 'add_list reference_lst)(end_list)
      (start_list "product")(mapcar 'add_list product_name_lst)(end_list)
      (start_list "product_catelog")(mapcar 'add_list product_catelog_lst)(end_list)
      (start_list "spec")(mapcar 'add_list spec_lst)(end_list)
      (start_list "operation_flow")(mapcar 'add_list operation_flow_lst)(end_list)
      (start_list "surface_treatment")(mapcar 'add_list surface_treatment_lst)(end_list)
      (start_list "color_name")(mapcar 'add_list color_name_lst)(end_list)
      (start_list "color_no")(mapcar 'add_list color_no_lst)(end_list)

      ;; (mode_tile "contract_item" 1)
      (mode_tile "color_no" 1)

      ;; (action_tile
      ;;     "contract"

      ;;     (strcat
      ;;         "(progn "
      ;;             "(setq reference $value) "
      ;;             "(setq SIZ (atoi reference))"
      ;;             "(setq contract_item (nth SIZ contract_item_lst)) "
      ;;             "(set_tile \"contract_item\" contract_item)"

      ;;         ")"
      ;;     )
      ;; )
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

              ;; (setq SIZ1 (fix SIZ1))
              ;; (setq reference (nth SIZ1 reference_lst))
              ;; (setq contract_item (nth SIZ1 contract_item_lst))
              ;; (update_block_text "reference" reference)
              ;; (update_block_text "contract_item" contract_item)
              (setq SIZ2 (fix SIZ2))
              (setq pt (list point1 point2))
              (setq product_name (nth SIZ2 product_name_lst))
              (SetBlockAttribute pt "product_name" product_name)
              (setq SIZ3 (fix SIZ3))
              (setq product_catelog (nth SIZ3 product_catelog_lst))
              (SetBlockAttribute pt "product_catelog" product_catelog)
              (setq SIZ4 (fix SIZ4))
              (setq spec (nth SIZ4 spec_lst))
              (SetBlockAttribute pt "spec" spec)
              (setq SIZ5 (fix SIZ5))
              (setq operation_flow (nth SIZ5 operation_flow_lst))
              (SetBlockAttribute pt "operation_flow" operation_flow)
              (setq SIZ6 (fix SIZ6))
              (setq surface_treatment (nth SIZ6 surface_treatment_lst))
              (SetBlockAttribute pt "surface_treatment" surface_treatment)

              (princ "SIZ7:")(princ SIZ7)(princ "\n")
              (setq SIZ7 (fix SIZ7))
              (setq color_name (nth SIZ7 color_name_lst))
              (setq color_no (nth SIZ7 color_no_lst))
              (SetBlockAttribute pt "color_name" color_name)
              (SetBlockAttribute pt "color_no" color_no)

              (alert (strcat "您選取的: \n" product_name "\n" product_catelog "\n" spec "\n" operation_flow "\n" surface_treatment "\n" color_name "\n 已更新！！！！"))
              ;display the Day

          );progn

      );if userclick

    );t
  );cond

(princ)
  
);defun