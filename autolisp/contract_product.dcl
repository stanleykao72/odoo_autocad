
//Here's how to design and code a Popup Listbox in DCL :

//DCL CODING STARTS HERE
contract_product : dialog {
label = "材料及加工方式" ;

: spacer { width = 5; }

// : boxed_column {			//define boxed column
//         label = "合約工項";			//give it a label

//         : row {
//           : popup_list {
//           label = "合約工項編:";
//           key = "contract";
//           value = "0" ;
//           edit_width = 12;
//           }

//         : edit_box {				//*define edit box
//           key = "contract_item" ;				//*give it a name
//           //label = "?X???u??" ;	//*give it a label
//           edit_width = 50 ;			//*6 characters only
//           }					//*end edit box

//         }

//         : spacer { width = 3; }
// }

// : spacer { width = 5; }

: boxed_column {			//define boxed column
        label = "材料";			//give it a label

        : popup_list {
        label = "材料:";
        key = "product";
        value = "0" ;
        edit_width = 50;
        }

        : popup_list {
        label = "材質:";
        key = "spec";
        value = "0" ;
        edit_width = 20;
        }

        : popup_list {
        label = "材料分類:";
        key = "product_catelog";
        value = "0" ;
        edit_width = 20;
        }

        : spacer { width = 3; }

}

: spacer { width = 5; }

: boxed_column {			//define boxed column
        label = "加工";			//give it a label

        : popup_list {
        label = "加工流程:";
        key = "operation_flow";
        value = "0" ;
        edit_width = 30;
        }

        : popup_list {
        label = "表面處理:";
        key = "surface_treatment";
        value = "0" ;
        edit_width = 20;
        }

        : spacer { width = 3; }

}

: spacer { width = 5; }

: boxed_column {			//define boxed column
        label = "顏色";			//give it a label

        : row {
            : popup_list {
            label = "顏色名稱:";
            key = "color_name";
            value = "0" ;
            edit_width = 20;
            }

            : edit_box {				//*define edit box
              key = "color_no";				//*give it a name
              //label = "?X???u??" ;	//*give it a label
              edit_width = 50 ;			//*6 characters only
              }					//*end edit box

        }

        : spacer { width = 3; }

}

: spacer { width = 5; }

ok_cancel ;

}
//DCL CODING ENDS HERE
//--------------------------------------------
//Save this as "contract_product.dcl"
