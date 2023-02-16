
//Here's how to design and code a Popup Listbox in DCL :

//DCL CODING STARTS HERE
contract_product : dialog {
label = "���ƤΥ[�u�覡" ;

: spacer { width = 5; }

// : boxed_column {			//define boxed column
//         label = "�X���u��";			//give it a label

//         : row {
//           : popup_list {
//           label = "�X���u���s:";
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
        label = "����";			//give it a label

        : popup_list {
        label = "����:";
        key = "product";
        value = "0" ;
        edit_width = 50;
        }

        : popup_list {
        label = "����:";
        key = "spec";
        value = "0" ;
        edit_width = 20;
        }

        : popup_list {
        label = "���Ƥ���:";
        key = "product_catelog";
        value = "0" ;
        edit_width = 20;
        }

        : spacer { width = 3; }

}

: spacer { width = 5; }

: boxed_column {			//define boxed column
        label = "�[�u";			//give it a label

        : popup_list {
        label = "�[�u�y�{:";
        key = "operation_flow";
        value = "0" ;
        edit_width = 30;
        }

        : popup_list {
        label = "���B�z:";
        key = "surface_treatment";
        value = "0" ;
        edit_width = 20;
        }

        : spacer { width = 3; }

}

: spacer { width = 5; }

: boxed_column {			//define boxed column
        label = "�C��";			//give it a label

        : row {
            : popup_list {
            label = "�C��W��:";
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
