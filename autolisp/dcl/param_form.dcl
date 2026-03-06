// param_form.dcl — Parameter selection form
// Product uses edit_box + list_box for keyword search
// Other fields use popup_list dropdowns

ob_param_form : dialog {
  label = "Set Parameters";
  key = "param_dlg";

  : boxed_column {
    label = "Material";

    : row {
      : edit_box {
        label = "Product Search:";
        key = "product_search";
        edit_width = 40;
      }
      : edit_box {
        key = "product_uom";
        edit_width = 10;
      }
    }

    : list_box {
      key = "product_list";
      height = 8;
      width = 55;
    }

    : popup_list {
      label = "Spec:";
      key = "spec";
      value = "0";
      edit_width = 30;
    }

    : popup_list {
      label = "Product Catalog:";
      key = "product_catelog";
      value = "0";
      edit_width = 30;
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Processing";

    : popup_list {
      label = "Operation Flow:";
      key = "operation_flow";
      value = "0";
      edit_width = 30;
    }

    : popup_list {
      label = "Surface Treatment:";
      key = "surface_treatment";
      value = "0";
      edit_width = 30;
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Color";

    : row {
      : popup_list {
        label = "Color Name:";
        key = "color_name";
        value = "0";
        edit_width = 20;
      }

      : edit_box {
        key = "color_no";
        edit_width = 30;
      }
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  ok_cancel;
}
