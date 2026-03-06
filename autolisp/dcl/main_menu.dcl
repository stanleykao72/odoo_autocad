// main_menu.dcl — Main menu dialog
// Odoo-AutoCAD Integration

ob_main_menu : dialog {
  label = "Odoo-AutoCAD Integration";
  key = "main_dlg";

  : boxed_column {
    label = "Status";

    : text {
      key = "status_text";
      value = "Odoo: Not Connected";
    }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Connection";

    : button {
      label = "Test Odoo Connection";
      key = "btn_connect";
      width = 36;
      fixed_width = true;
    }

    : button {
      label = "Connection Settings...";
      key = "btn_config";
      width = 36;
      fixed_width = true;
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Main Functions";

    : button {
      label = "Set Parameters from Odoo";
      key = "btn_set_params";
      width = 36;
      fixed_width = true;
    }

    : button {
      label = "Push to BOQ";
      key = "btn_push_boq";
      width = 36;
      fixed_width = true;
    }

    : button {
      label = "Transfer BOQ to PR";
      key = "btn_create_pr";
      width = 36;
      fixed_width = true;
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Tools";

    : button {
      label = "Clear Current Layout Table IDs";
      key = "btn_clear_current";
      width = 36;
      fixed_width = true;
    }

    : button {
      label = "Clear ALL Layout Table IDs";
      key = "btn_clear_all";
      width = 36;
      fixed_width = true;
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  cancel_button;
}
