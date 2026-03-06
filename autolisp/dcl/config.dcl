// config.dcl — Connection settings dialog

ob_config : dialog {
  label = "Connection Settings";
  key = "config_dlg";

  : boxed_column {
    label = "Odoo Server";

    : edit_box {
      label = "URL:";
      key = "cfg_url";
      edit_width = 40;
      value = "";
    }

    : edit_box {
      label = "Database:";
      key = "cfg_db";
      edit_width = 40;
      value = "";
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  : boxed_column {
    label = "Credentials";

    : edit_box {
      label = "Username:";
      key = "cfg_user";
      edit_width = 40;
      value = "";
    }

    : edit_box {
      label = "Password/API Key:";
      key = "cfg_pass";
      edit_width = 40;
      value = "";
    }

    : spacer { height = 0.3; }
  }

  : spacer { height = 0.5; }

  ok_cancel;
}
