// result.dcl — Result display dialog (title + message + OK)

ob_result : dialog {
  key = "result_dlg";

  : text {
    key = "result_title";
    value = "Result";
  }

  : spacer { height = 0.5; }

  : paragraph {
    : text_part {
      key = "result_line1";
      value = "";
    }
    : text_part {
      key = "result_line2";
      value = "";
    }
    : text_part {
      key = "result_line3";
      value = "";
    }
    : text_part {
      key = "result_line4";
      value = "";
    }
    : text_part {
      key = "result_line5";
      value = "";
    }
  }

  : spacer { height = 0.5; }

  ok_only;
}
