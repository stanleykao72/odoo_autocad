
(vl-load-com)
;; call python connect_odoo2csv.py
(defun c:O1 ()
  (setq
    python "C:/Python/Python310/python.exe"
    pyscript "C:/odoo/odoo_autocad/python/connect_odoo2csv.py"
   
  )
  (startapp python pyscript)
)
;; call python get_info_from_acad.py
(defun c:O2 ()
  (setq
    python "C:/Python/Python310/python.exe"
    pyscript "C:/odoo/odoo_autocad/python/get_info_from_acad.py"
   
  )
  (startapp python pyscript)
)
;; call python import2odoo_boq.py
(defun c:O3 ()
  (setq
    python "C:/Python/Python310/python.exe"
    pyscript "C:/odoo/odoo_autocad/python/import2odoo_boq.py"
   
  )
  (startapp python pyscript)
)
;; call python import2odoo_pr.py
(defun c:O4 ()
  (setq
    python "C:/Python/Python310/python.exe"
    pyscript "C:/odoo/odoo_autocad/python/import2odoo_pr.py"
   
  )
  (startapp python pyscript)
)
