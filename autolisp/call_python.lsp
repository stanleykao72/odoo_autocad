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

(defun delay_time (msec)
  (command "_.delay" msec)
)

(defun c:L0()
  (setq
    python "C:/Python/Python310/python.exe"
    pyscript "C:/odoo/odoo_autocad/python/python2com.py"
  )
  (startapp python pyscript)
  
  (delay_time 2000)
  
  (if (vlax-get-or-create-object "Python.ComServer")
    (progn
      (setq python_com (vlax-get-or-create-object "Python.ComServer"))
      (setq server_cfg (vlax-invoke-method python_com 'odoo_connection ))
    )
    (progn
      (princ "L0沒有執行成功，請再執行一次L0.....")
    )
  )
  (princ)
)

(defun c:LREG()
  (setq
    python "C:/Python/Python310/python.exe"
    ;; pyscript_unreg "C:/odoo/odoo_autocad/python/python2com.py  --unregister"
    pyscript_reg "C:/odoo/odoo_autocad/python/python2com.py  --user"
  )
  ;; (startapp python pyscript_unreg)  
  (delay_time 2000)
  (startapp python pyscript_reg)
  (princ)
)