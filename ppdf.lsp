(vl-load-com)
(defun c:ppdf (/ dwg file hnd i len llpt lst mn mx ss tab urpt)
(setq num (getint "1 Untuk Long 2 Untuk Cross:"))
	(if (= num 1)
	(
    (if (setq ss (ssget "_X" '((0 . "INSERT") (2 . "kop"))))
        (progn
            (repeat (setq i (sslength ss))
                (setq hnd (ssname ss (setq i (1- i)))
                      tab (cdr (assoc 410 (entget hnd)))
                      lst (cons (cons tab hnd) lst)
                )
            )
            (setq lst (vl-sort lst '(lambda (x y) (> (car x) (car y)))))
            (setq i 0)
            (foreach x lst
                (setq file (strcat (getvar 'DWGPREFIX)
                                   (substr (setq dwg (getvar 'DWGNAME)) 1 (- (strlen dwg) 4))
                                   "-"
                                   (itoa (setq i (1+ i)))
                                   ".pdf"
                           )
                )
                (if (findfile file)
                    (vl-file-delete file)
                )
                (vla-getboundingbox (vlax-ename->vla-object (cdr x)) 'mn 'mx)
                (setq llpt (vlax-safearray->list mn)
                      urpt (vlax-safearray->list mx)
                      len  (distance llpt (list (car urpt) (cadr llpt)))
                )
                (command "-plot"
                         "yes"
                         (car x)
                         "DWG TO PDF.PC3"
                         "ISO full bleed A3 (420.00 x 297.00 MM)"
                         "Millimeters"
                         "Landscape"
                         "No"
                         "Window"
                         llpt
                         urpt
                         "1=4"
                         "Center"
                         "yes"
                         "hitam A3.ctb"
                         "yes"
                         ""
                )
                (if (/= (car x) "Model")
                    (command "No" "No" file "no" "Yes")
                    (command
                        file
                        "no"
                        "Yes"
                    )
                )
            )
        )
    )
    (princ)
)
)
	(if (= num 2)
		(
		(if (setq ss (ssget "_X" '((0 . "INSERT") (2 . "kop"))))
			(progn
				(repeat (setq i (sslength ss))
					(setq hnd (ssname ss (setq i (1- i)))
						  tab (cdr (assoc 410 (entget hnd)))
						  lst (cons (cons tab hnd) lst)
					)
				)
				(setq lst (vl-sort lst '(lambda (x y) (> (car x) (car y)))))
				(setq i 0)
				(foreach x lst
					(setq file (strcat (getvar 'DWGPREFIX)
									   (substr (setq dwg (getvar 'DWGNAME)) 1 (- (strlen dwg) 4))
									   "-"
									   (itoa (setq i (1+ i)))
									   ".pdf"
							   )
					)
					(if (findfile file)
						(vl-file-delete file)
					)
					(vla-getboundingbox (vlax-ename->vla-object (cdr x)) 'mn 'mx)
					(setq llpt (vlax-safearray->list mn)
						  urpt (vlax-safearray->list mx)
						  len  (distance llpt (list (car urpt) (cadr llpt)))
					)
					(command "-plot"
							 "yes"
							 (car x)
							 "DWG TO PDF.PC3"
							 "ISO full bleed A3 (420.00 x 297.00 MM)"
							 "Millimeters"
							 "Landscape"
							 "No"
							 "Window"
							 llpt
							 urpt
							 "1=0.2"
							 "Center"
							 "yes"
							 "hitam A3.ctb"
							 "yes"
							 ""
					)
					(if (/= (car x) "Model")
						(command "No" "No" file "no" "Yes")
						(command
							file
							"no"
							"Yes"
						)
					)
				)
			)
		)
		(princ)
	)
	)
)