; Credit to Denon Deterding
; for VPverts script on autoCAD forum

(defun c:aq (/ vp cvp a b c d) 
  (if (not (atoms-family 1 '("C:ALIGNSPACE"))) 
    (if (eq "Uh Oh:" (load "aspace.lsp" "Uh Oh:")) 
      (progn (prompt " ALIGNSPACE lisp not found.") (exit))
    )
  )
  ;works only on paper space
  (if (and (= 0 (getvar 'TILEMODE)) (= 1 (getvar 'CVPORT))) 
    (progn 
      (setq vp (car (entsel "\nSelect Viewport: ")))
      (setq vertices (VPverts vp))
      (setq b (nth 2 vertices)
            a (nth 3 vertices)
      )
      (setq cvp (cdr (assoc 69 (entget vp))))
      (command "_.MSPACE")
      (setvar 'CVPORT cvp)
      (setq mr (car (entsel "\nSelect ModelRectangle: ")))
      (setq verticesmr (VPverts mr))
      (setq c (nth 0 verticesmr)
            d (nth 1 verticesmr)
      )
      (command "_.PSPACE")
      (if (and vp a b c d) 
        (progn 
          (alignspace c d a b)
          (command "_.PSPACE")
          (prompt "\nSuccess!")
        )
        (prompt "\n...something went wrong.")
      )
    )
    (prompt "\n...you're not in Paper Space.")
  )
  (princ)
)


(defun VPverts (vp / vertex vList tmp i) 
  ;vp - viewport ename
  ;returns - list of vertices or nil
  (setq vp    (entget vp)
        vList '()
  )
  (cond 
    ((member '(0 . "POLYLINE") vp)
     (setq vertex vp)
     (while (setq vertex (entnext (cdr (assoc -1 vertex)))) 
       (setq vertex (entget vertex))
       (if (member '(0 . "VERTEX") vertex) 
         (setq vList (cons (cdr (assoc 10 vertex)) vList))
       )
     )
    ) ;jika 3dPolyline
    ((member '(0 . "LWPOLYLINE") vp)
     (foreach i vp 
       (if (= 10 (car i)) 
         (setq vList (cons (cdr i) vList))
       ) ;if
     ) ;foreach
    ) ;jika LWPOLYLINE 2d
    ((member '(0 . "VIEWPORT") vp)
     (setq vpNum (cdr (assoc 69 vp)))
     (setq vertex (car (cdr (assoc vpNum (vports)))))
     (setq tmp (cadr (cdr (assoc vpNum (vports)))))
     (setq vList (cons vertex vList)
           vList (cons (list (car tmp) (cadr vertex)) vList)
     )
     (setq vList (cons tmp vList)
           vList (cons (list (car vertex) (cadr tmp)) vList)
     )
    ) ;cond VIEWPORT
    (t (prompt "\n...unable to extract Viewport vertices."))
  ) ;cond
  (if (< 0 (length vList)) vList nil)
);defun
