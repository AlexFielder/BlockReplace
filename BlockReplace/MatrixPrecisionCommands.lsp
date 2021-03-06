﻿
(setq MPECmdsVersion "1.0")
(defun c:atout ( / dn pa padn ss)
(vl-load-com)
(load "attout")
(setq dn (vl-filename-base (getvar "dwgname")))
(setq pa (getvar "dwgprefix"))
(setq padn (strcat pa dn ".txt"))
(setq ss (ssget "_X" (list (cons 0 "INSERT")'(66 . 1))))
(bns_attout padn ss)
);atout
(defun c:atin (/ dn pa padn)
(vl-load-com)
(load "attout")
(setq dn (vl-filename-base (getvar "dwgname")))
(setq pa (getvar "dwgprefix"))
(setq padn (strcat pa dn ".txt"))
(bns_attin padn false)
);atin
(defun c:uatts (/ ss)
(setq ss (ssget "_X" (list (cons 0 "INSERT")'(66 . 1))))
(updateatts ss)
);uatts
(defun c:srevs (/ ss rn rd)
(setq ss (ssget "_X" (list (cons 0 "INSERT")'(66 . 1))))
(setq rn "RN240129")
(setq rd (menucmd "m=$(edtime,$(getvar,DATE),DD-MON-YY)"))
(sortrevs ss rn rd)
);srevs
;;;(defun c:client ( dwgnum sht rev / newdn )
(defun client ( dwgnum sht rev / newdn psht prev DwgName )
(if (not (= (type dwgnum) 'STR))
    (setq dwgnum "DWGNUM")
  );if
  (if (not (= (type sht) 'STR))
    (setq sht "001")
  );if
  (if (not (= (type rev) 'STR))
    (setq rev "REV")
  );if
  (setq old-echo (getvar "CMDECHO"))
  (setvar "cmdecho" 1)
  (setvar "draworderctl" 1)
;;;Step 1. A quick save of the current drawing. (works)
  (command "qsave")
;;;Step 2. A save as to a drawing with file name "CLIENT_export.dwg" in the same directory as the current drawing.
;PAD SHT TO 3 DIGITS:
(setq psht (setanumber "_sht_" sht 3))
;PAD REV TO 3 DIGITS:
(setq prev (setanumber "_iss-" rev 3))
(setq newdn (strcat dwgnum psht prev ".dwg"))
(setq DwgName (strcat (getvar "dwgprefix") newdn))
(if (findfile DwgName);Check file exists
(command "_.saveas" "" DwgName "Y");If file exits then Overwrite
  (command "_.saveas" "" DwgName)
  );if
  (setvar "CMDECHO" old-echo)
   (setvar "DRAWORDERCTL" 3)
);CLIENT
(defun setanumber ( prfx n1 d1 / )
(if (< (strlen (itoa n1)) d1)
(strcat prfx (substr "00000000" 1 (- d1 (strlen (itoa n1))))(itoa
n1))
(strcat prfx (itoa n1))
)
);SETANUMBER
(vl-load-com)
(princ
    (strcat
        "\n:: MatrixPrecisionCommands.lsp | Version "
        MPECmdsVersion
        " | © Matrix Precision "
        (menucmd "m=$(edtime,$(getvar,DATE),YYYY)")
        " www.matrixprecision.com ::"
        "\n:: Type \"atout\" OR \"atin\" OR \"uatts\" OR \"delsel\" OR \"STXT\" to Invoke ::"
    )
)
(princ)
