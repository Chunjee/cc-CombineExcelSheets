#NoTrayIcon
#SingleInstance, force
SetBatchLines, -1

#Include node_modules
#Include biga.ahk\export.ahk
#Include json.ahk\export.ahk
#Include transformStringVars.ahk\export.ahk


; NOTES:
; grab col F, K
; ignore line when col G is UR97, UR93, ""
; ignore line when tracking number is blank
; "yeah the only things in the outfile that will ever matter to me will be cells A3-B104"


; class instances
A := new biga()
excel1 := ComObjCreate("Excel.Application") ;writer
excel2 := ComObjCreate("Excel.Application") ;reader
; ~~~ variables ~~~
FileRead, outputVar, % A_ScriptDir "\settings.json"
settings := JSON.parse(outputVar)
outfileWriteIndex := 1

;/--\--/--\--/--\--/--\
; MAIN
;\--/--\--/--\--/--\--/

; create a new excel file to write everything to
excel1.Workbooks.Add
; make Excel Application visible
excel1.Visible := true 
; label, size, and color columns here
excel1.Columns("A").ColumnWidth := 16
excel1.Range("A1").Value := "DIST_ORDER"
excel1.Range("B1").Value := "TRACKING_NBR"
excel1.Range("C1").Value := "UPS_STATUS"
excel1.Range("D1").Value := "SHIP_VIA"
excel1.Range("E1").Value := "COMMENT"


; loop each file in the inputdir location
loop, Files, % transformStringVars(settings.inputdir)
{
    ; skip temporary excel files
    if (InStr(A_LoopFilePath, "~")) {
        continue 
    }
    ; open the excel file to be read
    excel2.Workbooks.Open(A_LoopFilePath)

    ; read each line
    While, (true) {
        ; skip line 1 as this is mostly labels
        if (A_Index == 1) {
            continue
        }

        line := {}
        line.DIST_ORDER := excel2.Range("A" A_Index).Value
        ; check that we are still seeing orders, excel doesn't go on forever
        if (A.size(line.DIST_ORDER) < 5 || ( settings.debuglinelimit != "" && A_Index > settings.debuglinelimit)) {
            break
        }

        ; ok this line has activity, read the other data we're interested in
        line.TRACKING_NBR   := excel2.Range("F" A_Index).Value
        line.SHIP_VIA       := excel2.Range("G" A_Index).Value
        line.UPS_STATUS     := excel2.Range("K" A_Index).Value
        line.COMMENT        := excel2.Range("L" A_Index).Value

        ; FILTER OUT LINES WE DON'T CARE ABOUT HERE
        if (A.indexOf(settings.filtershipvia, line.SHIP_VIA) != -1) {
            ; msgbox, % "skipped: " line.SHIP_VIA
            continue
        }
        if (line.TRACKING_NBR = "") {
            ; msgbox, blank tracking number
            continue
        }

        ; write data out to OUT FILE
        outfileWriteIndex++
        excel1.Range("A" outfileWriteIndex).Value := line.DIST_ORDER
        excel1.Range("B" outfileWriteIndex).Value := line.TRACKING_NBR
        excel1.Range("C" outfileWriteIndex).Value := line.UPS_STATUS
        excel1.Range("D" outfileWriteIndex).Value := line.SHIP_VIA
        excel1.Range("E" outfileWriteIndex).Value := line.COMMENT
    }
}





; save the new excel file
FormatTime, Systemtime, A_Now, yyyyMMddhhmm
savePath := transformStringVars(settings.outfile)
excel1.ActiveWorkbook.SaveAs(savePath)
excel1.ActiveWorkbook.saved := true

; Exit Excel COM objects 
excel1.Quit
excel2.Quit

exitapp, 1



;/--\--/--\--/--\--/--\
; Helper functions
;\--/--\--/--\--/--\--/



;/--\--/--\--/--\--/--\
; functions
;\--/--\--/--\--/--\--/
