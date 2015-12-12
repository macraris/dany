## Macro Registri Dany
****************
``` vb
Sub registri()
Attribute registri.VB_ProcData.VB_Invoke_Func = " \n14"
'
'***********************************************************************************
'~~Questa semplice macro esegue operazioni basilari quali selezioni dinamici di
'~~intervallo di dati,eliminazione di intervallo di dati (range) vuoti,
'~~qualche formattazione qua e la con uso di condizionale if...then , cicli For...Next
'***************************************************************************************

Dim infos As Variant
    infos = MsgBox("Elaborazione Registri..." & vbNewLine & vbNewLine & _
    "Qui per sbaglio -->  Click 'NO'", _
                    vbYesNo + vbInformation + vbDefaultButton2, "Macr@ris Registri")
                    
If infos = vbNo Then
    Exit Sub
    End If
'
''''
Application.StatusBar = "Goditi un Caffe' mentre lavoro per Te...."

Application.ScreenUpdating = False
On Error GoTo ErrorHandler

'#A
Rows("1:13").Delete shift:=xlUp
       
 '#B
 Dim lCel As Integer 'definizione variabile per avere numero ultima riga di un range
    
   lCel = [B3].End(xlDown).Row
   
    Range("A1:A" & lCel & ",C1:F" & lCel & ",H1:I" & lCel & ",K1:N" & lCel & ",P1:P" & lCel & _
          ",R1:S" & lCel & ",V1:V" & lCel & ",X1:Y" & lCel) _
    .Delete shift:=xlToLeft
''''

'#C  ''area secondo quadrante dati da non considerare quindi cancellare
lCel = Cells.Find(What:="INCASSO IVA", After:=[A1], LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Row 'activate
 
Range(Cells(lCel, "A"), Cells(lCel, "A").Offset(12, 0)).EntireRow.Delete shift:=xlUp
'''

'#D  'Quadrante Dati di subtotale con formattazione numeri e dati negativi in positivi
   
  Dim pCel As Integer ''definizione della variabile che conterra' il numero posizione prima riga del range
    pCel = Cells.Find(What:="ContoIVA", After:=[A1], LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Row
    lCel = Cells.Find(What:="Società 0221", After:=[A1], LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Row
 ''cancellazione intervalli non necessari con spostamento celle a sinistra
    
 Range("A" & pCel & ":A" & lCel & ",B" & pCel & ":C" & lCel & ",E" & pCel & ":E" & lCel & _
        ",G" & pCel & ":G" & lCel & ",I" & pCel & ":I" & lCel & ",K" & pCel & ":R" & lCel & _
        ",T" & pCel & ":U" & lCel & ",W" & pCel & ":W" & lCel) _
  .Delete shift:=xlToLeft

''''''
'#E  Transformazione e formattazione numeri negativi in positivi
Range(Cells(pCel, "E"), Cells(lCel, "G")).Select
    Dim rng As Range 'Definisce variabile per ciclo For Each...next
    
    For Each rng In Selection
        If IsNumeric(rng) And Not IsEmpty(rng) Then
        rng.Value = rng * -1
        rng.NumberFormat = "#,##0.00"
    End If
    
    Next
     
  '#F  'Cancella ultime righe in quanto non necessarie
  
  pCel = lCel + 1
  lCel = [A1].SpecialCells(xlCellTypeLastCell).Row
Range(Cells(pCel, "A"), Cells(lCel, "A")).EntireRow.Delete shift:=xlUp

Cells.EntireColumn.AutoFit
[A1].Select

 Application.StatusBar = ""
 Application.ScreenUpdating = True
 Exit Sub
 
ErrorHandler:
MsgBox "Interruzione Macro Causa Errore in Registri" & vbNewLine & "Contattare Macr@ris" & vbNewLine & _
    vbCrLf & "Error number:  # " & Err.Number & vbNewLine & _
      "Description:==> " & Err.Description, vbCritical, "Macr@ris \Error Macro"
 
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = ""

End Sub
```
