Attribute VB_Name = "WdanyFisc"

Private Sub registri()
Attribute registri.VB_ProcData.VB_Invoke_Func = " \n14"
'
'***********************************************************************************
'~~Questa semplice macro esegue operazioni basilari quali selezioni dinamici di     *
'~~intervallo di dati,eliminazione di intervallo di dati (range) vuoti,              *
'~~qualche formattazione qua e la con uso di condizionale if...then , cicli For...Next *
'****************************************************************************************

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

 Exit Sub
 
ErrorHandler:
MsgBox "Interruzione Macro Causa Errore in Registri" & vbNewLine & "Contattare Macr@ris" & vbNewLine & _
    vbCrLf & "Error number:  # " & Err.Number & vbNewLine & _
      "Description:==> " & Err.Description, vbCritical, "Macr@ris \Error Macro"
 
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = ""

End Sub


Sub loopFile()
   ' Sets up the variable "MyFile" to be each file in the directory
      ' This example looks for all the files that have an .xls extension.
      ' This can be changed to whatever extension is needed. Also, this
      ' macro searches the current directory. This can be changed to any
      ' directory.

Dim infos As Variant
    infos = MsgBox("Elaborazione Registri..." & vbNewLine & vbNewLine & _
    "Qui per sbaglio -->  Click 'NO'", _
                    vbYesNo + vbInformation + vbDefaultButton2, "Macr@ris Registri")

If infos = vbNo Then
    Exit Sub
    End If

'''
Application.StatusBar = "Goditi un Caffe' mentre lavoro per Te...."

Application.ScreenUpdating = False

On Error GoTo ErrorHandler

      ' Test for Windows or Macintosh platform. Make the directory request.
      Dim myFile As String, Sep As String
      
      Sep = Application.PathSeparator

      If Sep = "\" Then
         ' Windows platform search syntax.
         'MyFile = Dir(CurDir() & Sep & "*.xls")
         stPath = "C:\Users\kwemarit\Desktop\REG"
            myFile = Dir(stPath & Sep & "*.xls*")
      Else

         ' Macintosh platform search syntax.
         myFile = Dir("", MacID("XLS5"))
      End If

    ''''@''''''''''''''''''''
     '''' CHECK IF FOLDER IS EMPTY
    If myFile = "" Then
            MsgBox "Nessun File presente in " & stPath & vbNewLine & _
            "Cartella --> 'REG'" & vbNewLine & _
             vbNewLine & "Caricare la cartella e rilanciare Macro." _
        , vbExclamation, "Errore!RegistriLoopFile. By @ris"
      Exit Sub
      End If
      
      ' Starts the loop, which will continue until there are no more files
      ' found.

'@ Add New Workbook
'''---@ Definizione di Variabili per la sub routine
    Dim newBook As Workbook
    Dim StFile As String
    Dim stPathDest As String
 Dim sht As Worksheet, shtCount As Integer
    
        stPathDest = "C:\Users\kwemarit\Desktop\REG\elaborato\"
       StFile = stPathDest & "regLavorato_" & Format(Now, "dd-mmm-yy hh-mm-ss") & ".xlsx"
  
  Application.DisplayAlerts = False
  
    Set newBook = Workbooks.Add
    With newBook
        .Title = "Registri"
        .Subject = "Fiscalita'"
        shtCount = .Sheets.Count
        .SaveAs StFile
    End With
    Application.DisplayAlerts = True
    
      
      Do While myFile <> "" ' And InStr(1, myFile, "ATT")


Workbooks.Open Filename:=stPath & Sep & myFile
  
''-------@@@
''invocazione del modulo registri
registri

''-------###

On Error GoTo ErrorHandler
 Set sht = ActiveSheet
sht.Move After:=newBook.Sheets(shtCount)
                           shtCount = shtCount + 1
         myFile = Dir()
         
      Loop

Application.DisplayAlerts = False
                        With newBook
                            .Sheets("sheet1").Delete
                            .Sheets("sheet2").Delete
                            .SaveAs StFile
                       End With
                Application.DisplayAlerts = True

'@ DELETING ALL FILES IN CURRENT FOLDER

Dim MyFile_KillAll As String
    MyFile_KillAll = Dir(stPath & Sep & "*.*")

Do While MyFile_KillAll <> ""
    Kill stPath & Sep & MyFile_KillAll
    MyFile_KillAll = Dir()
Loop


Application.StatusBar = ""
 Application.ScreenUpdating = True
Exit Sub

ErrorHandler:
MsgBox "Interruzione Macro Causa Errore in loopFile" & vbNewLine & "Contattare Macr@ris" & vbNewLine & _
    vbCrLf & "Error number:  # " & Err.Number & vbNewLine & _
      "Description:==> " & Err.Description, vbCritical, "Macr@ris \Error Macro"
 
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = ""

End Sub
