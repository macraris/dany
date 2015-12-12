Attribute VB_Name = "Modulo1"
Sub registri()
Attribute registri.VB_ProcData.VB_Invoke_Func = " \n14"
'
' registri Macro
'

'
    Range("A1:A21").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:11").Select
    Selection.Delete Shift:=xlUp
    Range("B1:E10").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1:D10").Select
    Selection.Delete Shift:=xlToLeft
    Range("D1:G9").Select
    Selection.Delete Shift:=xlToLeft
    Range("E1:E9").Select
    Selection.Delete Shift:=xlToLeft
    Range("F1:G9").Select
    Selection.Delete Shift:=xlToLeft
    Range("H1:H9").Select
    Selection.Delete Shift:=xlToLeft
    Range("I1:J9").Select
    Selection.Delete Shift:=xlToLeft
    Rows("11:21").Select
    Selection.Delete Shift:=xlUp
    Range("A11:B11").Select
    Selection.Delete Shift:=xlUp
    Range("A13:C35").Select
    Selection.Delete Shift:=xlToLeft
    Range("B13:B34").Select
    Selection.Delete Shift:=xlToLeft
    Range("C13:C22").Select
    Selection.Delete Shift:=xlToLeft
    Range("D13:D23").Select
    Selection.Delete Shift:=xlToLeft
    Range("E13:L22").Select
    Selection.Delete Shift:=xlToLeft
    Range("F13:G22").Select
    Selection.Delete Shift:=xlToLeft
    Range("G13:G22").Select
    Selection.Delete Shift:=xlToLeft
    Range("E24:N24").Select
    Selection.Delete Shift:=xlToLeft
    Range("F24:G24").Select
    Selection.Delete Shift:=xlToLeft
    Range("G24").Select
    Selection.Delete Shift:=xlToLeft
    Rows("27:38").Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("F5:G9").Select
    Selection.NumberFormat = "00000000000"
    Range("C5:C9").Select
    Selection.NumberFormat = "0"
End Sub
Sub registri1()
Attribute registri1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' registri1 Macro
'

'
    ActiveCell.Rows("1:13").EntireRow.Select
    Selection.Delete Shift:=xlUp
  
        
    
  [b1].Select
  Selection.End(xlDown).Select
  
  
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp).Offset(-2, 0)).Select
    Selection.Offset(0, -1).Select
    
'    ActiveCell.Offset(0, -1).Range("A1:A23").Select

    Selection.Delete Shift:=xlToLeft
    Range(Selection, Selection.Offset(0, 3)).Select
    Selection.Offset(0, 1).Select
'
'    ActiveCell.Offset(0, 1).Range("A1:D23").Select
'    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    
    Range("b1").Select
    [b1].Select
  Selection.End(xlDown).Select
  
  
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp).Offset(-2, 0)).Select
    
    Selection.Offset(0, 1).Select
    Range(Selection, Selection.Offset(0, 1)).Delete Shift:=xlToLeft
   
   
   'Da riferimento
'       Range("c1").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlDown)).Select
     Range(Selection, Selection.Offset(0, 3)).Select
    Selection.Offset(0, 1).Select
     Selection.Delete Shift:=xlToLeft
  '" finisce riferimento
  
  'range D
  Range("d1").Select
  Selection.End(xlDown).Select
  
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlUp).Offset(-2, 0)).Select
    Selection.Offset(0, 1).Select
     Selection.Delete Shift:=xlToLeft
 'end rng d
 'Range E
'fin qui funge
Selection.Offset(0, 1).Select

     Range(Selection, Selection.Offset(0, 1)).Select
    Selection.Offset(0, 1).Select
     Selection.Delete Shift:=xlToLeft
 'End rng E
 
 'Range g
 Range("G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Offset(0, 1).Select
     Selection.Delete Shift:=xlToLeft
 'End rng G
 
  'Range i
 Range("h1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.Offset(0, 1)).Select
    Selection.Offset(0, 1).Select
     Selection.Delete Shift:=xlToLeft
 'End rng i
  
   
 
 
 
    ActiveCell.Offset(0, -8).Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
   Selection.End(xlDown).Select
   Range(Selection, Selection.Offset(12, 0)).Select
   Selection.EntireRow.Delete Shift:=xlUp
   
  
  Range("D1").Select

    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
  Dim ci, endLine As String
  ci = ActiveCell.Address
  
    
    Cells.Find(What:="società 0221", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
endLine = ActiveCell.Address

Range(ci & ":" & endLine).Select
Range(Selection, Selection.Offset(0, -2)).Select
Selection.Offset(0, -1).Delete Shift:=xlToLeft

Range(ci & ":" & endLine).Offset(0, -2).Select
    Selection.Delete Shift:=xlToLeft
    
    Selection.Offset(0, 1).Delete Shift:=xlToLeft
    
  'conto iva
   Selection.Offset(0, 1).Delete Shift:=xlToLeft

'significato
 Selection.Offset(0, 1).Select
 Range(Selection, Selection.Offset(0, 8)).Delete Shift:=xlToLeft
 
 Range(ci & ":" & endLine).Offset(0, 2).Select
 Range(Selection, Selection.Offset(0, 1)).Delete Shift:=xlToLeft
 
 'colonna g
Range(ci & ":" & endLine).Offset(0, 3).Delete Shift:=xlToLeft
 
Range(endLine).Select
    Selection.End(xlDown).Select
    
    
    ActiveCell.Range("A1:A25").Select
    Selection.EntireRow.Delete
    
    ActiveCell.Columns("A:O").EntireColumn.Select
    ActiveCell.Columns("A:O").EntireColumn.EntireColumn.AutoFit
  
      [a1].Select
    ActiveWorkbook.Save
End Sub
