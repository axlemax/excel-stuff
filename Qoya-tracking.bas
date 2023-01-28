Attribute VB_Name = "Module1"
Dim SKULocation As Range
Dim LastHistoryEntry As String
Dim LastHistoryPrice As String
Dim LastPrice As String
Dim InventorySheet As String
Dim TrackingSheet As String


Function FindValueInColumn(MyColumn As Range, MyValue As Variant) As String
    On Error GoTo HandleError
    
    With MyColumn
        FindValueInColumn = .Find(What:=MyValue, After:=.Cells(.Cells.Count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
Done:
    Exit Function
    
HandleError:
    MsgBox "SKU Not found"
    End

End Function


Function FindSKU(MyValue As Variant) As String
    FindSKU = FindValueInColumn(Worksheets(InventorySheet).Range("C:C"), MyValue)
End Function


Sub SKUInput()
Attribute SKUInput.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim SKU As String
    InventorySheet = "Inventory"
    TrackingSheet = "Tracking"
    
    ' SKU = InputBox("Enter SKU", "Find SKU (using C Column)")
    SKU = Worksheets(TrackingSheet).Range("E2").Value & Worksheets(TrackingSheet).Range("F2").Value
    
    If Worksheets(TrackingSheet).Range("J2").Value <> SKU Then
    
       Set SKULocation = Range(FindSKU(SKU))
       
       If Worksheets(InventorySheet).Cells(SKULocation.Row, 1).Value = "0" Then
            MsgBox "SKU Already sold"
            End
       End If
       
       Dim PaymentMethod As String
       PaymentMethod = Worksheets(TrackingSheet).Range("D2").Value
       
       Dim DateString As String
       DateString = Worksheets(TrackingSheet).Range("C2").Value
       
       LastPrice = Worksheets(InventorySheet).Cells(SKULocation.Row, 5).Value
       
       Worksheets(InventorySheet).Cells(SKULocation.Row, 1).Value = "0"
       Worksheets(InventorySheet).Cells(SKULocation.Row, 7).Value = PaymentMethod
       Worksheets(InventorySheet).Cells(SKULocation.Row, 10).Value = DateString
    
       'Add history
    
       LastHistoryEntry = Worksheets(TrackingSheet).Range("J4").Value
       Worksheets(TrackingSheet).Range("J4").Value = Worksheets(TrackingSheet).Range("J3").Value
       Worksheets(TrackingSheet).Range("J3").Value = Worksheets(TrackingSheet).Range("J2").Value
'       Call Rotate(TrackingSheet, "J2", "J4")
       Worksheets(TrackingSheet).Range("J2").Value = SKU
       
       LastHistoryPrice = Worksheets(TrackingSheet).Range("K4").Value
       Worksheets(TrackingSheet).Range("K4").Value = Worksheets(TrackingSheet).Range("K3").Value
       Worksheets(TrackingSheet).Range("K3").Value = Worksheets(TrackingSheet).Range("K2").Value
       Worksheets(TrackingSheet).Range("K2").Value = LastPrice
       
       ' Enable undo
       Application.OnUndo "Undo sale", "ApplyUndo"
    End If
    
End Sub


Private Sub ApplyUndo()
    Dim ToUndo As String
    Dim ToUndoLocation As Range
    ToUndo = Worksheets(TrackingSheet).Range("J2").Value
    Set ToUndoLocation = Range(FindSKU(ToUndo))

    Worksheets(InventorySheet).Cells(ToUndoLocation.Row, 1).Value = "1"
    Worksheets(InventorySheet).Cells(ToUndoLocation.Row, 7).ClearContents
    Worksheets(InventorySheet).Cells(ToUndoLocation.Row, 10).ClearContents
    
    Worksheets(TrackingSheet).Range("J2").Value = Worksheets(TrackingSheet).Range("J3").Value
    Worksheets(TrackingSheet).Range("J3").Value = Worksheets(TrackingSheet).Range("J4").Value
    Worksheets(TrackingSheet).Range("J4").Value = LastHistoryEntry
    
    Worksheets(TrackingSheet).Range("K2").Value = Worksheets(TrackingSheet).Range("K3").Value
    Worksheets(TrackingSheet).Range("K3").Value = Worksheets(TrackingSheet).Range("K4").Value
    Worksheets(TrackingSheet).Range("K4").Value = LastHistoryPrice
End Sub


Private Sub Rotate(Sheet As String, StartCell As String, EndCell As String)
' TODO Go backwards from bottom up
    Dim Column As String
    Column = Col_Letter(Worksheets(Sheet).Range(StartCell).Column)
    Dim CurrentRow, EndRow As String
    CurrentRow = Worksheets(Sheet).Range(StartCell).Row
    EndRow = Worksheets(Sheet).Range(EndCell).Row

    Dim ToAdd As Integer
    ToAdd = IIf(CurrentRow < EndRow, 1, -1)
    ' MsgBox ("Column: " & Column & " Row: " & CurrentRow)

    While CurrentRow <= EndRow
        Worksheets(Sheet).Range(Column & CurrentRow).Value = Worksheets(Sheet).Range(Column & CurrentRow + ToAdd).Value
        CurrentRow = CurrentRow + ToAdd
    Wend

End Sub

Private Function Col_Letter(lngCol As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
