Public Sub ExtractNsEquipDataV1()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim row As Integer, column As Integer
  
  MyFName = "D:\dataflowcad\nsdata\tempEquip.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  row = 1
  column = 1
  Do While Range("B2:U100").Cells(row, 1).Value <> ""
    For column = 1 To 20
      myTxt.Write ","
      myTxt.Write Range("B2:U100").Cells(row, column).Value
    Next column
    myTxt.Write vbCr
    row = row + 1
  Loop

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub

' refactored at 2021-04-25
Public Sub ExtractNsEquipData()
  Dim MyFName As String
  Dim fso As Object
  Dim myTxt As Object
  
  MyFName = "D:\dataflowcad\nsdata\tempEquip2.csv"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Extract the data for special Range
  ' the frist row will be abandoned in autoLisp
  Call ExtractRangeData(Sheet1.Range("Q6:AR150"), 200, 28, myTxt)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub

Sub ExtractRangeData(range, rowNum, columnNum, myTxt)
  Dim row As Integer, column As Integer
  row = 1
  column = 1
  For row = 1 To rowNum
    ' Only extract the row that have the data
    if range.Cells(row, 1).Value <> "" Then 
      For column = 1 To columnNum
        myTxt.Write ","
        myTxt.Write range.Cells(row, column).Value
      Next column
      myTxt.Write vbCr
    End if
  Next row
End Sub