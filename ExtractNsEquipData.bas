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



' refactored at 2021-05-31
Public Sub ExtractNsEquipData1()
  Call ExtractNsEquipData(Sheet1)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData2()
  Call ExtractNsEquipData(Sheet2)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData3()
  Call ExtractNsEquipData(Sheet3)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData4()
  Call ExtractNsEquipData(Sheet4)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData5()
  Call ExtractNsEquipData(Sheet5)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData6()
  Call ExtractNsEquipData(Sheet6)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData7()
  Call ExtractNsEquipData(Sheet7)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData8()
  Call ExtractNsEquipData(Sheet8)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData9()
  Call ExtractNsEquipData(Sheet9)
End Sub

' refactored at 2021-05-31
Public Sub ExtractNsEquipData(Sheet)
  Dim MyFName As String
  Dim fso As Object
  Dim myTxt As Object
  
  MyFName = "D:\dataflowcad\nsdata\tempEquip.csv"
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Extract the data for special Range
  ' the frist row will be abandoned in autoLisp
  Call ExtractRangeData(Sheet.Range("Q6:AR150"), 500, 28, myTxt)

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