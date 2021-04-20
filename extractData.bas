Public Sub ExtractEquipDataToCSV1()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim row As Integer, column As Integer
  Dim extractedData As String
  
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

Public Sub ExtractEquipDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim row As Integer, column As Integer
  
  MyFName = "D:\dataflowcad\nsdata\tempEquip.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  row = 1
  column = 1
  For row = 1 To 150
    For column = 1 To 20
      myTxt.Write ","
      myTxt.Write Range("AE6:BB150").Cells(row, column).Value
    Next column
    myTxt.Write vbCr
  Next row

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub


' 2021-04-19
Public Sub ExtractBsGCTDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim row As Integer, column As Integer
  
  MyFName = "D:\dataflowcad\bsdata\bsGCT.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  row = 1
  column = 1
  Do While Sheet1.Range("B2:U100").Cells(row, 1).Value <> ""
    myTxt.Write ",Tank"
    For column = 1 To 22
      myTxt.Write ","
      myTxt.Write Sheet1.Range("B2:U100").Cells(row, column).Value
      
    Next column
    myTxt.Write vbCr
    row = row + 1
  Loop

  ' Extract the nozzle data
  row = 1
  column = 1
  Do While Sheet2.Range("B3:H3000").Cells(row, 1).Value <> ""
  myTxt.Write ",nozzle"
    For column = 1 To 7
      myTxt.Write ","
      myTxt.Write Sheet2.Range("B3:H3000").Cells(row, column).Value
      
    Next column
    myTxt.Write vbCr
    row = row + 1
  Loop

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub