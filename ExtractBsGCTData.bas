' refactored at 2021-04-21
Public Sub ExtractBsGCTDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\bsdata\bsGCT.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Extract main data
  ' the column in range can be wrong. eg [X100]
  Call ExtractColumnsData(Sheet1.Range("B2:X100"), 23, ",Tank", myTxt)
  ' Extract the nozzle data
  Call ExtractColumnsData(Sheet2.Range("B3:H3000"), 7, ",nozzle", myTxt)
  ' Extract the Tank Standard data
  Call ExtractOneColumnData(Sheet3.Range("C3:C12"), ",Tank-Standard", myTxt)
  ' Extract the Tank HeadStyle data
  Call ExtractOneColumnData(Sheet3.Range("D15:D19"), ",Tank-HeadStyle", myTxt)
  ' Extract the Tank HeadMaterial data
  Call ExtractOneColumnData(Sheet3.Range("D20:D24"), ",Tank-HeadMaterial", myTxt)
  ' Extract the Tank Other Request data
  Call ExtractOneColumnData(Sheet3.Range("C27:C40"), ",Tank-OtherRequest", myTxt)

  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
  MsgBox "Extract Sucess!"
End Sub

Sub ExtractOneColumnData(range, dataTypeString, myTxt)
  Dim row As Integer
  row = 1
  Do While range.Cells(row, 1).Value <> ""
    myTxt.Write dataTypeString
    myTxt.Write ","
    myTxt.Write range.Cells(row, 1).Value
    myTxt.Write vbCr
    row = row + 1
  Loop
End Sub

Sub ExtractColumnsData(range, columnNum, dataTypeString, myTxt)
  Dim row As Integer, column As Integer
  row = 1
  column = 1
  Do While range.Cells(row, 1).Value <> ""
    myTxt.Write dataTypeString
    For column = 1 To columnNum
      myTxt.Write ","
      myTxt.Write range.Cells(row, column).Value
    Next column
    myTxt.Write vbCr
    row = row + 1
  Loop
End Sub