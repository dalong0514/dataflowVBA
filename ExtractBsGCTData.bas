' refactored at 2021-05-06
Public Sub ExtractBsGCTDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\bsdata\bsGCT.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Extract main data
  ' the column in range can be wrong. eg [X100]
  Call ExtractColumnsData(Sheet1.Range("B4:X100"), 31, ",Tank", myTxt)
  Call ExtractOneRowData(Sheet1.Range("B3:X3"), ",Tank-MainKeys,BSGCT_TYPE", myTxt)
  ' Extract the nozzle data
  Call ExtractOneRowData(Sheet2.Range("B2:H2"), ",NozzleKeys", myTxt)
  Call ExtractColumnsData(Sheet2.Range("B4:H3000"), 7, ",Nozzle", myTxt)
  ' Extract the Tank PressureElement data
  Call ExtractOneRowData(Sheet3.Range("B3:F3"), ",Tank-PressureElementKeys", myTxt)
  Call ExtractColumnsData(Sheet3.Range("B5:F10"), 5, ",Tank-PressureElement", myTxt)
  ' Extract the Tank Standard data
  Call ExtractOneColumnData(Sheet6.Range("C3:C12"), ",Tank-Standard", myTxt)
  ' Extract the Tank HeadStyle data
  Call ExtractOneColumnData(Sheet6.Range("D15:D19"), ",Tank-HeadStyle", myTxt)
  ' Extract the Tank HeadMaterial data
  Call ExtractOneColumnData(Sheet6.Range("D20:D24"), ",Tank-HeadMaterial", myTxt)
  ' Extract the Tank Other Request data
  Call ExtractOneColumnData(Sheet6.Range("C27:C40"), ",Tank-OtherRequest", myTxt)

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

Sub ExtractOneRowData(range, dataTypeString, myTxt)
  Dim column As Integer
  column = 1
  myTxt.Write dataTypeString
  Do While range.Cells(1, column).Value <> ""
    myTxt.Write ","
    myTxt.Write range.Cells(1, column).Value
    column = column + 1
  Loop
  myTxt.Write vbCr
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