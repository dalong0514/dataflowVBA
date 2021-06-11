' refactored at 2021-06-11
Public Sub ExtractBsGCTDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\bsdata\bsGCT.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  Call SetTankInspectRate(Sheet1.Range("R7:AB150"), 9)
  Call SetTankInspectRate(Sheet2.Range("AK5:AK150"), 7)

  ' Extract main data
  ' the column in range could be wrong, still ok. eg [X100]
  Call ExtractColumnsData(Sheet1.Range("B6:X150"), 34, ",Tank", myTxt)
  Call ExtractOneRowData(Sheet1.Range("B5:X5"), ",Tank-MainKeys,BSGCT_TYPE", myTxt)
  ' because heat data need not to delete the first row, B5 not B4
  Call ExtractColumnsData(Sheet2.Range("B5:X100"), 52, ",Heater", myTxt)
  Call ExtractOneRowData(Sheet2.Range("B3:X3"), ",Heater-MainKeys,BSGCT_TYPE", myTxt)

  ' Extract the nozzle data
  Call ExtractOneRowData(Sheet3.Range("B2:H2"), ",NozzleKeys", myTxt)
  Call ExtractColumnsData(Sheet3.Range("B4:H3000"), 7, ",Nozzle", myTxt)

  ' Extract the Tank PressureElement data
  Call ExtractOneRowData(Sheet4.Range("B3:F3"), ",Tank-PressureElementKeys", myTxt)
  Call ExtractColumnsData(Sheet4.Range("B5:F12"), 5, ",Tank-PressureElement", myTxt)
  ' Extract the Heater PressureElement data
  Call ExtractColumnsData(Sheet4.Range("B16:F29"), 5, ",Heater-PressureElement", myTxt)

  ' Extract the support data
  Call ExtractOneRowData(Sheet5.Range("B2:G2"), ",SupportKeys", myTxt)
  Call ExtractColumnsData(Sheet5.Range("B4:G1000"), 6, ",Support", myTxt)

  ' Extract the Vertical Tank Standard data
  Call ExtractOneColumnData(Sheet6.Range("C3:C12"), ",Tank-Standard,verticalTank", myTxt)
  ' Extract the Vertical Tank HeadStyle data
  Call ExtractOneColumnData(Sheet6.Range("D15:D19"), ",Tank-HeadStyle,verticalTank", myTxt)
  ' Extract the Vertical Tank HeadMaterial data
  Call ExtractOneColumnData(Sheet6.Range("D20:D24"), ",Tank-HeadMaterial,verticalTank", myTxt)
  ' Extract the Vertical Tank Other Request data
  Call ExtractOneColumnData(Sheet6.Range("C27:C40"), ",Tank-OtherRequest,verticalTank", myTxt)

  ' Extract the Horizontal Tank Standard data
  Call ExtractOneColumnData(Sheet7.Range("C3:C12"), ",Tank-Standard,horizontalTank", myTxt)
  ' Extract the Horizontal Tank HeadStyle data
  Call ExtractOneColumnData(Sheet7.Range("D15:D19"), ",Tank-HeadStyle,horizontalTank", myTxt)
  ' Extract the Horizontal Tank HeadMaterial data
  Call ExtractOneColumnData(Sheet7.Range("D20:D24"), ",Tank-HeadMaterial,horizontalTank", myTxt)
  ' Extract the Horizontal Tank Other Request data
  Call ExtractOneColumnData(Sheet7.Range("C27:C40"), ",Tank-OtherRequest,horizontalTank", myTxt)

  ' Extract the Heater Standard data
  Call ExtractOneColumnData(Sheet8.Range("C3:C12"), ",Heater-Standard,Heater", myTxt)
  ' Extract the Heater HeadStyle data
  Call ExtractOneColumnData(Sheet8.Range("D15:D19"), ",Heater-HeadStyle,Heater", myTxt)
  ' Extract the Heater HeadMaterial data
  Call ExtractOneColumnData(Sheet8.Range("D20:D24"), ",Heater-HeadMaterial,Heater", myTxt)
  ' Extract the Heater Other Request data
  Call ExtractOneColumnData(Sheet8.Range("C27:C40"), ",Heater-OtherRequest,Heater", myTxt)

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

Sub SetTankInspectRate(range, columnNum)
  Dim row As Integer
  row = 1
  ' weld_joint is the frist column of the range
  Do While range.Cells(row, 1).Value <> ""
    ' barrel inspect_rate is the 7th column of the range
    Select Case True
      Case (range.Cells(row, 1).Value like "0.85/*") 
        range.Cells(row, columnNum).Value = "20%"
      Case (range.Cells(row, 1).Value like "1.0/*") 
        range.Cells(row, columnNum).Value = "100%"
      Case else 
        range.Cells(row, columnNum).Value = "/"
    End Select
    ' head inspect_rate is the 7th column of the range
    Select Case True
      Case (range.Cells(row, 1).Value like "*/0.85") 
        range.Cells(row, columnNum+1).Value = "20%"
      Case (range.Cells(row, 1).Value like "*/1.0") 
        range.Cells(row, columnNum+1).Value = "100%"
      Case else 
        range.Cells(row, columnNum+1).Value = "/"
    End Select
    row = row + 1
  Loop
End Sub