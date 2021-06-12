' refactored at 2021-06-12
Public Sub ExtractAllBsGCTData()
  Dim tankFileName As String
  Dim heaterFileName As String
  Dim projectFileName As String
  Dim nozzleFileName As String
  tankFileName = "D:\dataflowcad\bsdata\bsGCTTankMainData.txt"
  heaterFileName = "D:\dataflowcad\bsdata\bsGCTHeaterMainData.txt"
  projectFileName = "D:\dataflowcad\bsdata\bsGCTProjectData.txt"
  nozzleFileName = "D:\dataflowcad\bsdata\bsGCTNozzleData.txt"

  Call ExtractBsGCTOtherDataToCSV()
  Call ExtractBsGCTDataToCSV(projectFileName, Sheet1.Range("D4:K5"), 2, 8)
  Call ExtractBsGCTDataToCSV(tankFileName, Sheet1.Range("B8:X150"), 150, 36)
  Call ExtractBsGCTDataToCSV(heaterFileName, Sheet2.Range("B4:X150"), 150, 54)
  Call ExtractBsGCTDataToCSV(nozzleFileName, Sheet3.Range("B4:J150"), 2000, 9)
End Sub

' refactored at 2021-06-11
Public Sub ExtractBsGCTDataToCSV(gctFileName, range, rowNum, columnNum)
  Dim fso As Object
  Dim myTxt As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=gctFileName, OverWrite:=True)
  Call ExtractRangeNoNullData(range, rowNum, columnNum, myTxt)
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

Sub ExtractRangeNoNullData(range, rowNum, columnNum, myTxt)
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

' refactored at 2021-06-11
Public Sub ExtractBsGCTOtherDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  
  MyFName = "D:\dataflowcad\bsdata\bsGCT.csv"
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True)

  ' Call SetTankInspectRate(Sheet1.Range("R7:AB150"), 9)
  ' Call SetTankInspectRate(Sheet2.Range("AK5:AK150"), 7)

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