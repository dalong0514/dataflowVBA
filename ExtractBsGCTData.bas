' refactored at 2021-06-15
' refactored at 2021-08-11
Public Sub ExtractAllBsGCTData()
  Dim tankFileName As String
  Dim heaterFileName As String
  Dim projectFileName As String
  Dim nozzleFileName As String
  Dim supportData As String
  Dim reactorFileName As String
  Dim pressureElementFileName As String
  Dim standardFileName As String
  Dim requirementFileName As String
  Dim otherRequestFileName As String

  tankFileName = "D:\dataflowcad\bsdata\bsGCTTankMainData.txt"
  heaterFileName = "D:\dataflowcad\bsdata\bsGCTHeaterMainData.txt"
  projectFileName = "D:\dataflowcad\bsdata\bsGCTProjectData.txt"
  nozzleFileName = "D:\dataflowcad\bsdata\bsGCTNozzleData.txt"
  supportFileName = "D:\dataflowcad\bsdata\bsGCTSupportData.txt"
  reactorFileName = "D:\dataflowcad\bsdata\bsGCTReactorMainData.txt"
  pressureElementFileName = "D:\dataflowcad\bsdata\bsGCTPressureElementData.txt"
  standardFileName = "D:\dataflowcad\bsdata\bsGCTStandardData.txt"
  requirementFileName = "D:\dataflowcad\bsdata\bsGCTRequirementData.txt"
  otherRequestFileName = "D:\dataflowcad\bsdata\bsGCTOtherRequestData.txt"

  Call ExtractBsGCTProjectDataToCSV(projectFileName, Sheet1.Range("D4:K5"), 1, 8)
  Call ExtractBsGCTDataToCSV(tankFileName, Sheet1.Range("B8:X2000"), 200, 40)
  Call ExtractBsGCTDataToCSV(heaterFileName, Sheet2.Range("B4:X200"), 200, 58)
  Call ExtractBsGCTDataToCSV(nozzleFileName, Sheet3.Range("B4:J2000"), 2000, 11)
  Call ExtractBsGCTDataToCSV(supportFileName, Sheet5.Range("B4:G1000"), 1000, 6)
  Call ExtractBsGCTDataToCSV(reactorFileName, Sheet9.Range("B4:X200"), 200, 57)
  Call ExtractBsGCTDataToCSV(pressureElementFileName, Sheet4.Range("B4:H500"), 500, 7)
  Call ExtractBsGCTDataToCSV(standardFileName, Sheet6.Range("B4:D500"), 500, 3)
  Call ExtractBsGCTDataToCSV(requirementFileName, Sheet7.Range("B4:E500"), 500, 4)
  Call ExtractBsGCTDataToCSV(otherRequestFileName, Sheet8.Range("B4:D500"), 500, 3)

  MsgBox "Extract Sucess!"

End Sub

' 2021-08-11
Sub ExtractBsGCTProjectDataToCSV(gctFileName, range, rowNum, columnNum)
  Dim fso As Object
  Dim myTxt As Object
  Dim csvString As String
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myTxt = fso.CreateTextFile(Filename:=gctFileName, OverWrite:=True)
  Call ExtractRangeNoNullData(range, rowNum, columnNum, myTxt)
  csvString = "," & Sheet1.range("F2") & "," & Sheet1.range("O2") & "," & Sheet1.range("O3") & "," & Sheet1.range("U2") & "," & Sheet1.range("U3") & "," & Sheet1.range("X2") & "," & Sheet1.range("X3") & "," & Sheet1.range("AB2") 
  myTxt.Write csvString
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub

' refactored at 2021-06-11
Sub ExtractBsGCTDataToCSV(gctFileName, range, rowNum, columnNum)
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