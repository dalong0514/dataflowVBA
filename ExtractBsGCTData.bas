' refactored at 2021-04-21
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
  ' the column in range can be wrong. eg [X100]
  Do While Sheet1.Range("B2:X100").Cells(row, 1).Value <> ""
    myTxt.Write ",Tank"
    For column = 1 To 23
      myTxt.Write ","
      myTxt.Write Sheet1.Range("B2:X100").Cells(row, column).Value
      
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