Public Sub ExtractEquipDataToCSV()
  Dim fso As Object
  Dim myTxt As Object
  Dim MyFName As String
  Dim sht As Worksheet
  Dim extractedData As String
  
  MyFName = "D:\dataflowcad\NsTempData\equip.txt" 
  
  Set fso = CreateObject("Scripting.FileSystemObject") 
  Set myTxt = fso.CreateTextFile(Filename:=MyFName, OverWrite:=True) 

  extractedData = Range("K8").Value
  'Range("A4").Value = 200
  myTxt.Write extractedData
  
  myTxt.Close
  Set myTxt = Nothing
  Set fso = Nothing
End Sub
