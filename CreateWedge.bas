Attribute VB_Name = "CreateWedge"
Sub Ch8_CreateWedge()
    Dim wedgeObj As Acad3DSolid
    Dim center(0 To 2) As Double
    Dim length As Double
    Dim width As Double
    Dim height As Double

    ' Define the wedge
    center(0) = 5#: center(1) = 5#: center(2) = 0
    length = 10#: width = 15#: height = 20#

    ' Create the wedge in model space
    Set wedgeObj = ThisDrawing.ModelSpace. _
 AddWedge(center, length, width, height)

    ' Change the viewing direction of the viewport
    Dim NewDirection(0 To 2) As Double
    NewDirection(0) = -1
    NewDirection(1) = -1
    NewDirection(2) = 1
    ThisDrawing.ActiveViewport.Direction = NewDirection
    ThisDrawing.ActiveViewport = ThisDrawing.ActiveViewport
    ZoomAll
End Sub

