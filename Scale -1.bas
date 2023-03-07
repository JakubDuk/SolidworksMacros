Attribute VB_Name = "scala_record1"

Dim swApp                   As SldWorks.SldWorks
Dim swDraw                  As SldWorks.DrawingDoc
Dim swSheet                 As SldWorks.Sheet
Dim vSheetProperties        As Variant

Dim scale1 As Double
Dim scale2 As Double
Dim boolstatus As Boolean




Sub main()

    Set swApp = Application.SldWorks
    Set swDraw = swApp.ActiveDoc
    Set swSheet = swDraw.GetCurrentSheet

vSheetProperties = swSheet.GetProperties

scale1 = vSheetProperties(2)
scale2 = vSheetProperties(3)

 If scale2 = 1 Then
    boolstatus = swSheet.SetScale(scale1 + 1, scale2, True, False)
Else

boolstatus = swSheet.SetScale(scale1, scale2 - 1, True, False)

End If
End Sub
