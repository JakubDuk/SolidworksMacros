Attribute VB_Name = "Jakub_s_Auto_File_NAMESAV"
'' Made by Jakub Dukielski I hope it helps ;)

Option Explicit

Public Enum swDrawingViewTypes_e
    swDrawingSheet = 1
    swDrawingSectionView = 2
    swDrawingDetailView = 3
    swDrawingProjectedView = 4
    swDrawingAuxiliaryView = 5
    swDrawingStandardView = 6
    swDrawingNamedView = 7
    swDrawingRelativeView = 8

End Enum

Sub main1()

Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim swDraw                  As SldWorks.DrawingDoc
Dim swSheet                 As SldWorks.Sheet
Dim swView                  As SldWorks.View
Dim bRet                    As Boolean
Dim RefDoc                  As String
Dim File                    As String
Dim SplitFileName           As Variant
Dim FilePath                As String
Dim PathName                As String

 

 

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

   If (swModel.GetType = swDocDRAWING) Then

    Set swDraw = swModel
    swDraw.ClearSelection2 True
    Set swSheet = swDraw.GetCurrentSheet
    Set swView = swDraw.GetFirstView

While Not swView Is Nothing

    bRet = swDraw.ActivateView(swView.GetName2)

    SplitFileName = Split(swView.GetReferencedModelName, "\")    ''Split modelname
    
    If UBound(SplitFileName) > 0 Then

        RefDoc = SplitFileName(UBound(SplitFileName))
        
        File = Left(RefDoc, InStrRev(RefDoc, "."))
        PathName = swView.GetReferencedModelName
        FilePath = Left(PathName, InStrRev(PathName, "\"))
        
        swModel.SaveAs (FilePath + "Flat " + File + "slddrw")

    End If

    Set swView = swView.GetNextView

Wend

swModel.GraphicsRedraw2

End If


End Sub

