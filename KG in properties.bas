Attribute VB_Name = "Kg2"
'' Made by Jakub Dukielski I hope it helps ;)

Dim swApp As Object
Dim part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Sub main()

Set swApp = _
Application.SldWorks

Set part = swApp.ActiveDoc
Dim myModelView As Object
Set myModelView = part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized

boolstatus = part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsLinearFractionDenominator, 0, 0)
boolstatus = part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUnitsLinearFeetAndInchesFormat, 0, False)
boolstatus = part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsDualLinearFractionDenominator, 0, 0)
boolstatus = part.Extension.SetUserPreferenceToggle(swUserPreferenceToggle_e.swUnitsDualLinearFeetAndInchesFormat, 0, False)
boolstatus = part.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, 0, swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms)
bRet = part.ForceRebuild
part.FileSummaryInfo
End Sub
