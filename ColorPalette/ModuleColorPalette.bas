Attribute VB_Name = "ModuleColorPalette"
Option Explicit
Public OptionColorValue As Boolean ' keeps the last status of frmColorPalette.OptionColor(1).Value

Public Sub ShowPalette(SControl As Control)
    
    frmColorPalette.IniPalette SControl
    
' "SControl" is the control where
' we want apply the color. Can be any control
' that has "ForeColor" and "BackColor" properties

    frmColorPalette.Show
    
End Sub
