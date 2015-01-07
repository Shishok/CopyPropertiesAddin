Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core


<ComVisible(True)> _
Partial Public Class Addin
    Implements IRibbonExtensibility
    Private _ribbon As Microsoft.Office.Core.IRibbonUI

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ribbonId As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return My.Resources.Ribbon
    End Function

#End Region

#Region "Ribbon Callbacks"

    Public Function IsRibbonCommandEnabled(ctrl As Microsoft.Office.Core.IRibbonControl) As Boolean
        Return IsCommandEnabled(ctrl.Id)
    End Function

    Public Sub OnRibbonButtonClick(control As Microsoft.Office.Core.IRibbonControl)
        OnCommand(control.Id)
    End Sub

    Public Sub OnRibbonLoad(ribbonUI As Microsoft.Office.Core.IRibbonUI)
        _ribbon = ribbonUI
    End Sub

#End Region

    Public Sub UpdateRibbon()
        _ribbon.Invalidate()
    End Sub

End Class