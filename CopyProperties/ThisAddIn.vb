Public Class ThisAddIn

    Private ReadOnly _addin As Addin = New Addin()

    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addin
    End Function
    
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _addin.Startup(Application)
    End Sub

End Class
