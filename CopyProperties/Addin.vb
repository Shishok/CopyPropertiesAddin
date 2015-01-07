Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class Addin
    Public Property Application As Microsoft.Office.Interop.Visio.Application

    Public Sub OnCommand(commandId As String)
        On Error GoTo errD
        Select Case commandId
            Case "Command2"
                If Application.ActiveWindow.Selection.Count > 0 Then Call MainModule.Load_frmCopyProps()
                Return
        End Select
        Exit Sub

errD:
        MessageBox.Show("OnCommand" & vbNewLine & Err.Description)
    End Sub

    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "Command2"
                Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
            Case Else
                Return True
        End Select
    End Function

    Sub Startup(app As Object)
        Application = DirectCast(app, Microsoft.Office.Interop.Visio.Application)
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
    End Sub

    Sub UpdateUI()
        UpdateRibbon()
    End Sub
    
End Class