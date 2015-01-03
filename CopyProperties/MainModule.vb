Imports System.Drawing
Imports System.Windows.Forms

Module MainModule

    Dim frmCopyProps As System.Windows.Forms.Form = New dlgCopyProps
    Public vsoApp As Visio.Application = Globals.ThisAddIn.Application

    Public Sub Load_frmCopyProps()
        frmCopyProps = New dlgCopyProps
        frmCopyProps.Show()
    End Sub
End Module
