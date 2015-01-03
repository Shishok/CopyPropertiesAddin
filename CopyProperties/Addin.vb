Imports System.Drawing
Imports System.Windows.Forms


Partial Public Class Addin
    Public Property Application As Microsoft.Office.Interop.Visio.Application

    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String)
        On Error GoTo errD
        Select Case commandId
            Case "Command1"
                MessageBox.Show(commandId)
                Return

            Case "Command2"
                If Application.ActiveWindow.Selection.Count > 0 Then Call MainModule.Load_frmCopyProps()
                'Call MainModule.Load_frmCopyProps()
                Return

            Case "TogglePanel"
                TogglePanel()
                Return
        End Select
        Exit Sub

errD:
        MessageBox.Show("OnCommand" & vbNewLine & Err.Description)
    End Sub

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command shoudl be enabled in the user interface.
    ''' By default, all commands are enabled.
    ''' 
    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "Command1"
                ' make command1 always enabled
                Return True

            Case "Command2"
                ' make command2 enabled only if a window is opened
                Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
			
            Case "TogglePanel"
                ' make panel enabled only if we have an open drawing.
                Return IsPanelEnabled()
            Case Else
                Return True
        End Select
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
		
        If command = "TogglePanel" Then
            Return IsPanelVisible()
        End If

        Return False
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Returns a label associated with given command.
    ''' We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
    ''' 
    Public Function GetCommandLabel(command As String) As String
        Return My.Resources.ResourceManager.GetString(command & "_Label")
    End Function

    ''' 
    ''' Returns a icon associated with given command.
    ''' We assume for simplicity that icon ids are named after command commandId.
    ''' 
    Public Function GetCommandIcon(command As String) As Icon
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Icon)
    End Function
	
#Region "Panel"

    Private Sub TogglePanel()
        _panelManager.TogglePanel(Application.ActiveWindow)
    End Sub

    Private Function IsPanelEnabled() As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    Private Function IsPanelVisible() As Boolean
        Return Application IsNot Nothing AndAlso _panelManager.IsPanelOpened(Application.ActiveWindow)
    End Function

    Private _panelManager As PanelManager

#End Region
	
    Sub Startup(app As Object)
        Application = DirectCast(app, Microsoft.Office.Interop.Visio.Application)
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        _panelManager = New PanelManager(Me)
		End Sub

    Sub Shutdown()
	    _panelManager.Dispose()
    End Sub

    Sub UpdateUI()
        UpdateRibbon()
    
    End Sub
    
End Class