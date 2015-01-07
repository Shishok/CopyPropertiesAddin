Imports System.Windows.Forms
Imports System.Collections

Public Class dlgCopyProps

    Dim winObj As Visio.Window
    Dim strError As String = ""
    Dim arrNotSel As New ArrayList
    Dim intActiveSection As Integer

    Private Sub dlgCopyProps_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ctl As System.Windows.Forms.Control
        winObj = vsoApp.ActiveWindow
        Me.Text = Me.Text & " - " & winObj.Selection(1).Name
        btnRuEng.Text = "Eng"

        For Each ctl In Me.Controls
            If InStr(1, ctl.Name, "ckb_", 1) <> 0 Then ctl.Enabled = CheckSection(ctl.Tag)
        Next

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim txtRE() As String = {"Для выполнения операции необходимо минимум два объекта! Операция прервана.",
                                 "The operation requires a minimum of two objects! Operation aborted.",
                                 "Ошибка при копировании формул ячеек:",
                                 "Error when copying formulas cells:",
                                 "Повторите попытку для этих ячеек.",
                                 "Please try again for these cells.",
                                 "CopyProperties Addin"}

        If winObj.Selection.Count < 2 Then
            If btnRuEng.Text = "Eng" Then MsgBox(txtRE(0), vbInformation + vbOKOnly, txtRE(6))
            If btnRuEng.Text = "Ru" Then MsgBox(txtRE(1), vbInformation + vbOKOnly, txtRE(6))
            Exit Sub
        End If

        Dim rC As Boolean = ckbReplaceContent.Checked
        Dim rM As Boolean = ckbRemoveMissing.Checked

        If ckb_3.Checked Then Call CopyProp(ckb_3.Tag, rC, rM)
        If ckb_4.Checked Then Call CopyProp(ckb_4.Tag, rC, rM)
        If ckb_5.Checked Then Call CopyProp(ckb_5.Tag, rC, rM)
        If ckb_6.Checked Then Call CopyProp(ckb_6.Tag, rC, rM)
        If ckb_7.Checked Then Call CopyProp(ckb_7.Tag, rC, rM)
        If ckb_8.Checked Then Call CopyProp(ckb_8.Tag, rC, rM)
        If ckb_10.Checked Then Call CopyProp(ckb_10.Tag, rC, rM)
        If ckb_30.Checked Then Call CopyProp(ckb_30.Tag, rC, rM)

        If strError <> "" Then
            If btnRuEng.Text = "Eng" Then MsgBox(txtRE(2) & vbNewLine & Strings.Left(strError, Strings.Len(strError) - 2) & vbNewLine & txtRE(4), vbInformation + vbOKOnly, txtRE(6))
            If btnRuEng.Text = "Ru" Then MsgBox(txtRE(3) & vbNewLine & Strings.Left(strError, Strings.Len(strError) - 2) & vbNewLine & txtRE(5), vbInformation + vbOKOnly, txtRE(6))
            strError = ""
        Else
            arrNotSel.Clear()
            Me.Close()
        End If
    End Sub

    Private Sub CopyProp(arg, repC, remM)
        Dim vsoSel As Visio.Selection
        Dim vsoShpFst As Visio.Shape
        Dim vsoShpSec As Visio.Shape
        Dim vsoCellF As Visio.Cell
        Dim booDelRow As Boolean, booAddRow As Boolean
        Dim intSecShp As Integer, i As Integer, j As Integer, q As Integer

        vsoSel = winObj.Selection
        vsoShpFst = vsoSel(1)

        On Error GoTo err

        For intSecShp = 2 To vsoSel.Count ' Перебор вторичных шейпов
            vsoShpSec = vsoSel(intSecShp)

            If arg = 6 Or arg = 7 Then ' Если секция Scratch или Connection Points
                vsoShpSec.DeleteSection(arg) ' Удалить секцию
            End If

            If Not vsoShpSec.SectionExists(arg, 0) Then ' Если секция отсутствует

                If arg = 6 Or arg = 7 Then ' Если секция Scratch или Connection Points
                    vsoShpSec.AddRows(arg, 0, 0, vsoShpFst.RowCount(arg))
                Else ' Другие секции
                    vsoShpSec.AddSection(arg)
                    For i = 0 To vsoShpFst.RowCount(arg) - 1
                        If Not arrNotSel.Contains(vsoShpFst.CellsSRC(arg, i, 0).Name) Then vsoShpSec.AddNamedRow(arg, vsoShpFst.CellsSRC(arg, i, 0).RowName, 0)
                    Next
                End If

                For i = 0 To vsoShpSec.RowCount(arg) - 1 ' Перебор ячеек и запись значений в них
                    For j = 0 To vsoShpSec.RowsCellCount(arg, i)
                        vsoShpSec.CellsSRC(arg, i, j).FormulaU = vsoShpFst.CellsSRC(arg, i, j).FormulaU
                    Next
                Next
            Else ' Если секция существует

                For q = 0 To vsoShpFst.RowCount(arg) - 1
                    vsoCellF = vsoShpFst.CellsSRC(arg, q, 0)

                    If Not arrNotSel.Contains(vsoCellF.Name) Then
                        booAddRow = RowNameExists(arg, vsoShpSec, vsoCellF.RowName)

                        If repC Then ' Перебор ячеек и запись значений в них если требуется
                            i = vsoShpSec.CellsRowIndex(vsoCellF.Name)
                            For j = 0 To vsoShpSec.RowsCellCount(arg, i)
                                vsoShpSec.CellsSRC(arg, i, j).FormulaU = vsoShpFst.CellsSRC(arg, vsoShpFst.CellsRowIndex(vsoCellF.Name), j).FormulaU
                            Next
                        End If
                    End If
                Next

                If remM Then ' Удаление ненужных строк из секции если требуется
                    For i = vsoShpSec.RowCount(arg) - 1 To 0 Step -1
                        booDelRow = True
                        For j = 0 To vsoShpFst.RowCount(arg) - 1
                            If vsoShpSec.CellsSRC(arg, i, 0).RowName = vsoShpFst.CellsSRC(arg, j, 0).RowName Then booDelRow = False
                        Next
                        If booDelRow Then vsoShpSec.DeleteRow(arg, i)
                    Next
                End If
            End If
        Next

        Exit Sub

err:
        strError = strError & vsoShpSec.Name & ": " & vsoShpSec.CellsSRC(arg, i, j).Name & ", "
        'strError = strError & Switch(arg = "242", "User-defined Cells, ", arg = "243", "Shape Data, ", arg = "244", "Hyperlinks, ", _
        '            arg = "240", "Actions, ", arg = "9", "Controls, ", arg = "247", "Action Tags, ")
    End Sub

    Private Function RowNameExists(arg, vsoShpSec, strName)

        On Error GoTo err
        vsoShpSec.AddNamedRow(arg, strName, 0)
        RowNameExists = False
        Exit Function

err:
        RowNameExists = True
    End Function

    Private Function CheckSection(arg)
        Return winObj.Selection(1).SectionExists(Val(arg), 0)
    End Function

    Private Sub btnRuEng_Click(sender As Object, e As EventArgs) Handles btnRuEng.Click
        Dim ctl As System.Windows.Forms.Control

        If btnRuEng.Text = "Eng" Then
            For Each ctl In Me.Controls
                ctl.Text = ctl.AccessibleName
            Next
            Me.Text = Me.AccessibleName & " - " & winObj.Selection(1).NameU
        Else
            For Each ctl In Me.Controls
                ctl.Text = ctl.AccessibleDescription
            Next
            Me.Text = Me.AccessibleDescription & " - " & winObj.Selection(1).Name
        End If

    End Sub

    Private Sub ViewSectionsRowsCustom(IndSection)
        intActiveSection = IndSection
        Dim iCell As Integer = 0
        Dim IndCell As Integer = winObj.Selection(1).RowCount(IndSection) - 1

        With ckbSectionsRows.Items
            .Clear()
            For iCell = 0 To IndCell
                .Add(winObj.Selection(1).CellsSRC(IndSection, iCell, 0).Name, True)
                If arrNotSel.Contains(winObj.Selection(1).CellsSRC(IndSection, iCell, 0).Name) Then
                    ckbSectionsRows.SetItemChecked(ckbSectionsRows.Items.Count - 1, False)
                End If
            Next
        End With

    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub ckbSectionsRows_LostFocus(sender As Object, e As EventArgs) Handles ckbSectionsRows.LostFocus
        Dim ind As Integer

        For ind = 0 To ckbSectionsRows.Items.Count - 1
            If Not ckbSectionsRows.GetItemChecked(ind) Then
                arrNotSel.Remove(ckbSectionsRows.Items(ind))
                arrNotSel.Add(ckbSectionsRows.Items(ind))
            End If
        Next
    End Sub

    Private Sub btnClipBoard_Click(sender As Object, e As EventArgs) Handles btnClipBoard.Click
        Dim shObj As Visio.Shape = winObj.Selection(1)
        Dim strArrSectionData() As String = {""}
        Dim intArrSectionData() As Integer = {0}
        Dim strTxt As String = ""
        Dim strActiveSection As String = Switch(intActiveSection = "242", "User-defined Cells", intActiveSection = "243", "Shape Data", intActiveSection = "244", "Hyperlinks, ", _
                                        intActiveSection = "240", "Actions", intActiveSection = "6", "Scratch", intActiveSection = "7", "Connection Points", intActiveSection = "9", "Controls", intActiveSection = "247", "Action Tags")

        If ckbSectionsRows.Items.Count <> 0 Then
            Const r = "RowName"
            Select Case intActiveSection
                Case 242
                    strArrSectionData = {r, "Value", "Promt"}
                    intArrSectionData = {0, 1}
                Case 243
                    strArrSectionData = {r, "Label", "Promt", "Type", "Format", "Value", "SortKey", "Invisible", "Ask", "LangID", "Calendar"}
                    intArrSectionData = {2, 1, 5, 3, 0, 4, 6, 7, 14, 15}
                Case 244
                    strArrSectionData = {r, "Description", "Address", "SubAddress", "ExtraInfo", "Frame", "SortKey", "NewWindow", "Default", "Invisible"}
                    intArrSectionData = {0, 1, 2, 3, 4, 15, 5, 7, 8}
                Case 7
                    strArrSectionData = {r, "X", "Y", "DirX / A", "DirY / B", "Type / C", "D"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5}
                Case 240
                    strArrSectionData = {r, "Action", "Menu", "TagName", "ButtonFace", "SortKey", "Checked", "Disabled", "ReadOnly", "Invisible", "BeginGroup", "FlyoutChild"}
                    intArrSectionData = {3, 0, 14, 15, 16, 4, 5, 6, 7, 8, 9}
                Case 9
                    strArrSectionData = {r, "X", "Y", "X Dinamics", "Y Dinamics", "X Behavior", "Y Behavior", "CanGlue", "Tip"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5, 6, 8}
                Case 6
                    strArrSectionData = {r, "X", "Y", "A", "B", "C", "D"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5}
                Case 247
                    strArrSectionData = {r, "X", "Y", "TagName", "X Justify", "Y Justify", "DisplayMode", "ButtonFace", "Discription", "Disabled"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5, 6, 15, 7}
            End Select

            strTxt = shObj.Name & " (" & shObj.NameU & ")" & ", "
            strTxt = strTxt & "Section - " & strActiveSection & ":" & vbCrLf

            For j = 0 To UBound(strArrSectionData)
                If j <> UBound(strArrSectionData) Then strTxt = strTxt & strArrSectionData(j) & vbTab
                If j = UBound(strArrSectionData) Then strTxt = strTxt & strArrSectionData(j) & vbCrLf
            Next

            For ind = 0 To ckbSectionsRows.Items.Count - 1
                If ckbSectionsRows.GetItemChecked(ind) Then
                    For j = 0 To UBound(intArrSectionData)
                        If j = 0 Then strTxt = strTxt & ckbSectionsRows.Items(ind) & vbTab
                        strTxt = strTxt & shObj.CellsSRC(intActiveSection, shObj.CellsRowIndex(ckbSectionsRows.Items(ind)), intArrSectionData(j)).FormulaU
                        If j <> UBound(intArrSectionData) Then strTxt = strTxt & vbTab
                    Next
                    strTxt = strTxt & vbCr
                End If
            Next

        End If

        My.Computer.Clipboard.SetText(strTxt)
    End Sub

    Private Sub lbl2(s)
        With Label2
            .AccessibleName = s.AccessibleName
            .AccessibleDescription = s.AccessibleDescription
            .Text = s.text
        End With
    End Sub

#Region "Control's Events"

    'Private Sub CheckBox_CheckedChanged(sender As Button, e As System.EventArgs) _
    '    Handles ckb_3.CheckedChanged, ckb_4.CheckedChanged, ckb_5.CheckedChanged, ckb_6.CheckedChanged, _
    '    ckb_7.CheckedChanged, ckb_8.CheckedChanged, ckb_10.CheckedChanged, ckb_30.CheckedChanged
    '    Call ViewSectionsRowsCustom(sender.Tag)
    'End Sub

    Private Sub ckb_3_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_3.CheckedChanged
        Call ViewSectionsRowsCustom(242) : lbl2(sender)
    End Sub

    Private Sub ckb_4_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_4.CheckedChanged
        Call ViewSectionsRowsCustom(243) : lbl2(sender)
    End Sub

    Private Sub ckb_5_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_5.CheckedChanged
        Call ViewSectionsRowsCustom(244) : lbl2(sender)
    End Sub

    Private Sub ckb_6_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_6.CheckedChanged
        Call ViewSectionsRowsCustom(7) : lbl2(sender)
    End Sub

    Private Sub ckb_7_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_7.CheckedChanged
        Call ViewSectionsRowsCustom(240) : lbl2(sender)
    End Sub

    Private Sub ckb_8_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_8.CheckedChanged
        Call ViewSectionsRowsCustom(9) : lbl2(sender)
    End Sub

    Private Sub ckb_10_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_10.CheckedChanged
        Call ViewSectionsRowsCustom(6) : lbl2(sender)
    End Sub

    Private Sub ckb_30_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_30.CheckedChanged
        Call ViewSectionsRowsCustom(247) : lbl2(sender)
    End Sub

#End Region

End Class