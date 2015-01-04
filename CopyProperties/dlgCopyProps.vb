Imports System.Windows.Forms
Imports System.Collections

Public Class dlgCopyProps
    ' Сделать с указанием имени/ID шейпа

    Dim winObj As Visio.Window '= vsoApp.ActiveWindow
    Dim strError As String = ""
    Dim arrNotSel As New ArrayList
    Dim intActiveSection As Integer

    Private Sub dlgCopyProps_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        winObj = vsoApp.ActiveWindow
        Me.Text = Me.Text & " - " & winObj.Selection(1).Name
        btnRuEng.Text = "Eng"
        Dim namesel As String = "ckb_"
        Dim ctl As System.Windows.Forms.Control

        For Each ctl In Me.Controls
            If InStr(1, ctl.Name, namesel, 1) <> 0 Then ctl.Enabled = CheckSection(ctl.Tag)
        Next

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim txtRE() As String = {"Для выполнения операции необходимо минимум два объекта! Операция прервана.",
                                 "The operation requires a minimum of two objects! Operation aborted.",
                                 "Ошибка при копировании формул секций:",
                                 "Error when copying formulas sections:",
                                 "Повторите попытку для этих секций.",
                                 "Please try again for these sections.",
                                 "CopyProperties Addin"}

        If winObj.Selection.Count < 2 Then
            If btnRuEng.Text = "Eng" Then MsgBox(txtRE(0), vbInformation + vbOKOnly, txtRE(6))
            If btnRuEng.Text = "Ru" Then MsgBox(txtRE(1), vbInformation + vbOKOnly, txtRE(6))
            Me.Close()
            Exit Sub
        End If

        If ckb_3.Checked Then Call CopyProp(ckb_3.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        If ckb_4.Checked Then Call CopyProp(ckb_4.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        If ckb_5.Checked Then Call CopyProp(ckb_5.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        'If ckb_6.Checked Then Call CopyProp(ckb_6.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        If ckb_7.Checked Then Call CopyProp(ckb_7.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        If ckb_8.Checked Then Call CopyProp(ckb_8.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        'If ckb_10.Checked Then Call CopyProp(ckb_10.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)
        If ckb_30.Checked Then Call CopyProp(ckb_30.Tag, ckbReplaceContent.Checked, ckbRemoveMissing.Checked)

        If strError <> "" Then
            If btnRuEng.Text = "Eng" Then MsgBox(txtRE(2) & vbNewLine & Strings.Left(strError, Strings.Len(strError) - 2) & vbNewLine & txtRE(4), vbInformation + vbOKOnly, txtRE(6))
            If btnRuEng.Text = "Ru" Then MsgBox(txtRE(3) & vbNewLine & Strings.Left(strError, Strings.Len(strError) - 2) & vbNewLine & txtRE(5), vbInformation + vbOKOnly, txtRE(6))
        End If

        arrNotSel.Clear()
        Me.Close()
    End Sub

    Private Sub CopyProp(arg, repC, remM) 'Копировать свойства секции Prop
        Dim vsoSel As Visio.Selection
        Dim vsoShpFst As Visio.Shape
        Dim vsoShpSec As Visio.Shape
        Dim vsoCellF As Visio.Cell
        Dim booDelRow As Boolean, booAddRow As Boolean, cRi As Integer
        Dim iRF As Integer, intSecShp As Integer

        vsoSel = winObj.Selection
        vsoShpFst = vsoSel(1)

        On Error GoTo err

        For intSecShp = 2 To vsoSel.Count 'Перебор выделенных элементов (вторичных)
            vsoShpSec = vsoSel(intSecShp)

            If Not vsoShpSec.SectionExists(arg, 0) Then
                vsoShpSec.AddSection(arg)

                For i = 0 To vsoShpFst.RowCount(arg) - 1 'Перебор строк 1 шейпа и добавление именованных строк во вторичн. шейп
                    If Not arrNotSel.Contains(vsoShpFst.CellsSRC(arg, i, 0).Name) Then vsoShpSec.AddNamedRow(arg, vsoShpFst.CellsSRC(arg, i, 0).RowName, 0) 'visTagDefault
                Next

                For i = 0 To vsoShpSec.RowCount(arg) - 1 ' Перебор строк секции
                    For j = 0 To vsoShpSec.RowsCellCount(arg, i)  ' Перебор ячеек строк секций и запись значений в них
                        vsoShpSec.CellsSRC(arg, i, j).FormulaU = vsoShpFst.CellsSRC(arg, i, j).FormulaU
                    Next
                Next
            Else
                For iRF = 0 To vsoShpFst.RowCount(arg) - 1 'Перебор строк секции Prop первичного элемента
                    vsoCellF = vsoShpFst.CellsSRC(arg, iRF, 0)

                    If Not arrNotSel.Contains(vsoCellF.Name) Then
                        booAddRow = RowNameExists(arg, vsoShpSec, vsoCellF.RowName)

                        If repC Then ' Перебор ячеек и запись значений в них если требуется
                            cRi = vsoShpSec.CellsRowIndex(vsoCellF.Name)
                            For i = 0 To vsoShpSec.RowsCellCount(arg, cRi)
                                vsoShpSec.CellsSRC(arg, cRi, i).FormulaU = vsoShpFst.CellsSRC(arg, vsoShpFst.CellsRowIndex(vsoCellF.Name), i).FormulaU
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

        strError = strError & Switch(arg = "242", "User-defined Cells, ", arg = "243", "Shape Data, ", arg = "244", "Hyperlinks, ", _
                    arg = "240", "Actions, ", arg = "9", "Controls, ", arg = "247", "Action Tags, ")

    End Sub

    Private Function RowNameExists(x, vsoShpSec, strName) ' As Boolean

        On Error GoTo err
        vsoShpSec.AddNamedRow(x, strName, 0)  'visTagDefault
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
            Select Case intActiveSection
                Case 242
                    strArrSectionData = {"RowName", "Value", "Promt"}
                    intArrSectionData = {0, 1}
                Case 243
                    strArrSectionData = {"RowName", "Label", "Promt", "Type", "Format", "Value", "SortKey", "Invisible", "Ask", "LangID", "Calendar"}
                    intArrSectionData = {2, 1, 5, 3, 0, 4, 6, 7, 14, 15}
                Case 244
                    strArrSectionData = {"RowName", "Description", "Address", "SubAddress", "ExtraInfo", "Frame", "SortKey", "NewWindow", "Default", "Invisible"}
                    intArrSectionData = {0, 1, 2, 3, 4, 15, 5, 7, 8}
                Case 7
                    strArrSectionData = {"RowName", "X", "Y", "DirX / A", "DirY / B", "Type / C", "D"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5}
                Case 240
                    strArrSectionData = {"RowName", "Action", "Menu", "TagName", "ButtonFace", "SortKey", "Checked", "Disabled", "ReadOnly", "Invisible", "BeginGroup", "FlyoutChild"}
                    intArrSectionData = {3, 0, 14, 15, 16, 4, 5, 6, 7, 8, 9}
                Case 9
                    strArrSectionData = {"RowName", "X", "Y", "X Dinamics", "Y Dinamics", "X Behavior", "Y Behavior", "CanGlue", "Tip"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5, 6, 8}
                Case 6
                    strArrSectionData = {"RowName", "X", "Y", "A", "B", "C", "D"}
                    intArrSectionData = {0, 1, 2, 3, 4, 5}
                Case 247
                    strArrSectionData = {"RowName", "X", "Y", "TagName", "X Justify", "Y Justify", "DisplayMode", "ButtonFace", "Discription", "Disabled"}
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

#Region "Control's Events"

    Private Sub ckb_3_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_3.CheckedChanged
        Call ViewSectionsRowsCustom(242)
    End Sub

    Private Sub ckb_4_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_4.CheckedChanged
        Call ViewSectionsRowsCustom(243)
    End Sub

    Private Sub ckb_5_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_5.CheckedChanged
        Call ViewSectionsRowsCustom(244)
    End Sub

    Private Sub ckb_6_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_6.CheckedChanged
        Call ViewSectionsRowsCustom(7)
    End Sub

    Private Sub ckb_7_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_7.CheckedChanged
        Call ViewSectionsRowsCustom(240)
    End Sub

    Private Sub ckb_8_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_8.CheckedChanged
        Call ViewSectionsRowsCustom(9)
    End Sub

    Private Sub ckb_10_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_10.CheckedChanged
        Call ViewSectionsRowsCustom(6)
    End Sub

    Private Sub ckb_30_CheckedChanged(sender As Object, e As EventArgs) Handles ckb_30.CheckedChanged
        Call ViewSectionsRowsCustom(247)
    End Sub

#End Region

End Class