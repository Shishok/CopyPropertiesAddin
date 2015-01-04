<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgCopyProps
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ckb_3 = New System.Windows.Forms.CheckBox()
        Me.ckb_4 = New System.Windows.Forms.CheckBox()
        Me.ckb_5 = New System.Windows.Forms.CheckBox()
        Me.ckb_6 = New System.Windows.Forms.CheckBox()
        Me.ckb_7 = New System.Windows.Forms.CheckBox()
        Me.ckb_8 = New System.Windows.Forms.CheckBox()
        Me.ckb_10 = New System.Windows.Forms.CheckBox()
        Me.btnRuEng = New System.Windows.Forms.Button()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.ckbSectionsRows = New System.Windows.Forms.CheckedListBox()
        Me.ckbReplaceContent = New System.Windows.Forms.CheckBox()
        Me.ckbRemoveMissing = New System.Windows.Forms.CheckBox()
        Me.btnClipBoard = New System.Windows.Forms.Button()
        Me.ckb_30 = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AccessibleDescription = "Разделы"
        Me.Label1.AccessibleName = "Sections"
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Разделы"
        '
        'ckb_3
        '
        Me.ckb_3.AccessibleDescription = "Пользовательские ячейки"
        Me.ckb_3.AccessibleName = "User-defined Cells"
        Me.ckb_3.AutoSize = True
        Me.ckb_3.Enabled = False
        Me.ckb_3.Location = New System.Drawing.Point(15, 32)
        Me.ckb_3.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_3.Name = "ckb_3"
        Me.ckb_3.Size = New System.Drawing.Size(161, 17)
        Me.ckb_3.TabIndex = 1
        Me.ckb_3.Tag = "242"
        Me.ckb_3.Text = "Пользовательские ячейки"
        Me.ckb_3.UseVisualStyleBackColor = True
        '
        'ckb_4
        '
        Me.ckb_4.AccessibleDescription = "Данные фигуры"
        Me.ckb_4.AccessibleName = "Shape Data"
        Me.ckb_4.AutoSize = True
        Me.ckb_4.Enabled = False
        Me.ckb_4.Location = New System.Drawing.Point(15, 51)
        Me.ckb_4.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_4.Name = "ckb_4"
        Me.ckb_4.Size = New System.Drawing.Size(108, 17)
        Me.ckb_4.TabIndex = 2
        Me.ckb_4.Tag = "243"
        Me.ckb_4.Text = "Данные фигуры"
        Me.ckb_4.UseVisualStyleBackColor = True
        '
        'ckb_5
        '
        Me.ckb_5.AccessibleDescription = "Гиперссылки"
        Me.ckb_5.AccessibleName = "Hyperlinks"
        Me.ckb_5.AutoSize = True
        Me.ckb_5.Enabled = False
        Me.ckb_5.Location = New System.Drawing.Point(15, 70)
        Me.ckb_5.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_5.Name = "ckb_5"
        Me.ckb_5.Size = New System.Drawing.Size(94, 17)
        Me.ckb_5.TabIndex = 3
        Me.ckb_5.Tag = "244"
        Me.ckb_5.Text = "Гиперссылки"
        Me.ckb_5.UseVisualStyleBackColor = True
        '
        'ckb_6
        '
        Me.ckb_6.AccessibleDescription = "Точки соединений"
        Me.ckb_6.AccessibleName = "Connection Points"
        Me.ckb_6.AutoSize = True
        Me.ckb_6.Enabled = False
        Me.ckb_6.ForeColor = System.Drawing.Color.Red
        Me.ckb_6.Location = New System.Drawing.Point(15, 89)
        Me.ckb_6.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_6.Name = "ckb_6"
        Me.ckb_6.Size = New System.Drawing.Size(119, 17)
        Me.ckb_6.TabIndex = 4
        Me.ckb_6.Tag = "7"
        Me.ckb_6.Text = "Точки соединений"
        Me.ckb_6.UseVisualStyleBackColor = True
        '
        'ckb_7
        '
        Me.ckb_7.AccessibleDescription = "Действия"
        Me.ckb_7.AccessibleName = "Actions"
        Me.ckb_7.AutoSize = True
        Me.ckb_7.Enabled = False
        Me.ckb_7.Location = New System.Drawing.Point(15, 108)
        Me.ckb_7.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_7.Name = "ckb_7"
        Me.ckb_7.Size = New System.Drawing.Size(76, 17)
        Me.ckb_7.TabIndex = 5
        Me.ckb_7.Tag = "240"
        Me.ckb_7.Text = "Действия"
        Me.ckb_7.UseVisualStyleBackColor = True
        '
        'ckb_8
        '
        Me.ckb_8.AccessibleDescription = "Элементы управления"
        Me.ckb_8.AccessibleName = "Controls"
        Me.ckb_8.AutoSize = True
        Me.ckb_8.Enabled = False
        Me.ckb_8.Location = New System.Drawing.Point(15, 127)
        Me.ckb_8.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_8.Name = "ckb_8"
        Me.ckb_8.Size = New System.Drawing.Size(140, 17)
        Me.ckb_8.TabIndex = 6
        Me.ckb_8.Tag = "9"
        Me.ckb_8.Text = "Элементы управления"
        Me.ckb_8.UseVisualStyleBackColor = True
        '
        'ckb_10
        '
        Me.ckb_10.AccessibleDescription = "Вспомогательный"
        Me.ckb_10.AccessibleName = "Scratch"
        Me.ckb_10.AutoSize = True
        Me.ckb_10.Enabled = False
        Me.ckb_10.ForeColor = System.Drawing.Color.Red
        Me.ckb_10.Location = New System.Drawing.Point(15, 146)
        Me.ckb_10.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_10.Name = "ckb_10"
        Me.ckb_10.Size = New System.Drawing.Size(119, 17)
        Me.ckb_10.TabIndex = 7
        Me.ckb_10.Tag = "6"
        Me.ckb_10.Text = "Вспомогательный"
        Me.ckb_10.UseVisualStyleBackColor = True
        '
        'btnRuEng
        '
        Me.btnRuEng.AccessibleDescription = "Eng"
        Me.btnRuEng.AccessibleName = "Ru"
        Me.btnRuEng.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRuEng.AutoSize = True
        Me.btnRuEng.Location = New System.Drawing.Point(15, 271)
        Me.btnRuEng.Name = "btnRuEng"
        Me.btnRuEng.Size = New System.Drawing.Size(50, 23)
        Me.btnRuEng.TabIndex = 15
        Me.btnRuEng.Text = "Ru"
        Me.btnRuEng.UseVisualStyleBackColor = True
        '
        'btn_OK
        '
        Me.btn_OK.AccessibleDescription = "OK"
        Me.btn_OK.AccessibleName = "OK"
        Me.btn_OK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_OK.Location = New System.Drawing.Point(225, 271)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(75, 23)
        Me.btn_OK.TabIndex = 12
        Me.btn_OK.Text = "OK"
        Me.btn_OK.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.AccessibleDescription = "Отмена"
        Me.btn_Cancel.AccessibleName = "Cancel"
        Me.btn_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Cancel.Location = New System.Drawing.Point(306, 271)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_Cancel.TabIndex = 13
        Me.btn_Cancel.Text = "Отмена"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'ckbSectionsRows
        '
        Me.ckbSectionsRows.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckbSectionsRows.CheckOnClick = True
        Me.ckbSectionsRows.FormattingEnabled = True
        Me.ckbSectionsRows.Location = New System.Drawing.Point(180, 32)
        Me.ckbSectionsRows.Name = "ckbSectionsRows"
        Me.ckbSectionsRows.Size = New System.Drawing.Size(201, 154)
        Me.ckbSectionsRows.TabIndex = 9
        '
        'ckbReplaceContent
        '
        Me.ckbReplaceContent.AccessibleDescription = "При совпадении имен строк заменять содержимое ячеек"
        Me.ckbReplaceContent.AccessibleName = "With the same name to replace the contents of the cell lines"
        Me.ckbReplaceContent.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ckbReplaceContent.AutoSize = True
        Me.ckbReplaceContent.Checked = True
        Me.ckbReplaceContent.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbReplaceContent.Location = New System.Drawing.Point(15, 208)
        Me.ckbReplaceContent.Margin = New System.Windows.Forms.Padding(1)
        Me.ckbReplaceContent.Name = "ckbReplaceContent"
        Me.ckbReplaceContent.Size = New System.Drawing.Size(321, 17)
        Me.ckbReplaceContent.TabIndex = 10
        Me.ckbReplaceContent.Tag = "6"
        Me.ckbReplaceContent.Text = "При совпадении имен строк заменять содержимое ячеек"
        Me.ckbReplaceContent.UseVisualStyleBackColor = True
        '
        'ckbRemoveMissing
        '
        Me.ckbRemoveMissing.AccessibleDescription = "Удалять отсутствующие в источнике  строки"
        Me.ckbRemoveMissing.AccessibleName = "Remove missing in the source string"
        Me.ckbRemoveMissing.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ckbRemoveMissing.AutoSize = True
        Me.ckbRemoveMissing.Checked = True
        Me.ckbRemoveMissing.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbRemoveMissing.Location = New System.Drawing.Point(15, 227)
        Me.ckbRemoveMissing.Margin = New System.Windows.Forms.Padding(1)
        Me.ckbRemoveMissing.Name = "ckbRemoveMissing"
        Me.ckbRemoveMissing.Size = New System.Drawing.Size(255, 17)
        Me.ckbRemoveMissing.TabIndex = 11
        Me.ckbRemoveMissing.Tag = "6"
        Me.ckbRemoveMissing.Text = "Удалять отсутствующие в источнике  строки"
        Me.ckbRemoveMissing.UseVisualStyleBackColor = True
        '
        'btnClipBoard
        '
        Me.btnClipBoard.AccessibleDescription = "Буфер обмена"
        Me.btnClipBoard.AccessibleName = "Clipboard"
        Me.btnClipBoard.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClipBoard.AutoSize = True
        Me.btnClipBoard.Location = New System.Drawing.Point(71, 271)
        Me.btnClipBoard.Name = "btnClipBoard"
        Me.btnClipBoard.Size = New System.Drawing.Size(105, 23)
        Me.btnClipBoard.TabIndex = 14
        Me.btnClipBoard.Text = "Буфер обмена"
        Me.btnClipBoard.UseVisualStyleBackColor = True
        '
        'ckb_30
        '
        Me.ckb_30.AccessibleDescription = "Тэги действий"
        Me.ckb_30.AccessibleName = "Action Tags"
        Me.ckb_30.AutoSize = True
        Me.ckb_30.Enabled = False
        Me.ckb_30.Location = New System.Drawing.Point(15, 165)
        Me.ckb_30.Margin = New System.Windows.Forms.Padding(1)
        Me.ckb_30.Name = "ckb_30"
        Me.ckb_30.Size = New System.Drawing.Size(100, 17)
        Me.ckb_30.TabIndex = 8
        Me.ckb_30.Tag = "247"
        Me.ckb_30.Text = "Теги действий"
        Me.ckb_30.UseVisualStyleBackColor = True
        '
        'dlgCopyProps
        '
        Me.AccessibleDescription = "Копирование свойств фигуры"
        Me.AccessibleName = "Copy shape's properties"
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(394, 306)
        Me.Controls.Add(Me.ckb_30)
        Me.Controls.Add(Me.btnClipBoard)
        Me.Controls.Add(Me.ckbRemoveMissing)
        Me.Controls.Add(Me.ckbReplaceContent)
        Me.Controls.Add(Me.ckbSectionsRows)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.btnRuEng)
        Me.Controls.Add(Me.ckb_10)
        Me.Controls.Add(Me.ckb_8)
        Me.Controls.Add(Me.ckb_7)
        Me.Controls.Add(Me.ckb_6)
        Me.Controls.Add(Me.ckb_5)
        Me.Controls.Add(Me.ckb_4)
        Me.Controls.Add(Me.ckb_3)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(500, 500)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(410, 344)
        Me.Name = "dlgCopyProps"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Копирование свойств фигуры"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ckb_3 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_4 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_5 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_6 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_7 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_8 As System.Windows.Forms.CheckBox
    Friend WithEvents ckb_10 As System.Windows.Forms.CheckBox
    Friend WithEvents btnRuEng As System.Windows.Forms.Button
    Friend WithEvents btn_OK As System.Windows.Forms.Button
    Friend WithEvents btn_Cancel As System.Windows.Forms.Button
    Friend WithEvents ckbSectionsRows As System.Windows.Forms.CheckedListBox
    Friend WithEvents ckbReplaceContent As System.Windows.Forms.CheckBox
    Friend WithEvents ckbRemoveMissing As System.Windows.Forms.CheckBox
    Friend WithEvents btnClipBoard As System.Windows.Forms.Button
    Friend WithEvents ckb_30 As System.Windows.Forms.CheckBox
End Class
