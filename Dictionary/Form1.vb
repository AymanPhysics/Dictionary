Imports System.IO
Imports Microsoft.VisualBasic.Devices

Public Class Form1
    Inherits Form

    Dim dt As New DataTable With {.TableName = "Dictionary"}
    Dim dst As New DataGridViewCellStyle
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Dim Dv As New DataView
    Private Sub MySettingsBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        DG.EndEdit()
        dt.WriteXml("data.xml", XmlWriteMode.WriteSchema, True)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            dt.ReadXml("data.xml")
            Dv.Table = dt
            DG.DataSource = Dv
            DG.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
            Label1.Text = DG.Rows.Count
            DG.Columns("wav").ReadOnly = True
            DG.Columns("wav").DefaultCellStyle = (New DataGridViewLinkCell).Style
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs)
        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return
        dt.Columns.Add(xx)
        Dv.Table = dt
        DG.DataSource = Dv
    End Sub

    Private Sub DG_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG.CellDoubleClick
        Try
            If dt.Columns(DG.CurrentCell.ColumnIndex).DataType = GetType(System.Byte()) Then
                Dim o As New OpenFileDialog
                o.Filter = "Images (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*"
                If o.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim img As Image = Image.FromFile(o.FileName)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = False
                    dt.Rows(DG.CurrentCell.RowIndex)(DG.CurrentCell.ColumnIndex) = imageToByteArray(img)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = True
                End If
            End If

            If DG.Columns(DG.CurrentCell.ColumnIndex).Name.ToLower = "wav" Then
                Dim o As New OpenFileDialog
                o.Filter = "wav files (*.wav)|*.wav|All files (*.*)|*.*"
                If o.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim s As String = o.FileName.Split("\")(o.FileName.Split("\").Length - 1)
                    IO.File.Copy(o.FileName, Application.StartupPath & "\audio\" & s, True)
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = False
                    dt.Rows(DG.CurrentCell.RowIndex)(DG.CurrentCell.ColumnIndex) = s
                    dt.Columns(DG.CurrentCell.ColumnIndex).ReadOnly = True
                End If
            End If
        Catch
        End Try
    End Sub

    Function imageToByteArray(imageIn As System.Drawing.Image) As Byte()
        Dim ms As New MemoryStream()
        imageIn.GetThumbnailImage(100, 100, System.Drawing.Image.GetThumbnailImageAbort.Combine, System.IntPtr.Zero).Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
        Return ms.ToArray()
    End Function


    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs)
        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return

        Dim d As New DataColumn With {.Caption = xx, .ColumnName = xx, .ReadOnly = True}
        d.DataType = GetType(System.Byte())
        dt.Columns.Add(d)
        dt.Columns(d.ColumnName).Caption = xx
        Dv.Table = dt
        DG.DataSource = Dv
    End Sub

    Private Sub ToolStripButton3_Click_1(sender As Object, e As EventArgs)

        Dim xx As String = InputBox("أدخل اسم العمود")
        If xx.Trim = "" Then Return

        Dim d As New DataColumn With {.Caption = xx, .ColumnName = xx}
        dt.Columns.Add(d)
        dt.Columns(d.ColumnName).Caption = xx
        Dv.Table = dt
        DG.DataSource = Dv
        DG.Columns(d.ColumnName).ReadOnly = True


    End Sub

    Private Sub DG_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG.CellClick
        Try
            If DG.Columns(DG.CurrentCell.ColumnIndex).Name.ToLower = "wav" Then
                Dim a As New Audio
                a.Play(Application.StartupPath & "\audio\" & DG.CurrentCell.Value, AudioPlayMode.Background)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs)
        dt.Columns.RemoveAt(DG.CurrentCell.ColumnIndex)
    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Try
            Dv.RowFilter = "([" & dt.Columns(0).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox1.Text.Trim & "%' or [" & dt.Columns(0).ColumnName & "] like '%/ " & TextBox1.Text.Trim & "%' or [" & dt.Columns(0).ColumnName & "] like '%(" & TextBox1.Text.Trim & "%') and [" & dt.Columns(1).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox2.Text.Trim & "%' and [" & dt.Columns(2).ColumnName & "] like '" & IIf(RadioButton2.Checked, "%", "") & TextBox3.Text.Trim & "%'"
            Label1.Text = DG.Rows.Count
        Catch ex As Exception
        End Try
    End Sub


    Private Sub DG_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DG.DataError

    End Sub

    Private Sub DG_SelectionChanged(sender As Object, e As EventArgs) Handles DG.SelectionChanged

    End Sub

    Private Sub DG_RowValidated(sender As Object, e As DataGridViewCellEventArgs) Handles DG.RowValidated
        Try

            If dt.Rows(DG.CurrentRow.Index)(0).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(0) = "-"
            End If
            If dt.Rows(DG.CurrentRow.Index)(1).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(1) = "-"
            End If
            If dt.Rows(DG.CurrentRow.Index)(2).ToString = "" Then
                dt.Rows(DG.CurrentRow.Index)(2) = "-"
            End If

        Catch ex As Exception

        End Try
    End Sub
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.DG = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        CType(Me.DG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DG
        '
        Me.DG.AllowUserToOrderColumns = True
        Me.DG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DG.Location = New System.Drawing.Point(0, 74)
        Me.DG.Name = "DG"
        Me.DG.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DG.Size = New System.Drawing.Size(622, 250)
        Me.DG.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(239, 346)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(13, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "0"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(472, 48)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 4
        '
        'RadioButton1
        '
        Me.RadioButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(472, 25)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(74, 17)
        Me.RadioButton1.TabIndex = 5
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Start With"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(404, 25)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(62, 17)
        Me.RadioButton2.TabIndex = 5
        Me.RadioButton2.Text = "Contain"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Location = New System.Drawing.Point(366, 48)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 6
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox3.Location = New System.Drawing.Point(260, 48)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(100, 20)
        Me.TextBox3.TabIndex = 7
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton2, Me.ToolStripButton1, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 0)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(634, 25)
        Me.BindingNavigator1.TabIndex = 8
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "Add new"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "Add new"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(634, 381)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DG)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DG As New System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As New System.Windows.Forms.Label
    Friend WithEvents TextBox1 As New System.Windows.Forms.TextBox
    Friend WithEvents RadioButton1 As New System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As New System.Windows.Forms.RadioButton
    Friend WithEvents TextBox2 As New System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As New System.Windows.Forms.TextBox
    Private components As System.ComponentModel.IContainer
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

End Class
