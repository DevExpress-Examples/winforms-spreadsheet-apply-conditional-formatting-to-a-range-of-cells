Namespace ConditionalFormatting_Examples

    Partial Class Form1

        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

'#Region "Windows Form Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.spreadsheetControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.treeListColumn1 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
            Me.splitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.button1 = New System.Windows.Forms.Button()
            Me.splitContainerControl2 = New DevExpress.XtraEditors.SplitContainerControl()
            CType((Me.treeList1), System.ComponentModel.ISupportInitialize).BeginInit()
            CType((Me.splitContainerControl1), System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl1.SuspendLayout()
            CType((Me.splitContainerControl2), System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl2.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' spreadsheetControl1
            ' 
            Me.spreadsheetControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.spreadsheetControl1.Location = New System.Drawing.Point(0, 0)
            Me.spreadsheetControl1.Name = "spreadsheetControl1"
            Me.spreadsheetControl1.Options.Culture = New System.Globalization.CultureInfo("en-US")
            Me.spreadsheetControl1.Options.Export.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Export.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Export.Txt.ValueSeparator = ","c
            Me.spreadsheetControl1.Options.Import.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Import.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Size = New System.Drawing.Size(832, 721)
            Me.spreadsheetControl1.TabIndex = 0
            Me.spreadsheetControl1.Text = "spreadsheetControl1"
            ' 
            ' treeList1
            ' 
            Me.treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            Me.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.DarkGreen
            Me.treeList1.Appearance.FocusedCell.Options.UseFont = True
            Me.treeList1.Appearance.FocusedCell.Options.UseForeColor = True
            Me.treeList1.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.treeListColumn1})
            Me.treeList1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.treeList1.Location = New System.Drawing.Point(0, 0)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.OptionsBehavior.Editable = False
            Me.treeList1.OptionsView.ShowColumns = False
            Me.treeList1.OptionsView.ShowIndicator = False
            Me.treeList1.Size = New System.Drawing.Size(307, 666)
            Me.treeList1.TabIndex = 1
            ' 
            ' treeListColumn1
            ' 
            Me.treeListColumn1.Caption = "Name"
            Me.treeListColumn1.FieldName = "Name"
            Me.treeListColumn1.Name = "treeListColumn1"
            Me.treeListColumn1.Visible = True
            Me.treeListColumn1.VisibleIndex = 0
            Me.treeListColumn1.Width = 92
            ' 
            ' splitContainerControl1
            ' 
            Me.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.splitContainerControl1.Horizontal = False
            Me.splitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl1.Name = "splitContainerControl1"
            Me.splitContainerControl1.Panel1.Controls.Add(Me.treeList1)
            Me.splitContainerControl1.Panel1.Text = "Panel1"
            Me.splitContainerControl1.Panel2.Controls.Add(Me.button1)
            Me.splitContainerControl1.Panel2.MinSize = 50
            Me.splitContainerControl1.Panel2.Text = "Panel2"
            Me.splitContainerControl1.Size = New System.Drawing.Size(307, 721)
            Me.splitContainerControl1.SplitterPosition = 716
            Me.splitContainerControl1.TabIndex = 2
            Me.splitContainerControl1.Text = "splitContainerControl1"
            ' 
            ' button1
            ' 
            Me.button1.BackColor = System.Drawing.SystemColors.ButtonHighlight
            Me.button1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte((0))))
            Me.button1.Location = New System.Drawing.Point(0, 0)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(307, 50)
            Me.button1.TabIndex = 0
            Me.button1.Text = "Run"
            Me.button1.UseVisualStyleBackColor = False
            AddHandler Me.button1.Click, New System.EventHandler(AddressOf Me.button1_Click)
            ' 
            ' splitContainerControl2
            ' 
            Me.splitContainerControl2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.splitContainerControl2.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl2.Name = "splitContainerControl2"
            Me.splitContainerControl2.Panel1.Controls.Add(Me.splitContainerControl1)
            Me.splitContainerControl2.Panel1.Text = "Panel1"
            Me.splitContainerControl2.Panel2.Controls.Add(Me.spreadsheetControl1)
            Me.splitContainerControl2.Panel2.Text = "Panel2"
            Me.splitContainerControl2.Size = New System.Drawing.Size(1144, 721)
            Me.splitContainerControl2.SplitterPosition = 307
            Me.splitContainerControl2.TabIndex = 3
            Me.splitContainerControl2.Text = "splitContainerControl2"
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1144, 721)
            Me.Controls.Add(Me.splitContainerControl2)
            Me.MaximumSize = New System.Drawing.Size(1160, 760)
            Me.MinimumSize = New System.Drawing.Size(1160, 760)
            Me.Name = "Form1"
            Me.Text = "Conditional Formatting Examples"
            CType((Me.treeList1), System.ComponentModel.ISupportInitialize).EndInit()
            CType((Me.splitContainerControl1), System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl1.ResumeLayout(False)
            CType((Me.splitContainerControl2), System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl2.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

'#End Region
        Private spreadsheetControl1 As DevExpress.XtraSpreadsheet.SpreadsheetControl

        Private treeList1 As DevExpress.XtraTreeList.TreeList

        Private splitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl

        Private button1 As System.Windows.Forms.Button

        Private splitContainerControl2 As DevExpress.XtraEditors.SplitContainerControl

        Private treeListColumn1 As DevExpress.XtraTreeList.Columns.TreeListColumn
    End Class
End Namespace
