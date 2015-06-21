<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.SDI = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Barcode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AXAXCD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AXA1CD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ColumnHeadersVisible = False
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SDI, Me.Barcode, Me.AXAXCD, Me.AXA1CD})
        Me.DataGridView1.Location = New System.Drawing.Point(12, 76)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(219, 314)
        Me.DataGridView1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 396)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Return"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(172, 396)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(57, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Damage"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(219, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Label1"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(219, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Label2"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Label3"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 60)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(219, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Label4"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(93, 396)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(57, 23)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Unreturn"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'SDI
        '
        Me.SDI.DataPropertyName = "SDI"
        Me.SDI.HeaderText = "SDI"
        Me.SDI.Name = "SDI"
        Me.SDI.ReadOnly = True
        Me.SDI.Visible = False
        '
        'Barcode
        '
        Me.Barcode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Barcode.DataPropertyName = "Barcode"
        Me.Barcode.HeaderText = "Barcode"
        Me.Barcode.Name = "Barcode"
        Me.Barcode.ReadOnly = True
        '
        'AXAXCD
        '
        Me.AXAXCD.DataPropertyName = "AXAXCD"
        Me.AXAXCD.HeaderText = "Custnum"
        Me.AXAXCD.Name = "AXAXCD"
        Me.AXAXCD.ReadOnly = True
        Me.AXAXCD.Visible = False
        '
        'AXA1CD
        '
        Me.AXA1CD.DataPropertyName = "AXA1CD"
        Me.AXA1CD.HeaderText = "LocNum"
        Me.AXA1CD.Name = "AXA1CD"
        Me.AXA1CD.ReadOnly = True
        Me.AXA1CD.Visible = False
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(241, 445)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "Form2"
        Me.Text = "SDI Details"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents SDI As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Barcode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AXAXCD As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AXA1CD As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
