<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class sfrmHolder
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
        Me.components = New System.ComponentModel.Container()
        Me.lbFolder = New System.Windows.Forms.Label()
        Me.lbProNumber = New System.Windows.Forms.Label()
        Me.lbFBNumber = New System.Windows.Forms.Label()
        Me.lbStarttime = New System.Windows.Forms.Label()
        Me.lbElapsetime = New System.Windows.Forms.Label()
        Me.lbRemark = New System.Windows.Forms.Label()
        Me.lbFullname = New System.Windows.Forms.Label()
        Me.lbUsername = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.tm = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnQuerytoVALC = New System.Windows.Forms.Button()
        Me.btnQuery = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lbFolder
        '
        Me.lbFolder.AutoSize = True
        Me.lbFolder.Location = New System.Drawing.Point(75, 58)
        Me.lbFolder.Name = "lbFolder"
        Me.lbFolder.Size = New System.Drawing.Size(38, 12)
        Me.lbFolder.TabIndex = 0
        Me.lbFolder.Text = "Folder:"
        '
        'lbProNumber
        '
        Me.lbProNumber.AutoSize = True
        Me.lbProNumber.Location = New System.Drawing.Point(75, 81)
        Me.lbProNumber.Name = "lbProNumber"
        Me.lbProNumber.Size = New System.Drawing.Size(33, 12)
        Me.lbProNumber.TabIndex = 1
        Me.lbProNumber.Text = "Pro #:"
        '
        'lbFBNumber
        '
        Me.lbFBNumber.AutoSize = True
        Me.lbFBNumber.Location = New System.Drawing.Point(75, 105)
        Me.lbFBNumber.Name = "lbFBNumber"
        Me.lbFBNumber.Size = New System.Drawing.Size(30, 12)
        Me.lbFBNumber.TabIndex = 2
        Me.lbFBNumber.Text = "FB #:"
        '
        'lbStarttime
        '
        Me.lbStarttime.AutoSize = True
        Me.lbStarttime.Location = New System.Drawing.Point(75, 130)
        Me.lbStarttime.Name = "lbStarttime"
        Me.lbStarttime.Size = New System.Drawing.Size(56, 12)
        Me.lbStarttime.TabIndex = 3
        Me.lbStarttime.Text = "Start Time:"
        '
        'lbElapsetime
        '
        Me.lbElapsetime.AutoSize = True
        Me.lbElapsetime.Location = New System.Drawing.Point(75, 153)
        Me.lbElapsetime.Name = "lbElapsetime"
        Me.lbElapsetime.Size = New System.Drawing.Size(71, 12)
        Me.lbElapsetime.TabIndex = 4
        Me.lbElapsetime.Text = "Elapsed Time:"
        '
        'lbRemark
        '
        Me.lbRemark.AutoSize = True
        Me.lbRemark.Location = New System.Drawing.Point(75, 188)
        Me.lbRemark.Name = "lbRemark"
        Me.lbRemark.Size = New System.Drawing.Size(44, 12)
        Me.lbRemark.TabIndex = 5
        Me.lbRemark.Text = "Remark:"
        '
        'lbFullname
        '
        Me.lbFullname.AutoSize = True
        Me.lbFullname.Location = New System.Drawing.Point(75, 12)
        Me.lbFullname.Name = "lbFullname"
        Me.lbFullname.Size = New System.Drawing.Size(50, 12)
        Me.lbFullname.TabIndex = 6
        Me.lbFullname.Text = "Fullname:"
        '
        'lbUsername
        '
        Me.lbUsername.AutoSize = True
        Me.lbUsername.Location = New System.Drawing.Point(75, 33)
        Me.lbUsername.Name = "lbUsername"
        Me.lbUsername.Size = New System.Drawing.Size(57, 12)
        Me.lbUsername.TabIndex = 7
        Me.lbUsername.Text = "Username:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(66, 248)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 21)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(147, 248)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 21)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save Bill"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'tm
        '
        Me.tm.Interval = 500
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 14)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Username:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(11, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 14)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Fullname:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(18, 188)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 14)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Remark:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(22, 153)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 28)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Elapsed" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Time:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(5, 130)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(62, 14)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Start Time:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(35, 105)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 14)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "FB #:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(32, 81)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(35, 14)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Pro #:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Calibri", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(26, 58)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 14)
        Me.Label8.TabIndex = 10
        Me.Label8.Text = "Folder:"
        '
        'btnQuerytoVALC
        '
        Me.btnQuerytoVALC.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnQuerytoVALC.Enabled = False
        Me.btnQuerytoVALC.Location = New System.Drawing.Point(66, 221)
        Me.btnQuerytoVALC.Name = "btnQuerytoVALC"
        Me.btnQuerytoVALC.Size = New System.Drawing.Size(156, 21)
        Me.btnQuerytoVALC.TabIndex = 18
        Me.btnQuerytoVALC.Text = "Send Query to VALC"
        Me.btnQuerytoVALC.UseVisualStyleBackColor = True
        Me.btnQuerytoVALC.Visible = False
        '
        'btnQuery
        '
        Me.btnQuery.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnQuery.Location = New System.Drawing.Point(66, 221)
        Me.btnQuery.Name = "btnQuery"
        Me.btnQuery.Size = New System.Drawing.Size(156, 21)
        Me.btnQuery.TabIndex = 19
        Me.btnQuery.Text = "Send Query"
        Me.btnQuery.UseVisualStyleBackColor = True
        '
        'sfrmHolder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(232, 278)
        Me.Controls.Add(Me.btnQuery)
        Me.Controls.Add(Me.btnQuerytoVALC)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lbUsername)
        Me.Controls.Add(Me.lbFullname)
        Me.Controls.Add(Me.lbRemark)
        Me.Controls.Add(Me.lbElapsetime)
        Me.Controls.Add(Me.lbStarttime)
        Me.Controls.Add(Me.lbFBNumber)
        Me.Controls.Add(Me.lbProNumber)
        Me.Controls.Add(Me.lbFolder)
        Me.Font = New System.Drawing.Font("Liberation Sans", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "sfrmHolder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbFolder As System.Windows.Forms.Label
    Friend WithEvents lbProNumber As System.Windows.Forms.Label
    Friend WithEvents lbFBNumber As System.Windows.Forms.Label
    Friend WithEvents lbStarttime As System.Windows.Forms.Label
    Friend WithEvents lbElapsetime As System.Windows.Forms.Label
    Friend WithEvents lbRemark As System.Windows.Forms.Label
    Friend WithEvents lbFullname As System.Windows.Forms.Label
    Friend WithEvents lbUsername As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents tm As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents btnQuerytoVALC As System.Windows.Forms.Button
    Friend WithEvents btnQuery As System.Windows.Forms.Button
End Class
