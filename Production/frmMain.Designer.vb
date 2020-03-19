<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ControlsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnAddBill = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnRelogin = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSettings = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.lbFullname = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lbUsername = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cbRemark = New System.Windows.Forms.ComboBox()
        Me.tbEndtime = New System.Windows.Forms.TextBox()
        Me.tbStarttime = New System.Windows.Forms.TextBox()
        Me.tbFolder = New System.Windows.Forms.TextBox()
        Me.tbProNumber = New System.Windows.Forms.TextBox()
        Me.tbFBNumber = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvQueryAnswer = New System.Windows.Forms.DataGridView()
        Me.clSendTime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clTrip = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clProNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clFBNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clQueryContent = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clQueryAnswer = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dgvBatches = New System.Windows.Forms.DataGridView()
        Me.tm = New System.Windows.Forms.Timer(Me.components)
        Me.bgRefresher = New System.ComponentModel.BackgroundWorker()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.dgvProduction = New System.Windows.Forms.DataGridView()
        Me.clDatetime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clFolder = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clProdTripNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clProdProNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clProdFBNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clProdRemark = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clStatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clStarttime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clEndtime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tbTripNumber = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.clFilename = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsTripNumber = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clTA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsBillCount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsForEntry = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clQuery = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsBilled = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clReject = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsOngoing = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clsDatetime = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clSender = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clBatchRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvQueryAnswer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBatches, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvProduction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ControlsToolStripMenuItem, Me.mnSettings})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(982, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ControlsToolStripMenuItem
        '
        Me.ControlsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnAddBill, Me.mnRelogin})
        Me.ControlsToolStripMenuItem.Name = "ControlsToolStripMenuItem"
        Me.ControlsToolStripMenuItem.Size = New System.Drawing.Size(64, 20)
        Me.ControlsToolStripMenuItem.Text = "Controls"
        '
        'mnAddBill
        '
        Me.mnAddBill.Name = "mnAddBill"
        Me.mnAddBill.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnAddBill.Size = New System.Drawing.Size(152, 22)
        Me.mnAddBill.Text = "Add Bill"
        '
        'mnRelogin
        '
        Me.mnRelogin.Name = "mnRelogin"
        Me.mnRelogin.Size = New System.Drawing.Size(152, 22)
        Me.mnRelogin.Text = "Re-Login"
        '
        'mnSettings
        '
        Me.mnSettings.Name = "mnSettings"
        Me.mnSettings.Size = New System.Drawing.Size(61, 20)
        Me.mnSettings.Text = "Settings"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lbFullname, Me.lbUsername})
        Me.StatusStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 399)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 16, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(982, 20)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lbFullname
        '
        Me.lbFullname.Name = "lbFullname"
        Me.lbFullname.Size = New System.Drawing.Size(71, 17)
        Me.lbFullname.Text = "USERNAME:"
        '
        'lbUsername
        '
        Me.lbUsername.Name = "lbUsername"
        Me.lbUsername.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lbUsername.Size = New System.Drawing.Size(68, 15)
        Me.lbUsername.Text = "USERNAME"
        Me.lbUsername.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'cbRemark
        '
        Me.cbRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbRemark.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbRemark.FormattingEnabled = True
        Me.cbRemark.Items.AddRange(New Object() {"Full Audit", "QAFA", "Processed in Synergize", "Skipped Bills"})
        Me.cbRemark.Location = New System.Drawing.Point(67, 241)
        Me.cbRemark.Name = "cbRemark"
        Me.cbRemark.Size = New System.Drawing.Size(116, 24)
        Me.cbRemark.TabIndex = 3
        '
        'tbEndtime
        '
        Me.tbEndtime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbEndtime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbEndtime.Location = New System.Drawing.Point(67, 197)
        Me.tbEndtime.Name = "tbEndtime"
        Me.tbEndtime.ReadOnly = True
        Me.tbEndtime.Size = New System.Drawing.Size(116, 22)
        Me.tbEndtime.TabIndex = 13
        '
        'tbStarttime
        '
        Me.tbStarttime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbStarttime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbStarttime.Location = New System.Drawing.Point(67, 171)
        Me.tbStarttime.Name = "tbStarttime"
        Me.tbStarttime.ReadOnly = True
        Me.tbStarttime.Size = New System.Drawing.Size(116, 22)
        Me.tbStarttime.TabIndex = 7
        '
        'tbFolder
        '
        Me.tbFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFolder.Enabled = False
        Me.tbFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbFolder.Location = New System.Drawing.Point(67, 77)
        Me.tbFolder.Name = "tbFolder"
        Me.tbFolder.ReadOnly = True
        Me.tbFolder.Size = New System.Drawing.Size(116, 22)
        Me.tbFolder.TabIndex = 0
        '
        'tbProNumber
        '
        Me.tbProNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbProNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbProNumber.Location = New System.Drawing.Point(67, 103)
        Me.tbProNumber.MaxLength = 8
        Me.tbProNumber.Name = "tbProNumber"
        Me.tbProNumber.Size = New System.Drawing.Size(116, 22)
        Me.tbProNumber.TabIndex = 1
        '
        'tbFBNumber
        '
        Me.tbFBNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbFBNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbFBNumber.Location = New System.Drawing.Point(67, 128)
        Me.tbFBNumber.MaxLength = 5
        Me.tbFBNumber.Name = "tbFBNumber"
        Me.tbFBNumber.Size = New System.Drawing.Size(116, 22)
        Me.tbFBNumber.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(17, 203)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(52, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "End Time"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(30, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(36, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Folder"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(35, 109)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Pro #"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 177)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Start Time"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 247)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Remark"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(38, 134)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "FB #"
        '
        'dgvQueryAnswer
        '
        Me.dgvQueryAnswer.AllowUserToAddRows = False
        Me.dgvQueryAnswer.AllowUserToDeleteRows = False
        Me.dgvQueryAnswer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvQueryAnswer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvQueryAnswer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clSendTime, Me.clTrip, Me.clProNumber, Me.clFBNumber, Me.clQueryContent, Me.clQueryAnswer})
        Me.dgvQueryAnswer.Location = New System.Drawing.Point(12, 301)
        Me.dgvQueryAnswer.MultiSelect = False
        Me.dgvQueryAnswer.Name = "dgvQueryAnswer"
        Me.dgvQueryAnswer.RowHeadersVisible = False
        Me.dgvQueryAnswer.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvQueryAnswer.Size = New System.Drawing.Size(963, 93)
        Me.dgvQueryAnswer.TabIndex = 14
        '
        'clSendTime
        '
        Me.clSendTime.HeaderText = "Time Sent"
        Me.clSendTime.Name = "clSendTime"
        Me.clSendTime.ReadOnly = True
        Me.clSendTime.Width = 140
        '
        'clTrip
        '
        Me.clTrip.HeaderText = "Trip #"
        Me.clTrip.Name = "clTrip"
        Me.clTrip.ReadOnly = True
        Me.clTrip.Width = 75
        '
        'clProNumber
        '
        Me.clProNumber.HeaderText = "Pro #"
        Me.clProNumber.Name = "clProNumber"
        Me.clProNumber.ReadOnly = True
        Me.clProNumber.Width = 75
        '
        'clFBNumber
        '
        Me.clFBNumber.HeaderText = "FB #"
        Me.clFBNumber.Name = "clFBNumber"
        Me.clFBNumber.ReadOnly = True
        Me.clFBNumber.Width = 75
        '
        'clQueryContent
        '
        Me.clQueryContent.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clQueryContent.HeaderText = "Query"
        Me.clQueryContent.Name = "clQueryContent"
        Me.clQueryContent.ReadOnly = True
        '
        'clQueryAnswer
        '
        Me.clQueryAnswer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clQueryAnswer.HeaderText = "Query Answer"
        Me.clQueryAnswer.Name = "clQueryAnswer"
        Me.clQueryAnswer.ReadOnly = True
        '
        'dgvBatches
        '
        Me.dgvBatches.AllowUserToAddRows = False
        Me.dgvBatches.AllowUserToDeleteRows = False
        Me.dgvBatches.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.dgvBatches.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBatches.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clFilename, Me.clsTripNumber, Me.clTA, Me.clsBillCount, Me.clsForEntry, Me.clQuery, Me.clsBilled, Me.clReject, Me.clsOngoing, Me.clsDatetime, Me.clSender, Me.clBatchRemarks})
        Me.dgvBatches.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvBatches.Location = New System.Drawing.Point(3, 3)
        Me.dgvBatches.MultiSelect = False
        Me.dgvBatches.Name = "dgvBatches"
        Me.dgvBatches.RowHeadersVisible = False
        Me.dgvBatches.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvBatches.Size = New System.Drawing.Size(777, 238)
        Me.dgvBatches.TabIndex = 0
        '
        'tm
        '
        '
        'bgRefresher
        '
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(191, 27)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(791, 272)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvBatches)
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(783, 244)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Batches"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvProduction)
        Me.TabPage2.Location = New System.Drawing.Point(4, 24)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(783, 244)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Production"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvProduction
        '
        Me.dgvProduction.AllowUserToAddRows = False
        Me.dgvProduction.AllowUserToDeleteRows = False
        Me.dgvProduction.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProduction.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clDatetime, Me.clFolder, Me.clProdTripNumber, Me.clProdProNumber, Me.clProdFBNumber, Me.clProdRemark, Me.clStatus, Me.clStarttime, Me.clEndtime})
        Me.dgvProduction.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvProduction.Location = New System.Drawing.Point(3, 3)
        Me.dgvProduction.Name = "dgvProduction"
        Me.dgvProduction.RowHeadersVisible = False
        Me.dgvProduction.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvProduction.Size = New System.Drawing.Size(777, 238)
        Me.dgvProduction.TabIndex = 0
        '
        'clDatetime
        '
        Me.clDatetime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clDatetime.HeaderText = "Datetime"
        Me.clDatetime.Name = "clDatetime"
        Me.clDatetime.ReadOnly = True
        Me.clDatetime.Width = 82
        '
        'clFolder
        '
        Me.clFolder.HeaderText = "Folder"
        Me.clFolder.Name = "clFolder"
        Me.clFolder.ReadOnly = True
        '
        'clProdTripNumber
        '
        Me.clProdTripNumber.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clProdTripNumber.HeaderText = "Trip #"
        Me.clProdTripNumber.Name = "clProdTripNumber"
        Me.clProdTripNumber.ReadOnly = True
        Me.clProdTripNumber.Width = 63
        '
        'clProdProNumber
        '
        Me.clProdProNumber.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clProdProNumber.HeaderText = "Pro #"
        Me.clProdProNumber.Name = "clProdProNumber"
        Me.clProdProNumber.ReadOnly = True
        Me.clProdProNumber.Width = 61
        '
        'clProdFBNumber
        '
        Me.clProdFBNumber.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clProdFBNumber.HeaderText = "FB #"
        Me.clProdFBNumber.Name = "clProdFBNumber"
        Me.clProdFBNumber.ReadOnly = True
        Me.clProdFBNumber.Width = 57
        '
        'clProdRemark
        '
        Me.clProdRemark.HeaderText = "Remark"
        Me.clProdRemark.Name = "clProdRemark"
        Me.clProdRemark.ReadOnly = True
        '
        'clStatus
        '
        Me.clStatus.HeaderText = "Status"
        Me.clStatus.Name = "clStatus"
        Me.clStatus.ReadOnly = True
        '
        'clStarttime
        '
        Me.clStarttime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clStarttime.HeaderText = "Start Time"
        Me.clStarttime.Name = "clStarttime"
        Me.clStarttime.ReadOnly = True
        Me.clStarttime.Width = 88
        '
        'clEndtime
        '
        Me.clEndtime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clEndtime.HeaderText = "End Time"
        Me.clEndtime.Name = "clEndtime"
        Me.clEndtime.ReadOnly = True
        Me.clEndtime.Width = 85
        '
        'tbTripNumber
        '
        Me.tbTripNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbTripNumber.Enabled = False
        Me.tbTripNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbTripNumber.Location = New System.Drawing.Point(67, 37)
        Me.tbTripNumber.Name = "tbTripNumber"
        Me.tbTripNumber.ReadOnly = True
        Me.tbTripNumber.Size = New System.Drawing.Size(116, 22)
        Me.tbTripNumber.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(33, 43)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(35, 13)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Trip #"
        '
        'clFilename
        '
        Me.clFilename.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.clFilename.HeaderText = "Filename"
        Me.clFilename.Name = "clFilename"
        Me.clFilename.ReadOnly = True
        Me.clFilename.Width = 84
        '
        'clsTripNumber
        '
        Me.clsTripNumber.HeaderText = "Trip#"
        Me.clsTripNumber.Name = "clsTripNumber"
        Me.clsTripNumber.ReadOnly = True
        Me.clsTripNumber.Width = 50
        '
        'clTA
        '
        Me.clTA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clTA.HeaderText = "TA"
        Me.clTA.Name = "clTA"
        Me.clTA.ReadOnly = True
        Me.clTA.Width = 46
        '
        'clsBillCount
        '
        Me.clsBillCount.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clsBillCount.HeaderText = "Bill Count"
        Me.clsBillCount.Name = "clsBillCount"
        Me.clsBillCount.ReadOnly = True
        Me.clsBillCount.Width = 84
        '
        'clsForEntry
        '
        Me.clsForEntry.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clsForEntry.HeaderText = "For Entry"
        Me.clsForEntry.Name = "clsForEntry"
        Me.clsForEntry.ReadOnly = True
        Me.clsForEntry.Width = 80
        '
        'clQuery
        '
        Me.clQuery.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clQuery.HeaderText = "Query"
        Me.clQuery.Name = "clQuery"
        Me.clQuery.ReadOnly = True
        Me.clQuery.Width = 64
        '
        'clsBilled
        '
        Me.clsBilled.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clsBilled.HeaderText = "Billed"
        Me.clsBilled.Name = "clsBilled"
        Me.clsBilled.ReadOnly = True
        Me.clsBilled.Width = 63
        '
        'clReject
        '
        Me.clReject.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clReject.HeaderText = "Reject"
        Me.clReject.Name = "clReject"
        Me.clReject.ReadOnly = True
        Me.clReject.Width = 67
        '
        'clsOngoing
        '
        Me.clsOngoing.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader
        Me.clsOngoing.HeaderText = "Ongoing"
        Me.clsOngoing.Name = "clsOngoing"
        Me.clsOngoing.ReadOnly = True
        Me.clsOngoing.Width = 79
        '
        'clsDatetime
        '
        Me.clsDatetime.HeaderText = "Datetime"
        Me.clsDatetime.Name = "clsDatetime"
        Me.clsDatetime.Width = 140
        '
        'clSender
        '
        Me.clSender.HeaderText = "Sent By"
        Me.clSender.Name = "clSender"
        Me.clSender.ReadOnly = True
        '
        'clBatchRemarks
        '
        Me.clBatchRemarks.HeaderText = "Remarks"
        Me.clBatchRemarks.Name = "clBatchRemarks"
        Me.clBatchRemarks.ReadOnly = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(982, 419)
        Me.Controls.Add(Me.tbTripNumber)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cbRemark)
        Me.Controls.Add(Me.dgvQueryAnswer)
        Me.Controls.Add(Me.tbEndtime)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.tbStarttime)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.tbFolder)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.tbProNumber)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbFBNumber)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvQueryAnswer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBatches, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvProduction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tbEndtime As System.Windows.Forms.TextBox
    Friend WithEvents tbStarttime As System.Windows.Forms.TextBox
    Friend WithEvents tbFolder As System.Windows.Forms.TextBox
    Friend WithEvents tbProNumber As System.Windows.Forms.TextBox
    Friend WithEvents tbFBNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ControlsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cbRemark As System.Windows.Forms.ComboBox
    Friend WithEvents dgvBatches As System.Windows.Forms.DataGridView
    Friend WithEvents lbUsername As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mnAddBill As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSettings As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tm As System.Windows.Forms.Timer
    Friend WithEvents bgRefresher As System.ComponentModel.BackgroundWorker
    Friend WithEvents lbFullname As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mnRelogin As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dgvQueryAnswer As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvProduction As System.Windows.Forms.DataGridView
    Friend WithEvents tbTripNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents clDatetime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clFolder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clProdTripNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clProdProNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clProdFBNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clProdRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clStarttime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clEndtime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clSendTime As DataGridViewTextBoxColumn
    Friend WithEvents clTrip As DataGridViewTextBoxColumn
    Friend WithEvents clProNumber As DataGridViewTextBoxColumn
    Friend WithEvents clFBNumber As DataGridViewTextBoxColumn
    Friend WithEvents clQueryContent As DataGridViewTextBoxColumn
    Friend WithEvents clQueryAnswer As DataGridViewTextBoxColumn
    Friend WithEvents clFilename As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsTripNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clTA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsBillCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsForEntry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clQuery As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsBilled As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clReject As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsOngoing As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clsDatetime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clSender As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clBatchRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
