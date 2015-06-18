<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
                m_IBdata.Dispose()
                m_utils.Dispose()

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
        Me.tabControl = New System.Windows.Forms.TabControl()
        Me.tabConnection = New System.Windows.Forms.TabPage()
        Me.btnClrErrLog = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.lstErrors = New System.Windows.Forms.ListBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.lstServerResponses = New System.Windows.Forms.ListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.labelAccountNum = New System.Windows.Forms.Label()
        Me.labelServerTime = New System.Windows.Forms.Label()
        Me.labelConnectionStatus = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbboxLogLevel = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btnConnect = New System.Windows.Forms.Button()
        Me.btnDisconnect = New System.Windows.Forms.Button()
        Me.txtboxClientID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtboxPort = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtboxIPAddress = New System.Windows.Forms.TextBox()
        Me.tabAccount = New System.Windows.Forms.TabPage()
        Me.tabctrlAccount = New System.Windows.Forms.TabControl()
        Me.tabAcctSummary = New System.Windows.Forms.TabPage()
        Me.gridvwAcctSummary = New System.Windows.Forms.DataGridView()
        Me.tabPortfolio = New System.Windows.Forms.TabPage()
        Me.gridvwPortfolio = New System.Windows.Forms.DataGridView()
        Me.grpAcctSub = New System.Windows.Forms.GroupBox()
        Me.labelAcctTime = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnUnsubscribe = New System.Windows.Forms.Button()
        Me.btnSubscribe = New System.Windows.Forms.Button()
        Me.labelAccountNum2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tabOrders = New System.Windows.Forms.TabPage()
        Me.tabctrlOrders = New System.Windows.Forms.TabControl()
        Me.tabOpenOrders = New System.Windows.Forms.TabPage()
        Me.gridvwOpenOrders = New System.Windows.Forms.DataGridView()
        Me.tabOrderStatus = New System.Windows.Forms.TabPage()
        Me.gridvwOrderStatus = New System.Windows.Forms.DataGridView()
        Me.grpOrderDesc = New System.Windows.Forms.GroupBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.btnClearOrders = New System.Windows.Forms.Button()
        Me.btnGlobalCancel = New System.Windows.Forms.Button()
        Me.btnWhatIf = New System.Windows.Forms.Button()
        Me.btnPlaceModifyOrder = New System.Windows.Forms.Button()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.txtboxOrderID = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.btnExtOrderAttr = New System.Windows.Forms.Button()
        Me.txtboxPrice = New System.Windows.Forms.TextBox()
        Me.txtboxOrderType = New System.Windows.Forms.TextBox()
        Me.txtboxQuantity = New System.Windows.Forms.TextBox()
        Me.txtboxAction = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtboxTicker = New System.Windows.Forms.TextBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.btnExtTickerAttr = New System.Windows.Forms.Button()
        Me.txtboxCurrency = New System.Windows.Forms.TextBox()
        Me.txtboxPrimExch = New System.Windows.Forms.TextBox()
        Me.txtboxExchange = New System.Windows.Forms.TextBox()
        Me.txtboxType = New System.Windows.Forms.TextBox()
        Me.tabExecutions = New System.Windows.Forms.TabPage()
        Me.tabctrlExecutions = New System.Windows.Forms.TabControl()
        Me.tabExecReport = New System.Windows.Forms.TabPage()
        Me.gridvwExecutionsReport = New System.Windows.Forms.DataGridView()
        Me.grpReqExeReports = New System.Windows.Forms.GroupBox()
        Me.btnClrExecs = New System.Windows.Forms.Button()
        Me.btnReqExecution = New System.Windows.Forms.Button()
        Me.txtboxExecAction = New System.Windows.Forms.TextBox()
        Me.txtboxExecExch = New System.Windows.Forms.TextBox()
        Me.txtboxExecSecType = New System.Windows.Forms.TextBox()
        Me.txtboxExecSym = New System.Windows.Forms.TextBox()
        Me.txtboxExecTime = New System.Windows.Forms.TextBox()
        Me.txtboxExecAcct = New System.Windows.Forms.TextBox()
        Me.txtboxExecClientId = New System.Windows.Forms.TextBox()
        Me.txtboxExecReqId = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ctxtmenuCancelOrder = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CancelOrderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tabControl.SuspendLayout()
        Me.tabConnection.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tabAccount.SuspendLayout()
        Me.tabctrlAccount.SuspendLayout()
        Me.tabAcctSummary.SuspendLayout()
        CType(Me.gridvwAcctSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabPortfolio.SuspendLayout()
        CType(Me.gridvwPortfolio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAcctSub.SuspendLayout()
        Me.tabOrders.SuspendLayout()
        Me.tabctrlOrders.SuspendLayout()
        Me.tabOpenOrders.SuspendLayout()
        CType(Me.gridvwOpenOrders, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabOrderStatus.SuspendLayout()
        CType(Me.gridvwOrderStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpOrderDesc.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.tabExecutions.SuspendLayout()
        Me.tabctrlExecutions.SuspendLayout()
        Me.tabExecReport.SuspendLayout()
        CType(Me.gridvwExecutionsReport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpReqExeReports.SuspendLayout()
        Me.ctxtmenuCancelOrder.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabControl
        '
        Me.tabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabControl.Controls.Add(Me.tabConnection)
        Me.tabControl.Controls.Add(Me.tabAccount)
        Me.tabControl.Controls.Add(Me.tabOrders)
        Me.tabControl.Controls.Add(Me.tabExecutions)
        Me.tabControl.Location = New System.Drawing.Point(0, 0)
        Me.tabControl.Name = "tabControl"
        Me.tabControl.SelectedIndex = 0
        Me.tabControl.Size = New System.Drawing.Size(828, 434)
        Me.tabControl.TabIndex = 0
        '
        'tabConnection
        '
        Me.tabConnection.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabConnection.Controls.Add(Me.btnClrErrLog)
        Me.tabConnection.Controls.Add(Me.GroupBox5)
        Me.tabConnection.Controls.Add(Me.GroupBox4)
        Me.tabConnection.Controls.Add(Me.GroupBox2)
        Me.tabConnection.Controls.Add(Me.GroupBox1)
        Me.tabConnection.Location = New System.Drawing.Point(4, 22)
        Me.tabConnection.Name = "tabConnection"
        Me.tabConnection.Size = New System.Drawing.Size(820, 408)
        Me.tabConnection.TabIndex = 2
        Me.tabConnection.Text = "Connection"
        '
        'btnClrErrLog
        '
        Me.btnClrErrLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClrErrLog.Location = New System.Drawing.Point(726, 159)
        Me.btnClrErrLog.Name = "btnClrErrLog"
        Me.btnClrErrLog.Size = New System.Drawing.Size(88, 23)
        Me.btnClrErrLog.TabIndex = 5
        Me.btnClrErrLog.Text = "Clear Error Log"
        Me.btnClrErrLog.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.lstErrors)
        Me.GroupBox5.Location = New System.Drawing.Point(334, 179)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(486, 226)
        Me.GroupBox5.TabIndex = 4
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Error Log"
        '
        'lstErrors
        '
        Me.lstErrors.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstErrors.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstErrors.FormattingEnabled = True
        Me.lstErrors.HorizontalScrollbar = True
        Me.lstErrors.Location = New System.Drawing.Point(6, 19)
        Me.lstErrors.Name = "lstErrors"
        Me.lstErrors.Size = New System.Drawing.Size(474, 197)
        Me.lstErrors.TabIndex = 0
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.lstServerResponses)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 179)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(320, 226)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Server Log"
        '
        'lstServerResponses
        '
        Me.lstServerResponses.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstServerResponses.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstServerResponses.FormattingEnabled = True
        Me.lstServerResponses.HorizontalScrollbar = True
        Me.lstServerResponses.Location = New System.Drawing.Point(6, 19)
        Me.lstServerResponses.Name = "lstServerResponses"
        Me.lstServerResponses.Size = New System.Drawing.Size(308, 197)
        Me.lstServerResponses.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.labelAccountNum)
        Me.GroupBox2.Controls.Add(Me.labelServerTime)
        Me.GroupBox2.Controls.Add(Me.labelConnectionStatus)
        Me.GroupBox2.Location = New System.Drawing.Point(371, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(357, 65)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Connection Status"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(53, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Acct. Num :"
        '
        'labelAccountNum
        '
        Me.labelAccountNum.AutoSize = True
        Me.labelAccountNum.Location = New System.Drawing.Point(126, 46)
        Me.labelAccountNum.Name = "labelAccountNum"
        Me.labelAccountNum.Size = New System.Drawing.Size(0, 13)
        Me.labelAccountNum.TabIndex = 2
        '
        'labelServerTime
        '
        Me.labelServerTime.AutoSize = True
        Me.labelServerTime.Location = New System.Drawing.Point(66, 46)
        Me.labelServerTime.Name = "labelServerTime"
        Me.labelServerTime.Size = New System.Drawing.Size(0, 13)
        Me.labelServerTime.TabIndex = 1
        '
        'labelConnectionStatus
        '
        Me.labelConnectionStatus.AutoSize = True
        Me.labelConnectionStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelConnectionStatus.Location = New System.Drawing.Point(116, 16)
        Me.labelConnectionStatus.MaximumSize = New System.Drawing.Size(200, 30)
        Me.labelConnectionStatus.Name = "labelConnectionStatus"
        Me.labelConnectionStatus.Size = New System.Drawing.Size(136, 24)
        Me.labelConnectionStatus.TabIndex = 0
        Me.labelConnectionStatus.Text = "disconnected"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmbboxLogLevel)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.btnConnect)
        Me.GroupBox1.Controls.Add(Me.btnDisconnect)
        Me.GroupBox1.Controls.Add(Me.txtboxClientID)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtboxPort)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.txtboxIPAddress)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(359, 170)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Connection Settings"
        '
        'cmbboxLogLevel
        '
        Me.cmbboxLogLevel.FormattingEnabled = True
        Me.cmbboxLogLevel.Items.AddRange(New Object() {"System", "Error", "Warning", "Information", "Detail"})
        Me.cmbboxLogLevel.Location = New System.Drawing.Point(238, 34)
        Me.cmbboxLogLevel.Name = "cmbboxLogLevel"
        Me.cmbboxLogLevel.Size = New System.Drawing.Size(82, 21)
        Me.cmbboxLogLevel.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(236, 18)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(54, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Log Level"
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(9, 125)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(109, 36)
        Me.btnConnect.TabIndex = 6
        Me.btnConnect.Text = "Connect"
        Me.btnConnect.UseVisualStyleBackColor = True
        '
        'btnDisconnect
        '
        Me.btnDisconnect.Location = New System.Drawing.Point(220, 125)
        Me.btnDisconnect.Name = "btnDisconnect"
        Me.btnDisconnect.Size = New System.Drawing.Size(100, 36)
        Me.btnDisconnect.TabIndex = 7
        Me.btnDisconnect.Text = "Disconnect"
        Me.btnDisconnect.UseVisualStyleBackColor = True
        '
        'txtboxClientID
        '
        Me.txtboxClientID.Location = New System.Drawing.Point(156, 86)
        Me.txtboxClientID.Name = "txtboxClientID"
        Me.txtboxClientID.Size = New System.Drawing.Size(100, 20)
        Me.txtboxClientID.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(153, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Client ID"
        '
        'txtboxPort
        '
        Me.txtboxPort.Location = New System.Drawing.Point(9, 87)
        Me.txtboxPort.Name = "txtboxPort"
        Me.txtboxPort.Size = New System.Drawing.Size(109, 20)
        Me.txtboxPort.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Port number"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(187, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "IP Address (leave empty for local host)"
        '
        'txtboxIPAddress
        '
        Me.txtboxIPAddress.Location = New System.Drawing.Point(9, 35)
        Me.txtboxIPAddress.Name = "txtboxIPAddress"
        Me.txtboxIPAddress.Size = New System.Drawing.Size(184, 20)
        Me.txtboxIPAddress.TabIndex = 0
        '
        'tabAccount
        '
        Me.tabAccount.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabAccount.Controls.Add(Me.tabctrlAccount)
        Me.tabAccount.Controls.Add(Me.grpAcctSub)
        Me.tabAccount.Location = New System.Drawing.Point(4, 22)
        Me.tabAccount.Name = "tabAccount"
        Me.tabAccount.Padding = New System.Windows.Forms.Padding(3)
        Me.tabAccount.Size = New System.Drawing.Size(820, 408)
        Me.tabAccount.TabIndex = 1
        Me.tabAccount.Text = "Account"
        '
        'tabctrlAccount
        '
        Me.tabctrlAccount.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabctrlAccount.Controls.Add(Me.tabAcctSummary)
        Me.tabctrlAccount.Controls.Add(Me.tabPortfolio)
        Me.tabctrlAccount.Location = New System.Drawing.Point(8, 106)
        Me.tabctrlAccount.Name = "tabctrlAccount"
        Me.tabctrlAccount.SelectedIndex = 0
        Me.tabctrlAccount.Size = New System.Drawing.Size(813, 299)
        Me.tabctrlAccount.TabIndex = 1
        '
        'tabAcctSummary
        '
        Me.tabAcctSummary.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabAcctSummary.Controls.Add(Me.gridvwAcctSummary)
        Me.tabAcctSummary.Location = New System.Drawing.Point(4, 22)
        Me.tabAcctSummary.Name = "tabAcctSummary"
        Me.tabAcctSummary.Padding = New System.Windows.Forms.Padding(3)
        Me.tabAcctSummary.Size = New System.Drawing.Size(805, 273)
        Me.tabAcctSummary.TabIndex = 0
        Me.tabAcctSummary.Text = "Summary"
        '
        'gridvwAcctSummary
        '
        Me.gridvwAcctSummary.AllowUserToAddRows = False
        Me.gridvwAcctSummary.AllowUserToDeleteRows = False
        Me.gridvwAcctSummary.AllowUserToOrderColumns = True
        Me.gridvwAcctSummary.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridvwAcctSummary.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.gridvwAcctSummary.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.gridvwAcctSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridvwAcctSummary.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.gridvwAcctSummary.Location = New System.Drawing.Point(0, 0)
        Me.gridvwAcctSummary.Name = "gridvwAcctSummary"
        Me.gridvwAcctSummary.ReadOnly = True
        Me.gridvwAcctSummary.Size = New System.Drawing.Size(805, 273)
        Me.gridvwAcctSummary.TabIndex = 0
        '
        'tabPortfolio
        '
        Me.tabPortfolio.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabPortfolio.Controls.Add(Me.gridvwPortfolio)
        Me.tabPortfolio.Location = New System.Drawing.Point(4, 22)
        Me.tabPortfolio.Name = "tabPortfolio"
        Me.tabPortfolio.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPortfolio.Size = New System.Drawing.Size(805, 273)
        Me.tabPortfolio.TabIndex = 1
        Me.tabPortfolio.Text = "Portfolio"
        '
        'gridvwPortfolio
        '
        Me.gridvwPortfolio.AllowUserToAddRows = False
        Me.gridvwPortfolio.AllowUserToDeleteRows = False
        Me.gridvwPortfolio.AllowUserToOrderColumns = True
        Me.gridvwPortfolio.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridvwPortfolio.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.gridvwPortfolio.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.gridvwPortfolio.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridvwPortfolio.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.gridvwPortfolio.Location = New System.Drawing.Point(0, 0)
        Me.gridvwPortfolio.Name = "gridvwPortfolio"
        Me.gridvwPortfolio.ReadOnly = True
        Me.gridvwPortfolio.Size = New System.Drawing.Size(805, 273)
        Me.gridvwPortfolio.TabIndex = 0
        '
        'grpAcctSub
        '
        Me.grpAcctSub.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpAcctSub.Controls.Add(Me.labelAcctTime)
        Me.grpAcctSub.Controls.Add(Me.Label6)
        Me.grpAcctSub.Controls.Add(Me.btnUnsubscribe)
        Me.grpAcctSub.Controls.Add(Me.btnSubscribe)
        Me.grpAcctSub.Controls.Add(Me.labelAccountNum2)
        Me.grpAcctSub.Controls.Add(Me.Label5)
        Me.grpAcctSub.Location = New System.Drawing.Point(8, 6)
        Me.grpAcctSub.Name = "grpAcctSub"
        Me.grpAcctSub.Size = New System.Drawing.Size(802, 94)
        Me.grpAcctSub.TabIndex = 0
        Me.grpAcctSub.TabStop = False
        Me.grpAcctSub.Text = "Subscription"
        '
        'labelAcctTime
        '
        Me.labelAcctTime.AutoSize = True
        Me.labelAcctTime.Location = New System.Drawing.Point(95, 63)
        Me.labelAcctTime.Name = "labelAcctTime"
        Me.labelAcctTime.Size = New System.Drawing.Size(0, 13)
        Me.labelAcctTime.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(50, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(36, 13)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Time :"
        '
        'btnUnsubscribe
        '
        Me.btnUnsubscribe.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnUnsubscribe.Location = New System.Drawing.Point(682, 19)
        Me.btnUnsubscribe.Name = "btnUnsubscribe"
        Me.btnUnsubscribe.Size = New System.Drawing.Size(114, 30)
        Me.btnUnsubscribe.TabIndex = 3
        Me.btnUnsubscribe.Text = "Unsubscribe"
        Me.btnUnsubscribe.UseVisualStyleBackColor = True
        '
        'btnSubscribe
        '
        Me.btnSubscribe.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnSubscribe.Location = New System.Drawing.Point(524, 19)
        Me.btnSubscribe.Name = "btnSubscribe"
        Me.btnSubscribe.Size = New System.Drawing.Size(121, 30)
        Me.btnSubscribe.TabIndex = 2
        Me.btnSubscribe.Text = "Subscribe"
        Me.btnSubscribe.UseVisualStyleBackColor = True
        '
        'labelAccountNum2
        '
        Me.labelAccountNum2.AutoSize = True
        Me.labelAccountNum2.Location = New System.Drawing.Point(92, 32)
        Me.labelAccountNum2.Name = "labelAccountNum2"
        Me.labelAccountNum2.Size = New System.Drawing.Size(0, 13)
        Me.labelAccountNum2.TabIndex = 1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(23, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Acct. Num :"
        '
        'tabOrders
        '
        Me.tabOrders.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabOrders.Controls.Add(Me.tabctrlOrders)
        Me.tabOrders.Controls.Add(Me.grpOrderDesc)
        Me.tabOrders.Location = New System.Drawing.Point(4, 22)
        Me.tabOrders.Name = "tabOrders"
        Me.tabOrders.Size = New System.Drawing.Size(820, 408)
        Me.tabOrders.TabIndex = 3
        Me.tabOrders.Text = "Orders"
        '
        'tabctrlOrders
        '
        Me.tabctrlOrders.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabctrlOrders.Controls.Add(Me.tabOpenOrders)
        Me.tabctrlOrders.Controls.Add(Me.tabOrderStatus)
        Me.tabctrlOrders.Location = New System.Drawing.Point(9, 220)
        Me.tabctrlOrders.Name = "tabctrlOrders"
        Me.tabctrlOrders.SelectedIndex = 0
        Me.tabctrlOrders.Size = New System.Drawing.Size(801, 178)
        Me.tabctrlOrders.TabIndex = 1
        '
        'tabOpenOrders
        '
        Me.tabOpenOrders.Controls.Add(Me.gridvwOpenOrders)
        Me.tabOpenOrders.Location = New System.Drawing.Point(4, 22)
        Me.tabOpenOrders.Name = "tabOpenOrders"
        Me.tabOpenOrders.Padding = New System.Windows.Forms.Padding(3)
        Me.tabOpenOrders.Size = New System.Drawing.Size(793, 152)
        Me.tabOpenOrders.TabIndex = 0
        Me.tabOpenOrders.Text = "Open Orders"
        Me.tabOpenOrders.UseVisualStyleBackColor = True
        '
        'gridvwOpenOrders
        '
        Me.gridvwOpenOrders.AllowUserToAddRows = False
        Me.gridvwOpenOrders.AllowUserToDeleteRows = False
        Me.gridvwOpenOrders.AllowUserToOrderColumns = True
        Me.gridvwOpenOrders.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridvwOpenOrders.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.gridvwOpenOrders.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.gridvwOpenOrders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridvwOpenOrders.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.gridvwOpenOrders.Location = New System.Drawing.Point(0, 0)
        Me.gridvwOpenOrders.Name = "gridvwOpenOrders"
        Me.gridvwOpenOrders.ReadOnly = True
        Me.gridvwOpenOrders.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.gridvwOpenOrders.ShowEditingIcon = False
        Me.gridvwOpenOrders.Size = New System.Drawing.Size(793, 152)
        Me.gridvwOpenOrders.TabIndex = 0
        '
        'tabOrderStatus
        '
        Me.tabOrderStatus.AutoScroll = True
        Me.tabOrderStatus.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabOrderStatus.Controls.Add(Me.gridvwOrderStatus)
        Me.tabOrderStatus.Location = New System.Drawing.Point(4, 22)
        Me.tabOrderStatus.Name = "tabOrderStatus"
        Me.tabOrderStatus.Padding = New System.Windows.Forms.Padding(3)
        Me.tabOrderStatus.Size = New System.Drawing.Size(793, 152)
        Me.tabOrderStatus.TabIndex = 1
        Me.tabOrderStatus.Text = "Order Status"
        '
        'gridvwOrderStatus
        '
        Me.gridvwOrderStatus.AllowUserToAddRows = False
        Me.gridvwOrderStatus.AllowUserToDeleteRows = False
        Me.gridvwOrderStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridvwOrderStatus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.gridvwOrderStatus.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.gridvwOrderStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridvwOrderStatus.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.gridvwOrderStatus.Location = New System.Drawing.Point(0, 0)
        Me.gridvwOrderStatus.Name = "gridvwOrderStatus"
        Me.gridvwOrderStatus.ReadOnly = True
        Me.gridvwOrderStatus.ShowEditingIcon = False
        Me.gridvwOrderStatus.Size = New System.Drawing.Size(793, 152)
        Me.gridvwOrderStatus.TabIndex = 0
        '
        'grpOrderDesc
        '
        Me.grpOrderDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpOrderDesc.Controls.Add(Me.GroupBox9)
        Me.grpOrderDesc.Controls.Add(Me.GroupBox8)
        Me.grpOrderDesc.Controls.Add(Me.Label12)
        Me.grpOrderDesc.Controls.Add(Me.Label11)
        Me.grpOrderDesc.Controls.Add(Me.Label10)
        Me.grpOrderDesc.Controls.Add(Me.Label9)
        Me.grpOrderDesc.Controls.Add(Me.Label8)
        Me.grpOrderDesc.Controls.Add(Me.txtboxTicker)
        Me.grpOrderDesc.Controls.Add(Me.GroupBox7)
        Me.grpOrderDesc.Location = New System.Drawing.Point(8, 3)
        Me.grpOrderDesc.Name = "grpOrderDesc"
        Me.grpOrderDesc.Size = New System.Drawing.Size(802, 210)
        Me.grpOrderDesc.TabIndex = 0
        Me.grpOrderDesc.TabStop = False
        Me.grpOrderDesc.Text = "Order Description"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.btnClearOrders)
        Me.GroupBox9.Controls.Add(Me.btnGlobalCancel)
        Me.GroupBox9.Controls.Add(Me.btnWhatIf)
        Me.GroupBox9.Controls.Add(Me.btnPlaceModifyOrder)
        Me.GroupBox9.Location = New System.Drawing.Point(429, 20)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(367, 184)
        Me.GroupBox9.TabIndex = 8
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Place Order"
        '
        'btnClearOrders
        '
        Me.btnClearOrders.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearOrders.Location = New System.Drawing.Point(267, 133)
        Me.btnClearOrders.Name = "btnClearOrders"
        Me.btnClearOrders.Size = New System.Drawing.Size(85, 36)
        Me.btnClearOrders.TabIndex = 1
        Me.btnClearOrders.Text = "Clear Orders"
        Me.btnClearOrders.UseVisualStyleBackColor = True
        '
        'btnGlobalCancel
        '
        Me.btnGlobalCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGlobalCancel.ForeColor = System.Drawing.Color.Red
        Me.btnGlobalCancel.Location = New System.Drawing.Point(17, 133)
        Me.btnGlobalCancel.Name = "btnGlobalCancel"
        Me.btnGlobalCancel.Size = New System.Drawing.Size(144, 36)
        Me.btnGlobalCancel.TabIndex = 3
        Me.btnGlobalCancel.Text = "Global Cancel!!"
        Me.btnGlobalCancel.UseVisualStyleBackColor = True
        '
        'btnWhatIf
        '
        Me.btnWhatIf.Location = New System.Drawing.Point(17, 31)
        Me.btnWhatIf.Name = "btnWhatIf"
        Me.btnWhatIf.Size = New System.Drawing.Size(144, 36)
        Me.btnWhatIf.TabIndex = 1
        Me.btnWhatIf.Text = "What-If"
        Me.btnWhatIf.UseVisualStyleBackColor = True
        '
        'btnPlaceModifyOrder
        '
        Me.btnPlaceModifyOrder.Location = New System.Drawing.Point(208, 31)
        Me.btnPlaceModifyOrder.Name = "btnPlaceModifyOrder"
        Me.btnPlaceModifyOrder.Size = New System.Drawing.Size(144, 36)
        Me.btnPlaceModifyOrder.TabIndex = 0
        Me.btnPlaceModifyOrder.Text = "Place / Modify Order"
        Me.btnPlaceModifyOrder.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.txtboxOrderID)
        Me.GroupBox8.Controls.Add(Me.Label17)
        Me.GroupBox8.Controls.Add(Me.btnExtOrderAttr)
        Me.GroupBox8.Controls.Add(Me.txtboxPrice)
        Me.GroupBox8.Controls.Add(Me.txtboxOrderType)
        Me.GroupBox8.Controls.Add(Me.txtboxQuantity)
        Me.GroupBox8.Controls.Add(Me.txtboxAction)
        Me.GroupBox8.Controls.Add(Me.Label16)
        Me.GroupBox8.Controls.Add(Me.Label15)
        Me.GroupBox8.Controls.Add(Me.Label14)
        Me.GroupBox8.Controls.Add(Me.Label13)
        Me.GroupBox8.Location = New System.Drawing.Point(220, 20)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(203, 183)
        Me.GroupBox8.TabIndex = 7
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Order"
        '
        'txtboxOrderID
        '
        Me.txtboxOrderID.Location = New System.Drawing.Point(102, 32)
        Me.txtboxOrderID.Name = "txtboxOrderID"
        Me.txtboxOrderID.Size = New System.Drawing.Size(65, 20)
        Me.txtboxOrderID.TabIndex = 10
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(98, 16)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(47, 13)
        Me.Label17.TabIndex = 9
        Me.Label17.Text = "Order ID"
        '
        'btnExtOrderAttr
        '
        Me.btnExtOrderAttr.Location = New System.Drawing.Point(101, 84)
        Me.btnExtOrderAttr.Name = "btnExtOrderAttr"
        Me.btnExtOrderAttr.Size = New System.Drawing.Size(80, 46)
        Me.btnExtOrderAttr.TabIndex = 8
        Me.btnExtOrderAttr.Text = "Ext. Order Attr"
        Me.btnExtOrderAttr.UseVisualStyleBackColor = True
        '
        'txtboxPrice
        '
        Me.txtboxPrice.Location = New System.Drawing.Point(10, 149)
        Me.txtboxPrice.Name = "txtboxPrice"
        Me.txtboxPrice.Size = New System.Drawing.Size(65, 20)
        Me.txtboxPrice.TabIndex = 7
        '
        'txtboxOrderType
        '
        Me.txtboxOrderType.Location = New System.Drawing.Point(10, 110)
        Me.txtboxOrderType.Name = "txtboxOrderType"
        Me.txtboxOrderType.Size = New System.Drawing.Size(65, 20)
        Me.txtboxOrderType.TabIndex = 6
        '
        'txtboxQuantity
        '
        Me.txtboxQuantity.Location = New System.Drawing.Point(10, 71)
        Me.txtboxQuantity.Name = "txtboxQuantity"
        Me.txtboxQuantity.Size = New System.Drawing.Size(65, 20)
        Me.txtboxQuantity.TabIndex = 5
        '
        'txtboxAction
        '
        Me.txtboxAction.Location = New System.Drawing.Point(10, 31)
        Me.txtboxAction.Name = "txtboxAction"
        Me.txtboxAction.Size = New System.Drawing.Size(65, 20)
        Me.txtboxAction.TabIndex = 4
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(7, 133)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(31, 13)
        Me.Label16.TabIndex = 3
        Me.Label16.Text = "Price"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(7, 94)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(60, 13)
        Me.Label15.TabIndex = 2
        Me.Label15.Text = "Order Type"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(7, 54)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(46, 13)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = "Quantity"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 15)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(37, 13)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Action"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(118, 35)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(49, 13)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "Currency"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(15, 153)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(92, 13)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "Primary Exchange"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(15, 114)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(55, 13)
        Me.Label10.TabIndex = 3
        Me.Label10.Text = "Exchange"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(15, 74)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(31, 13)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Type"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(15, 35)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 13)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Symbol"
        '
        'txtboxTicker
        '
        Me.txtboxTicker.Location = New System.Drawing.Point(18, 51)
        Me.txtboxTicker.Name = "txtboxTicker"
        Me.txtboxTicker.Size = New System.Drawing.Size(80, 20)
        Me.txtboxTicker.TabIndex = 0
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.btnExtTickerAttr)
        Me.GroupBox7.Controls.Add(Me.txtboxCurrency)
        Me.GroupBox7.Controls.Add(Me.txtboxPrimExch)
        Me.GroupBox7.Controls.Add(Me.txtboxExchange)
        Me.GroupBox7.Controls.Add(Me.txtboxType)
        Me.GroupBox7.Location = New System.Drawing.Point(6, 19)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(207, 184)
        Me.GroupBox7.TabIndex = 6
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Ticker"
        '
        'btnExtTickerAttr
        '
        Me.btnExtTickerAttr.Location = New System.Drawing.Point(112, 85)
        Me.btnExtTickerAttr.Name = "btnExtTickerAttr"
        Me.btnExtTickerAttr.Size = New System.Drawing.Size(80, 46)
        Me.btnExtTickerAttr.TabIndex = 7
        Me.btnExtTickerAttr.Text = "Ext. Ticker Attr"
        Me.btnExtTickerAttr.UseVisualStyleBackColor = True
        '
        'txtboxCurrency
        '
        Me.txtboxCurrency.Location = New System.Drawing.Point(112, 32)
        Me.txtboxCurrency.Name = "txtboxCurrency"
        Me.txtboxCurrency.Size = New System.Drawing.Size(80, 20)
        Me.txtboxCurrency.TabIndex = 10
        '
        'txtboxPrimExch
        '
        Me.txtboxPrimExch.Location = New System.Drawing.Point(12, 150)
        Me.txtboxPrimExch.Name = "txtboxPrimExch"
        Me.txtboxPrimExch.Size = New System.Drawing.Size(80, 20)
        Me.txtboxPrimExch.TabIndex = 9
        '
        'txtboxExchange
        '
        Me.txtboxExchange.Location = New System.Drawing.Point(12, 111)
        Me.txtboxExchange.Name = "txtboxExchange"
        Me.txtboxExchange.Size = New System.Drawing.Size(80, 20)
        Me.txtboxExchange.TabIndex = 8
        '
        'txtboxType
        '
        Me.txtboxType.Location = New System.Drawing.Point(12, 71)
        Me.txtboxType.Name = "txtboxType"
        Me.txtboxType.Size = New System.Drawing.Size(80, 20)
        Me.txtboxType.TabIndex = 7
        '
        'tabExecutions
        '
        Me.tabExecutions.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabExecutions.Controls.Add(Me.tabctrlExecutions)
        Me.tabExecutions.Controls.Add(Me.grpReqExeReports)
        Me.tabExecutions.Location = New System.Drawing.Point(4, 22)
        Me.tabExecutions.Name = "tabExecutions"
        Me.tabExecutions.Size = New System.Drawing.Size(820, 408)
        Me.tabExecutions.TabIndex = 4
        Me.tabExecutions.Text = "Executions"
        '
        'tabctrlExecutions
        '
        Me.tabctrlExecutions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabctrlExecutions.Controls.Add(Me.tabExecReport)
        Me.tabctrlExecutions.Location = New System.Drawing.Point(4, 126)
        Me.tabctrlExecutions.Name = "tabctrlExecutions"
        Me.tabctrlExecutions.SelectedIndex = 0
        Me.tabctrlExecutions.Size = New System.Drawing.Size(813, 279)
        Me.tabctrlExecutions.TabIndex = 1
        '
        'tabExecReport
        '
        Me.tabExecReport.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.tabExecReport.Controls.Add(Me.gridvwExecutionsReport)
        Me.tabExecReport.Location = New System.Drawing.Point(4, 22)
        Me.tabExecReport.Name = "tabExecReport"
        Me.tabExecReport.Padding = New System.Windows.Forms.Padding(3)
        Me.tabExecReport.Size = New System.Drawing.Size(805, 253)
        Me.tabExecReport.TabIndex = 0
        Me.tabExecReport.Text = "Report"
        '
        'gridvwExecutionsReport
        '
        Me.gridvwExecutionsReport.AllowUserToAddRows = False
        Me.gridvwExecutionsReport.AllowUserToDeleteRows = False
        Me.gridvwExecutionsReport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gridvwExecutionsReport.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.gridvwExecutionsReport.BackgroundColor = System.Drawing.SystemColors.ButtonFace
        Me.gridvwExecutionsReport.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridvwExecutionsReport.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.gridvwExecutionsReport.Location = New System.Drawing.Point(1, 0)
        Me.gridvwExecutionsReport.Name = "gridvwExecutionsReport"
        Me.gridvwExecutionsReport.ReadOnly = True
        Me.gridvwExecutionsReport.Size = New System.Drawing.Size(804, 253)
        Me.gridvwExecutionsReport.TabIndex = 0
        '
        'grpReqExeReports
        '
        Me.grpReqExeReports.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpReqExeReports.Controls.Add(Me.btnClrExecs)
        Me.grpReqExeReports.Controls.Add(Me.btnReqExecution)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecAction)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecExch)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecSecType)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecSym)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecTime)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecAcct)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecClientId)
        Me.grpReqExeReports.Controls.Add(Me.txtboxExecReqId)
        Me.grpReqExeReports.Controls.Add(Me.Label25)
        Me.grpReqExeReports.Controls.Add(Me.Label24)
        Me.grpReqExeReports.Controls.Add(Me.Label23)
        Me.grpReqExeReports.Controls.Add(Me.Label22)
        Me.grpReqExeReports.Controls.Add(Me.Label21)
        Me.grpReqExeReports.Controls.Add(Me.Label20)
        Me.grpReqExeReports.Controls.Add(Me.Label19)
        Me.grpReqExeReports.Controls.Add(Me.Label18)
        Me.grpReqExeReports.Location = New System.Drawing.Point(4, 4)
        Me.grpReqExeReports.Name = "grpReqExeReports"
        Me.grpReqExeReports.Size = New System.Drawing.Size(813, 115)
        Me.grpReqExeReports.TabIndex = 0
        Me.grpReqExeReports.TabStop = False
        Me.grpReqExeReports.Text = "Request Execution Report"
        '
        'btnClrExecs
        '
        Me.btnClrExecs.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClrExecs.Location = New System.Drawing.Point(725, 73)
        Me.btnClrExecs.Name = "btnClrExecs"
        Me.btnClrExecs.Size = New System.Drawing.Size(81, 36)
        Me.btnClrExecs.TabIndex = 17
        Me.btnClrExecs.Text = "Clear Executions"
        Me.btnClrExecs.UseVisualStyleBackColor = True
        '
        'btnReqExecution
        '
        Me.btnReqExecution.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReqExecution.Location = New System.Drawing.Point(643, 16)
        Me.btnReqExecution.Name = "btnReqExecution"
        Me.btnReqExecution.Size = New System.Drawing.Size(163, 36)
        Me.btnReqExecution.TabIndex = 16
        Me.btnReqExecution.Text = "Request Execution Report"
        Me.btnReqExecution.UseVisualStyleBackColor = True
        '
        'txtboxExecAction
        '
        Me.txtboxExecAction.Location = New System.Drawing.Point(382, 84)
        Me.txtboxExecAction.Name = "txtboxExecAction"
        Me.txtboxExecAction.Size = New System.Drawing.Size(79, 20)
        Me.txtboxExecAction.TabIndex = 15
        '
        'txtboxExecExch
        '
        Me.txtboxExecExch.Location = New System.Drawing.Point(382, 45)
        Me.txtboxExecExch.Name = "txtboxExecExch"
        Me.txtboxExecExch.Size = New System.Drawing.Size(79, 20)
        Me.txtboxExecExch.TabIndex = 14
        '
        'txtboxExecSecType
        '
        Me.txtboxExecSecType.Location = New System.Drawing.Point(260, 84)
        Me.txtboxExecSecType.Name = "txtboxExecSecType"
        Me.txtboxExecSecType.Size = New System.Drawing.Size(72, 20)
        Me.txtboxExecSecType.TabIndex = 13
        '
        'txtboxExecSym
        '
        Me.txtboxExecSym.Location = New System.Drawing.Point(260, 45)
        Me.txtboxExecSym.Name = "txtboxExecSym"
        Me.txtboxExecSym.Size = New System.Drawing.Size(72, 20)
        Me.txtboxExecSym.TabIndex = 12
        '
        'txtboxExecTime
        '
        Me.txtboxExecTime.Location = New System.Drawing.Point(125, 84)
        Me.txtboxExecTime.Name = "txtboxExecTime"
        Me.txtboxExecTime.Size = New System.Drawing.Size(89, 20)
        Me.txtboxExecTime.TabIndex = 11
        '
        'txtboxExecAcct
        '
        Me.txtboxExecAcct.Location = New System.Drawing.Point(125, 45)
        Me.txtboxExecAcct.Name = "txtboxExecAcct"
        Me.txtboxExecAcct.Size = New System.Drawing.Size(89, 20)
        Me.txtboxExecAcct.TabIndex = 10
        '
        'txtboxExecClientId
        '
        Me.txtboxExecClientId.Location = New System.Drawing.Point(9, 84)
        Me.txtboxExecClientId.Name = "txtboxExecClientId"
        Me.txtboxExecClientId.Size = New System.Drawing.Size(66, 20)
        Me.txtboxExecClientId.TabIndex = 9
        '
        'txtboxExecReqId
        '
        Me.txtboxExecReqId.Location = New System.Drawing.Point(9, 45)
        Me.txtboxExecReqId.Name = "txtboxExecReqId"
        Me.txtboxExecReqId.Size = New System.Drawing.Size(66, 20)
        Me.txtboxExecReqId.TabIndex = 8
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(379, 68)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(37, 13)
        Me.Label25.TabIndex = 7
        Me.Label25.Text = "Action"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(379, 29)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(55, 13)
        Me.Label24.TabIndex = 6
        Me.Label24.Text = "Exchange"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(257, 68)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(50, 13)
        Me.Label23.TabIndex = 5
        Me.Label23.Text = "SecType"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(257, 29)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(41, 13)
        Me.Label22.TabIndex = 4
        Me.Label22.Text = "Symbol"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(122, 68)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(30, 13)
        Me.Label21.TabIndex = 3
        Me.Label21.Text = "Time"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(122, 29)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(75, 13)
        Me.Label20.TabIndex = 2
        Me.Label20.Text = "Account Code"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(6, 68)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(45, 13)
        Me.Label19.TabIndex = 1
        Me.Label19.Text = "Client Id"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(6, 29)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(59, 13)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "Request Id"
        '
        'ctxtmenuCancelOrder
        '
        Me.ctxtmenuCancelOrder.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.ctxtmenuCancelOrder.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.ctxtmenuCancelOrder.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CancelOrderToolStripMenuItem})
        Me.ctxtmenuCancelOrder.Name = "ctxtmenuCancelOrder"
        Me.ctxtmenuCancelOrder.Size = New System.Drawing.Size(147, 26)
        Me.ctxtmenuCancelOrder.Text = "Cancel Order"
        '
        'CancelOrderToolStripMenuItem
        '
        Me.CancelOrderToolStripMenuItem.Name = "CancelOrderToolStripMenuItem"
        Me.CancelOrderToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.CancelOrderToolStripMenuItem.Text = "Cancel Order"
        '
        'dlgMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(826, 432)
        Me.Controls.Add(Me.tabControl)
        Me.Name = "dlgMain"
        Me.Text = "Interactive Broker TWS VB.NET version (by The Portfolio Trader)"
        Me.tabControl.ResumeLayout(False)
        Me.tabConnection.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.tabAccount.ResumeLayout(False)
        Me.tabctrlAccount.ResumeLayout(False)
        Me.tabAcctSummary.ResumeLayout(False)
        CType(Me.gridvwAcctSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabPortfolio.ResumeLayout(False)
        CType(Me.gridvwPortfolio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAcctSub.ResumeLayout(False)
        Me.grpAcctSub.PerformLayout()
        Me.tabOrders.ResumeLayout(False)
        Me.tabctrlOrders.ResumeLayout(False)
        Me.tabOpenOrders.ResumeLayout(False)
        CType(Me.gridvwOpenOrders, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabOrderStatus.ResumeLayout(False)
        CType(Me.gridvwOrderStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpOrderDesc.ResumeLayout(False)
        Me.grpOrderDesc.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.tabExecutions.ResumeLayout(False)
        Me.tabctrlExecutions.ResumeLayout(False)
        Me.tabExecReport.ResumeLayout(False)
        CType(Me.gridvwExecutionsReport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpReqExeReports.ResumeLayout(False)
        Me.grpReqExeReports.PerformLayout()
        Me.ctxtmenuCancelOrder.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabControl As System.Windows.Forms.TabControl
    Friend WithEvents tabAccount As System.Windows.Forms.TabPage
    Friend WithEvents grpAcctSub As System.Windows.Forms.GroupBox
    Friend WithEvents labelAccountNum2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents labelAcctTime As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnUnsubscribe As System.Windows.Forms.Button
    Friend WithEvents btnSubscribe As System.Windows.Forms.Button
    Friend WithEvents tabctrlAccount As System.Windows.Forms.TabControl
    Friend WithEvents tabAcctSummary As System.Windows.Forms.TabPage
    Friend WithEvents tabPortfolio As System.Windows.Forms.TabPage
    Friend WithEvents tabConnection As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbboxLogLevel As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents btnDisconnect As System.Windows.Forms.Button
    Friend WithEvents txtboxClientID As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtboxPort As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtboxIPAddress As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents lstErrors As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents lstServerResponses As System.Windows.Forms.ListBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents labelAccountNum As System.Windows.Forms.Label
    Friend WithEvents labelServerTime As System.Windows.Forms.Label
    Friend WithEvents labelConnectionStatus As System.Windows.Forms.Label
    Friend WithEvents gridvwAcctSummary As System.Windows.Forms.DataGridView
    Friend WithEvents gridvwPortfolio As System.Windows.Forms.DataGridView
    Friend WithEvents tabOrders As System.Windows.Forms.TabPage
    Friend WithEvents grpOrderDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtboxTicker As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExtTickerAttr As System.Windows.Forms.Button
    Friend WithEvents txtboxCurrency As System.Windows.Forms.TextBox
    Friend WithEvents txtboxPrimExch As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExchange As System.Windows.Forms.TextBox
    Friend WithEvents txtboxType As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents txtboxPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtboxOrderType As System.Windows.Forms.TextBox
    Friend WithEvents txtboxQuantity As System.Windows.Forms.TextBox
    Friend WithEvents txtboxAction As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnExtOrderAttr As System.Windows.Forms.Button
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents btnWhatIf As System.Windows.Forms.Button
    Friend WithEvents btnPlaceModifyOrder As System.Windows.Forms.Button
    Friend WithEvents txtboxOrderID As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents tabctrlOrders As System.Windows.Forms.TabControl
    Friend WithEvents tabOpenOrders As System.Windows.Forms.TabPage
    Friend WithEvents tabOrderStatus As System.Windows.Forms.TabPage
    Friend WithEvents gridvwOpenOrders As System.Windows.Forms.DataGridView
    Friend WithEvents btnGlobalCancel As System.Windows.Forms.Button
    Friend WithEvents ctxtmenuCancelOrder As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CancelOrderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents gridvwOrderStatus As System.Windows.Forms.DataGridView
    Friend WithEvents tabExecutions As System.Windows.Forms.TabPage
    Friend WithEvents grpReqExeReports As System.Windows.Forms.GroupBox
    Friend WithEvents txtboxExecAction As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecExch As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecSecType As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecSym As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecTime As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecAcct As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecClientId As System.Windows.Forms.TextBox
    Friend WithEvents txtboxExecReqId As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents btnReqExecution As System.Windows.Forms.Button
    Friend WithEvents tabctrlExecutions As System.Windows.Forms.TabControl
    Friend WithEvents tabExecReport As System.Windows.Forms.TabPage
    Friend WithEvents gridvwExecutionsReport As System.Windows.Forms.DataGridView
    Friend WithEvents btnClrExecs As System.Windows.Forms.Button
    Friend WithEvents btnClrErrLog As System.Windows.Forms.Button
    Friend WithEvents btnClearOrders As System.Windows.Forms.Button

End Class
