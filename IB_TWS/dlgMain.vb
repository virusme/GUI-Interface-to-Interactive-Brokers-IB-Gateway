'=================================================================
'
' Author:  Whacky - The Portfolio Trader
' version: 1.0, 2014
'=================================================================

Option Strict Off
Option Explicit On

Imports System.Resources


Friend Class dlgMain
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        '
        Tws1 = New Tws(Me)
    End Sub


    '================================================================================
    ' With Events
    '================================================================================
    Public WithEvents Tws1 As Tws
    'Public WithEvents myTimer As System.Threading.Timer

    '================================================================================
    ' Private Members
    '================================================================================
    ' data members
    Private m_utils As New Utils
    Private m_IBdata As New IBData
    
    Private m_faAccount, faError As Boolean

    Private m_accountName As String
    Private m_complete As Boolean
    Dim faErrorCodes(5) As Short
    Dim MKT_DEPTH_DATA_RESET As Integer = 317

    '================================================================================
    ' Private Methods
    '================================================================================
    Private Sub dlgMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call m_utils.init(Me)
        Call m_IBdata.init(m_utils)
    

        ' Server log level
        cmbboxLogLevel.Items.Add(("System"))
        cmbboxLogLevel.Items.Add(("Error"))
        cmbboxLogLevel.Items.Add(("Warning"))
        cmbboxLogLevel.Items.Add(("Information"))
        cmbboxLogLevel.Items.Add(("Detail"))
        'cmbboxLogLevel.SelectedIndex = 4 ' Dfault log level is DETAIL

        ' disable button and grouo controls
        btnDisconnect.Enabled = False
        grpAcctSub.Enabled = False
        grpOrderDesc.Enabled = False
        grpReqExeReports.Enabled = False
        
        faErrorCodes(0) = 503
        faErrorCodes(1) = 504
        faErrorCodes(2) = 505
        faErrorCodes(3) = 522
        faErrorCodes(4) = 1100
        faErrorCodes(5) = 321

    End Sub


    '================================================================================
    ' Button Events
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Connect this API client to the TWS instance
    '--------------------------------------------------------------------------------
    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        '
        Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, _
                    "Connecting to Tws using clientId " & txtboxClientID.Text & " ...")
        If String.IsNullOrEmpty(txtboxIPAddress.Text) Then
            txtboxIPAddress.Text = "127.0.0.1"
        End If
        ' connect
        Call Tws1.connect(txtboxIPAddress.Text, txtboxPort.Text, txtboxClientID.Text, False, "")
        If (Tws1.serverVersion() > 0) Then   ' connected
            Dim msg As String
            msg = "Connected to Tws server version " & Tws1.serverVersion() & _
                  " at " & Tws1.TwsConnectionTime()
            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, msg)
            '
            ' update connection status
            labelConnectionStatus.Text = "Connected"
            labelConnectionStatus.ForeColor = Color.Green
            ' Get accounts
            Call Tws1.reqManagedAccts()
            ' Sets the TWS server logging level
            Call Tws1.setServerLogLevel(cmbboxLogLevel.SelectedIndex + 1)
            ' write connection settings to file
            With m_utils
                .m_IBsettings.Tables("Connection").Rows(0).Delete()
                .m_IBsettings = .addToDatatable(.m_IBsettings, "Connection", .m_connColumns, _
                                                  {txtboxIPAddress.Text, txtboxPort.Text, txtboxClientID.Text, cmbboxLogLevel.SelectedIndex})
                Try
                    .m_IBsettings.WriteXml(.AppPath & .IBSETFILE)
                Catch ex As Exception
                    Call m_utils.addListItem(Utils.List_Types.ERRORS, ex.Message)
                End Try
            End With
            '--------------------------------------------------------------------------------
            ' Request all the API open orders that were placed by this client. Note the
            ' clientID with a client id of 0 returns its, plus the TWS TWS open orders. Once
            ' requested the TWS orders retain their API asociation.
            '--------------------------------------------------------------------------------
            'Call Tws1.reqOpenOrders()

            ' disable/enable button and group controls
            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            grpAcctSub.Enabled = True
            btnUnsubscribe.Enabled = False
            grpOrderDesc.Enabled = True
            grpReqExeReports.Enabled = True
           
        Else           ' NOT CONNECTED
            ' do nothing
        End If
        

    End Sub

    '--------------------------------------------------------------------------------
    ' Disconnect this API client from the TWS instance
    '--------------------------------------------------------------------------------
    Private Sub btnDisconnect_Click(sender As Object, e As EventArgs) Handles btnDisconnect.Click
        ' Unsubscribe from Account updates if not performed, it will go into a infinite loop
        ' and you will not be able to subscribe again without restarting the application
        btnUnsubscribe_Click(sender, e)
        '
        Call Tws1.disconnect()
        ' update connection status
        labelConnectionStatus.Text = "disconnected"
        labelConnectionStatus.ForeColor = Color.Black
        '
        labelAccountNum.Text = ""

        ' disable/enable controls
        btnConnect.Enabled = True
        btnDisconnect.Enabled = False
        grpAcctSub.Enabled = False
        grpOrderDesc.Enabled = False
        grpReqExeReports.Enabled = False
        
    End Sub

    '--------------------------------------------------------------------------------
    ' Subscribe to account data
    '--------------------------------------------------------------------------------
    Private Sub btnSubscribe_Click(sender As Object, e As EventArgs) Handles btnSubscribe.Click
        m_complete = False
        Call Tws1.reqAccountUpdates(True, m_accountName)
        'disable/enable button and group controls
        btnSubscribe.Enabled = False
        btnUnsubscribe.Enabled = True
        ' request executions
        btnReqExecution_Click(sender, e)

    End Sub
    '--------------------------------------------------------------------------------
    ' Unsubscribe to account data
    '--------------------------------------------------------------------------------
    Private Sub btnUnsubscribe_Click(sender As Object, e As EventArgs) Handles btnUnsubscribe.Click
        m_complete = True
        Call Tws1.reqAccountUpdates(False, m_accountName)
        'disable/enable controls
        btnSubscribe.Enabled = True
        btnUnsubscribe.Enabled = False
        
    End Sub

    '--------------------------------------------------------------------------------
    ' Extended Ticker Attributes
    '--------------------------------------------------------------------------------
    Private Sub btnExtTickerAttr_Click(sender As Object, e As EventArgs) Handles btnExtTickerAttr.Click
        m_IBdata.showDlg("ExtTickerAttr")
    End Sub
    '--------------------------------------------------------------------------------
    ' Extended Order Attributes
    '--------------------------------------------------------------------------------
    Private Sub btnExtOrderAttr_Click(sender As Object, e As EventArgs) Handles btnExtOrderAttr.Click
        m_IBdata.showDlg("ExtOrderAttr")
    End Sub

    '--------------------------------------------------------------------------------
    ' Place a new or modify an existing order
    '--------------------------------------------------------------------------------
    Private Sub btnWhatIf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnWhatIf.Click
        m_IBdata.placeOrderImpl(True)
    End Sub
    Private Sub btnPlaceModifyOrder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPlaceModifyOrder.Click
        m_IBdata.placeOrderImpl(False)
    End Sub
    Public Sub cmdPlaceOrderEx(ByRef orderId As Integer, ByRef m_contractInfo As IBApi.Contract, ByRef m_orderInfo As IBApi.Order)
        Call Tws1.placeOrderEx(orderId, m_contractInfo, m_orderInfo)
    End Sub

    '--------------------------------------------------------------------------------
    ' Request global cancel
    '--------------------------------------------------------------------------------
    Private Sub btnGlobalCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGlobalCancel.Click
        Call Tws1.reqGlobalCancel()
    End Sub

    '--------------------------------------------------------------------------------
    ' Right Click Context Menu for CANCEL ORDER
    '--------------------------------------------------------------------------------
    Private Sub gridvwOpenOrders_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles gridvwOpenOrders.CellMouseDown
        Dim rowClicked As String
        If e.Button = Windows.Forms.MouseButtons.Right Then
            rowClicked = e.RowIndex
            ' select the clicked row
            Try
                gridvwOpenOrders.CurrentCell = gridvwOpenOrders.Rows(e.RowIndex).Cells(e.ColumnIndex)
                gridvwOpenOrders.Rows(rowClicked).Selected = True
                gridvwOpenOrders.Focus()
                ctxtmenuCancelOrder.Show(Cursor.Position)
            Catch ex As Exception
                ' do nothing, this exception occurs if you click outside the DataGridView TABLE
            End Try
        End If
    End Sub

    '--------------------------------------------------------------------------------
    '  Cancel Order _Click in ContextMenu
    '--------------------------------------------------------------------------------
    Private Sub CancelOrderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancelOrderToolStripMenuItem.Click
        Dim rowClicked As String
        Dim orderID As String
        rowClicked = gridvwOpenOrders.CurrentRow.Index
        With m_utils
            orderID = .m_dataset.Tables("OpenOrders").Rows(rowClicked)("OrderID").ToString()
        End With
        ' send request
        Call Tws1.cancelOrder(orderID)
    End Sub

    '--------------------------------------------------------------------------------
    '  Request Execution report for the contract
    '--------------------------------------------------------------------------------
    Private Sub btnReqExecution_Click(sender As Object, e As EventArgs) Handles btnReqExecution.Click

        Call m_IBdata.reqExecInfo(txtboxExecReqId.Text, txtboxExecClientId.Text, txtboxExecAcct.Text, txtboxExecTime.Text, _
                                  txtboxExecSym.Text, txtboxExecSecType.Text, txtboxExecExch.Text, txtboxExecAction.Text)
    End Sub
    Public Sub clickReqExecution()
        Call m_IBdata.reqExecInfo(txtboxExecReqId.Text, txtboxExecClientId.Text, txtboxExecAcct.Text, txtboxExecTime.Text, _
                                  txtboxExecSym.Text, txtboxExecSecType.Text, txtboxExecExch.Text, txtboxExecAction.Text)
    End Sub
    Public Sub cmdreqExections(ByRef reqId As Integer, ByRef m_execFilter As IBApi.ExecutionFilter)
        Call Tws1.reqExecutionsEx(reqId, m_execFilter)
    End Sub




    '--------------------------------------------------------------------------------
    '  Click to clear executions report
    '--------------------------------------------------------------------------------
    Private Sub btnClrExecs_Click(sender As Object, e As EventArgs) Handles btnClrExecs.Click
        Call m_IBdata.clearExecutions()
    End Sub

    '--------------------------------------------------------------------------------
    '  Click to clear Error Log
    '--------------------------------------------------------------------------------
    Private Sub btnClrErrLog_Click(sender As Object, e As EventArgs) Handles btnClrErrLog.Click
        Call m_utils.addListItem(Utils.List_Types.ERRORS, Nothing, True, True)
    End Sub

    '--------------------------------------------------------------------------------
    '  Click to clear Orders
    '--------------------------------------------------------------------------------
    Private Sub btnClearOrders_Click(sender As Object, e As EventArgs) Handles btnClearOrders.Click
        Call m_IBdata.clearOrderStatus()
        Call m_IBdata.clearOpenOrders()
    End Sub

    '================================================================================
    ' TWS Events
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Notification that the TWS-API connection has been broken.
    '--------------------------------------------------------------------------------
    Private Sub Tws1_connectionClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tws1.OnConnectionClosed
        Call m_utils.addListItem(Utils.List_Types.ERRORS, "Connection to Tws has been closed")

        ' move into view
        lstErrors.TopIndex = lstErrors.Items.Count - 1
    End Sub

    '--------------------------------------------------------------------------------
    ' Current Time event - triggered by the reqCurrentTime() methods
    '--------------------------------------------------------------------------------
    Private Sub Tws1_currentTime(ByVal sender As Object, ByVal eventArgs As AxTWSLib._DTwsEvents_currentTimeEvent) Handles Tws1.OncurrentTime

        Dim displayString As String
        displayString = "current time = " & eventArgs.time

        Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, displayString)

        ' move into view
        lstServerResponses.TopIndex = lstServerResponses.Items.Count - 1
    End Sub

    '--------------------------------------------------------------------------------
    ' Notification of the FA managed accounts (comma delimited list of account codes)
    '--------------------------------------------------------------------------------
    Private Sub Tws1_managedAccounts(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_managedAccountsEvent) Handles Tws1.OnmanagedAccounts
        'Dim msg As String
        'msg = "Connected : The list of managed accounts are : [" & eventArgs.accountsList & "]"
        'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, msg)
        m_faAccount = True
        m_accountName = eventArgs.accountsList
        m_IBdata.updateAccountList(eventArgs.accountsList)
    End Sub

    '--------------------------------------------------------------------------------
    ' Notification of a server time update - triggered by the reqAcctUpdates() method
    '--------------------------------------------------------------------------------
    Private Sub Tws1_updateAccountTime(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_updateAccountTimeEvent) Handles Tws1.OnupdateAccountTime
        m_IBdata.updateAccountTime(eventArgs.timestamp)
    End Sub

    '--------------------------------------------------------------------------------
    ' Notification of an account proprty update - triggered by the reqAcctUpdates() method
    '--------------------------------------------------------------------------------
    Private Sub Tws1_updateAccountValue(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_updateAccountValueEvent) Handles Tws1.OnupdateAccountValue
        m_IBdata.updateAccountValue(eventArgs.key, eventArgs.value, eventArgs.currency, eventArgs.accountName)
    End Sub

    '--------------------------------------------------------------------------------
    ' Notification of an updated/new portfolio position - triggered by the reqAcctUpdates() method
    '--------------------------------------------------------------------------------
    Private Sub Tws1_updatePortfolioEx(ByVal sender As Object, ByVal eventArgs As AxTWSLib._DTwsEvents_updatePortfolioExEvent) Handles Tws1.OnupdatePortfolioEx
        m_IBdata.updatePortfolio(eventArgs.contract, eventArgs.position, eventArgs.marketPrice, eventArgs.marketValue, eventArgs.averageCost, eventArgs.unrealisedPNL, eventArgs.realisedPNL, eventArgs.accountName)
    End Sub

    '--------------------------------------------------------------------------------
    ' Returns the next valid order id upon connection - triggered by the connect() and
    ' reqNextValidId() methods
    '--------------------------------------------------------------------------------
    Private Sub Tws1_nextValidId(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_nextValidIdEvent) Handles Tws1.OnNextValidId
        m_IBdata.updateOrderID(eventArgs.Id)
    End Sub
    '--------------------------------------------------------------------------------
    ' Requests the next avaliable order id for placing an order
    '--------------------------------------------------------------------------------
    Public Sub cmdReqIds()
        Call Tws1.reqIds(1)
    End Sub

    '--------------------------------------------------------------------------------
    ' Returns the details for an open order - triggered by the reqOpenOrders() method
    '--------------------------------------------------------------------------------
    Private Sub Tws1_openOrderEx(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_openOrderExEvent) Handles Tws1.OnopenOrderEx

        Call m_IBdata.updateOpenOrders(eventArgs.contract, eventArgs.order, eventArgs.orderState, eventArgs.orderId)

    End Sub
    Private Sub Tws1_openOrderEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Tws1.OnopenOrderEnd

        Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "============= end =============")

        ' move into view
        lstServerResponses.TopIndex = lstServerResponses.Items.Count - 1

    End Sub

    '--------------------------------------------------------------------------------
    ' Notification of an updates order status - triggered by an order state change.
    '--------------------------------------------------------------------------------
    Private Sub Tws1_orderStatus(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_orderStatusEvent) Handles Tws1.OnorderStatus
        'Dim msg As String
        'msg = "order status: orderId=" & eventArgs.orderId & " client id=" & eventArgs.clientId & " permId=" & eventArgs.permId & _
        '      " status=" & eventArgs.status & " filled=" & eventArgs.filled & " remaining=" & eventArgs.remaining & _
        '      " avgFillPrice=" & eventArgs.avgFillPrice & " lastFillPrice=" & eventArgs.lastFillPrice & _
        '      " parentId=" & eventArgs.parentId & " whyHeld=" & eventArgs.whyHeld

        m_IBdata.updateOrderStatus(eventArgs)
        Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "Order Status triggered, orderId=" & eventArgs.orderId)

    End Sub

    '--------------------------------------------------------------------------------
    ' An order execution report. This event is triggered by the explicit request for
    ' execution reports reqExecutionDetials(), and also by order state changes method
    '--------------------------------------------------------------------------------
    Private Sub Tws1_execDetailsEx(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_execDetailsExEvent) Handles Tws1.OnexecDetailsEx

        m_IBdata.updateExecutions(eventArgs.reqId, eventArgs.contract, eventArgs.execution)

    End Sub
    Private Sub Tws1_execDetailsEnd(ByVal eventSender As Object, ByVal eventArgs As AxTWSLib._DTwsEvents_execDetailsEndEvent) Handles Tws1.OnexecDetailsEnd

        Dim reqId As Long
        reqId = eventArgs.reqId

        Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "reqId = " & reqId & " =============== end ===============")

        ' move into view
        lstServerResponses.TopIndex = lstServerResponses.Items.Count - 1

    End Sub


    '--------------------------------------------------------------------------------
    ' Notify the users of any API request processing errors and displays them in the
    ' server responses listbox
    '--------------------------------------------------------------------------------
    Private Sub Tws1_errMsg(ByVal eventSender As System.Object, ByVal eventArgs As AxTWSLib._DTwsEvents_errMsgEvent) Handles Tws1.OnErrMsg
        Dim msg As String
        Dim ctr As Short

        'If eventArgs.errorCode = 399 Then   ' Order status warning
        '    Dim orderStatus As New AxTWSLib._DTwsEvents_orderStatusEvent
        '    orderStatus.Status = eventArgs.errorMsg
        '    orderStatus.orderId = eventArgs.id
        '    orderStatus.filled = Nothing
        '    orderStatus.avgFillPrice = Nothing
        '    orderStatus.lastFillPrice = Nothing
        '    orderStatus.whyHeld = Nothing
        '    Call m_IBdata.updateOrderStatus(orderStatus)

        'Else
        msg = "Time: " & Now.Date.Day.ToString & "/" & Now.Date.Month.ToString & "/" & Now.Date.Year.ToString & " @ " & Now.TimeOfDay.ToString & _
          " | id: " & eventArgs.id & " | Error Code: " & eventArgs.errorCode & " | Error Msg: " & eventArgs.errorMsg
        Call m_utils.addListItem(Utils.List_Types.ERRORS, msg)

        ' move into view
        lstErrors.TopIndex = lstErrors.Items.Count - 1

        For ctr = 0 To 5
            If eventArgs.errorCode = faErrorCodes(ctr) Then faError = True
        Next ctr

        If eventArgs.errorCode = MKT_DEPTH_DATA_RESET Then
            'm_dlgMktDepth.clear()
        End If
        'End If
    End Sub

    '--------------------------------------------------------------------------------
    ' Form Closing Event Handler
    '--------------------------------------------------------------------------------
    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        btnDisconnect_Click(sender, e)
    End Sub

    
End Class
