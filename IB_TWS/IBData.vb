'=================================================================
'
' Author:  Whacky - The Portfolio Trader
' version: 1.0, 2014
'=================================================================

Option Explicit On
Option Strict Off

Friend Class IBData
    Implements IDisposable

    '================================================================================
    ' Private Members
    '================================================================================
    ' data members
    Private m_utils As Utils
    Private m_dlgMain As dlgMain
    Private m_dlgExtTickAtrr As New dlgExtTickerAttr
    Private m_dlgExtOrderAttr As New dlgExtOrderAttr

    Public m_contractInfo As IBApi.Contract
    Public m_orderInfo As IBApi.Order
    Private m_execFilter As IBApi.ExecutionFilter
    Private m_underComp As IBApi.UnderComp

    Public i_orderID As Integer

    ' ========================================================
    ' Public methods
    ' ========================================================
    '--------------------------------------------------------------------------------
    ' Class initializer. Make utilites available to this class
    '--------------------------------------------------------------------------------
    Public Sub init(ByVal utilities As Utils)

        ' modules
        m_utils = utilities
        m_dlgMain = m_utils.m_dlgMain

        m_dlgExtTickAtrr.init(Me)
        m_dlgExtOrderAttr.init(Me)

        ' read settings
        loadContractInfo()
        loadOrderInfo()

        m_contractInfo = New IBApi.Contract
        m_orderInfo = New IBApi.Order
        m_execFilter = New IBApi.ExecutionFilter
        m_underComp = New IBApi.UnderComp

    End Sub



    '================================================================================
    ' Updates
    '================================================================================
    '--------------------------------------------------------------------------------
    ' server clock time
    '--------------------------------------------------------------------------------
    Public Sub updateAccountTime(ByRef timeStamp As String)
        With m_dlgMain.labelAcctTime
            .Text = timeStamp
            .Font = New Font(m_dlgMain.labelAcctTime.Font, FontStyle.Bold)
        End With
    End Sub
    '--------------------------------------------------------------------------------
    ' account list
    '--------------------------------------------------------------------------------
    Public Sub updateAccountList(ByRef accountList As String)
        With m_dlgMain.labelAccountNum
            .Text = accountList
            .Font = New Font(m_dlgMain.labelAccountNum.Font, FontStyle.Bold)
        End With
        With m_dlgMain.labelAccountNum2
            .Text = accountList
            .Font = New Font(m_dlgMain.labelAccountNum.Font, FontStyle.Bold)
        End With
    End Sub

    '--------------------------------------------------------------------------------
    ' Updates a user account property
    '--------------------------------------------------------------------------------
    Public Sub updateAccountValue(ByRef key As String, ByRef val_Renamed As String, ByRef curency As String, ByRef accountName As String)
        'Dim msg As String
        Dim colvalues As String()
        Dim drow() As DataRow
        
        colvalues = Nothing

        'msg = "key=" & key & " value=" & val_Renamed & " currency=" & curency & " account=" & accountName
        'Call m_utils.addListItem(Utils.List_Types.ACCOUNT_DATA, msg)

        ' find if the key + curency + accountName already exist
        With m_utils
            drow = .findRowInDatatable(.m_dataset, "Account", {"Key", "Currency", "Account"}, {key, curency, accountName})
        End With
        '
        If drow Is Nothing Then ' row does not exist, hence add-row to datatable
            colvalues = {key, val_Renamed, curency, accountName}
            With m_utils
                .m_dataset = .addToDatatable(.m_dataset, "Account", .m_acctColumns, colvalues)
            End With

        Else
            If drow.Count = 0 Then  ' row does not exist, hence add-row to datatable
                colvalues = {key, val_Renamed, curency, accountName}
                With m_utils
                    .m_dataset = .addToDatatable(.m_dataset, "Account", .m_acctColumns, colvalues)
                End With

            ElseIf drow.Count = 1 Then  ' row exist, update the value
                drow(0).Item("Value") = val_Renamed
                ' show the table
                With m_utils
                    .m_dataset = .addToDatatable(.m_dataset, "Account", Nothing, Nothing)
                End With

            Else
                m_utils.addListItem(Utils.List_Types.ERRORS, "Unknow If-Else conidition in IBData::updateAccount")
            End If
        End If

    End Sub

    '--------------------------------------------------------------------------------
    ' Updates a portfolio position details
    '--------------------------------------------------------------------------------
    Public Sub updatePortfolio(ByVal contract As IBApi.Contract, ByRef position As Integer, ByRef marketPrice As Double, ByRef marketValue As Double, ByRef averageCost As Double, ByRef unrealizedPNL As Double, ByRef realizedPNL As Double, ByRef accountName As String)
        'Dim msg As String
        Dim drow() As DataRow
        Dim colvalues As String()
        colvalues = Nothing

        'With contract
        '    msg = "conId=" & .ConId & " symbol=" & .Symbol & " secType=" & .SecType & " expiry=" & .Expiry & " strike=" & .Strike _
        '    & " right=" & .Right & " multiplier=" & .Multiplier & " primaryExch=" & .PrimaryExch & " currency=" & .Currency _
        '    & " localSymbol=" & .LocalSymbol & " tradingClass=" & .TradingClass & " position=" & position & " mktPrice=" & marketPrice & " mktValue=" & marketValue _
        '    & " avgCost=" & averageCost & " unrealizedPNL=" & unrealizedPNL & " realizedPNL=" & realizedPNL & " account=" & accountName
        'End With
        'Call m_utils.addListItem(Utils.List_Types.PORTFOLIO_DATA, msg)

        ' find if row Symbol+sectype+Exch+Currency+Account already existss
        With m_utils
            drow = .findRowInDatatable(.m_dataset, "Portfolio", _
                                       {"Symbol", "SecType", "PrimaryExch", "Currency", "Account", "TradingClass", "ConId"}, _
                                       {contract.Symbol, contract.SecType, contract.PrimaryExch, contract.Currency, accountName, contract.TradingClass, contract.ConId})
        End With
        '
        If Not IsNothing(contract.Currency) Then
            If drow Is Nothing Then ' row does not exist, add row
                With contract
                    colvalues = {.Symbol, .SecType, .PrimaryExch, .Currency, .LocalSymbol, position, marketPrice, marketValue, averageCost, unrealizedPNL, _
                                  realizedPNL, .TradingClass, .ConId, .Expiry, .Strike, .Right, .Multiplier, accountName}
                End With
                With m_utils
                    .m_dataset = .addToDatatable(.m_dataset, "Portfolio", .m_portfColumns, colvalues)
                End With

            Else
                If drow.Count = 0 Then ' row does not exist, add row
                    With contract
                        colvalues = {.Symbol, .SecType, .PrimaryExch, .Currency, .LocalSymbol, position, marketPrice, marketValue, averageCost, unrealizedPNL, _
                                      realizedPNL, .TradingClass, .ConId, .Expiry, .Strike, .Right, .Multiplier, accountName}
                    End With
                    With m_utils
                        .m_dataset = .addToDatatable(.m_dataset, "Portfolio", .m_portfColumns, colvalues)
                    End With

                ElseIf drow.Count = 1 Then ' row exists, update row
                    With drow(0)
                        .Item("Position") = position
                        .Item("MktPrice") = marketPrice
                        .Item("MktValue") = marketValue
                        .Item("AvgCost") = averageCost
                        .Item("UnrealizedPNL") = unrealizedPNL
                        .Item("RealizedPNL") = realizedPNL
                    End With
                    ' show the table
                    With m_utils
                        .m_dataset = .addToDatatable(.m_dataset, "Portfolio", Nothing, Nothing)
                    End With

                Else
                    m_utils.addListItem(Utils.List_Types.ERRORS, "Unknow If-Else conidition in IBData::updatePortfolio")
                End If
            End If

        End If

    End Sub

    '--------------------------------------------------------------------------------
    ' Updates order ID
    '--------------------------------------------------------------------------------
    Public Sub updateOrderID(ByRef id As Integer)
        i_orderID = id
        m_dlgMain.txtboxOrderID.Text = i_orderID
    End Sub

    '--------------------------------------------------------------------------------
    ' Updates Open Orders
    '--------------------------------------------------------------------------------
    Public Sub updateOpenOrders(ByRef contract As IBApi.Contract, ByRef order As IBApi.Order, _
                                ByRef orderState As IBApi.OrderState, ByRef orderId As Integer)

        Dim colvalues As String()
        Dim drow() As DataRow
        Dim rowind As Integer
        colvalues = Nothing
        rowind = Nothing

        ' find row, if the orderID exists in OpenOrders
        With m_utils
            'drow = .m_dataset.Tables("OpenOrders").Rows.Find(order.OrderId)
            drow = .findRowInDatatable(.m_dataset, "OpenOrders", {"OrderID"}, {order.OrderId})
        End With
        If drow Is Nothing Then
            ' row does not exist, add a new openOrder
            colvalues = {contract.Symbol, _
                     contract.SecType, _
                     contract.Exchange, _
                     contract.PrimaryExch, _
                     contract.Currency, _
                     order.Action, _
                     order.TotalQuantity, _
                     order.OrderType, _
                     DblMaxStr(order.LmtPrice), _
                     orderState.Status, _
                     orderState.WarningText, _
                     order.ClientId, _
                     order.OrderId, _
                     contract.LocalSymbol}
            ''---
            '' currently not using below fields
            ''---
            '             {order.PermId, _
            '             contract.ConId, _
            '             order.Tif, _
            '             order.GoodTillDate, _
            '             order.GoodAfterTime, _
            '             order.OutsideRth, _
            '             DblMaxStr(order.AuxPrice), _
            '             order.OcaGroup, _
            '             order.OcaType, _
            '             order.OrderRef, _
            '             order.ParentId, _
            '             order.BlockOrder, _
            '             order.SweepToFill, _
            '             order.DisplaySize, _
            '             order.TriggerMethod, _
            '             order.Hidden, _
            '             order.OverridePercentageConstraints, _
            '             order.Rule80A, _
            '             order.AllOrNone, _
            '             IntMaxStr(order.MinQty), _
            '             DblMaxStr(order.PercentOffset), _
            '             DblMaxStr(order.TrailStopPrice), _
            '             DblMaxStr(order.TrailingPercent), _
            '             order.WhatIf, _
            '             order.NotHeld, _
            '             orderState.InitMargin, _
            '             orderState.MaintMargin, _
            '             orderState.EquityWithLoan, _
            '             DblMaxStr(orderState.Commission), _
            '             DblMaxStr(orderState.MinCommission), _
            '             DblMaxStr(orderState.MaxCommission), _
            '             orderState.CommissionCurrency, _
            '             contract.Expiry, _
            '             contract.Strike, _
            '             contract.Right, _
            '             contract.Multiplier, _
            '             contract.TradingClass, _
            '             contract.ComboLegsDescription}

            With m_utils
                .m_dataset = .addToDatatable(.m_dataset, "OpenOrders", .m_openColumns, colvalues)
            End With

            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "OpenOrderEx called, orderId=" & orderId)
            '' --- 
            '' Currently not using the below fields to display OpenOrders
            ''----
            '' Financial advisors only
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  faGroup=" & order.FaGroup)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  faProfile=" & order.FaProfile)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  faMethod=" & order.FaMethod)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  faPercentage=" & order.FaPercentage)

            '' Clearing info
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  account=" & order.Account)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  settlingFirm=" & order.SettlingFirm)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  clearingAccount=" & order.ClearingAccount)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  clearingIntent=" & order.ClearingIntent)

            '' Institutional orders only
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  openClose=" & order.OpenClose)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  origin=" & order.Origin)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  shortSaleSlot=" & order.ShortSaleSlot)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  designatedLocation=" & order.DesignatedLocation)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  exemptCode=" & order.ExemptCode)

            '' SMART routing only
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  discretionaryAmt=" & order.DiscretionaryAmt)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  eTradeOnly=" & order.ETradeOnly)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  firmQuoteOnly=" & order.FirmQuoteOnly)
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  nbboPriceCap=" & DblMaxStr(order.NbboPriceCap))
            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "  optOutSmartRouting=" & order.OptOutSmartRouting)

            'Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "===============================")

        Else
            ' row exists, do nothing (this is take care of "dual"event triggers that happens due to "openOrdersEx" trigger
            ' just show the table
            With m_utils
                .m_dataset = .addToDatatable(.m_dataset, "OpenOrders", Nothing, Nothing)
            End With
        End If



    End Sub

    '--------------------------------------------------------------------------------
    ' Update order status
    '--------------------------------------------------------------------------------
    Public Sub updateOrderStatus(ByRef orderStatus As System.Object)
        Dim colvalues As String()
        Dim row As Integer
        Dim drow() As DataRow
        Dim drow_or() As DataRow

        colvalues = Nothing
        row = -1
        '{"OrderID", "Symbol", "SecType", "Exchange",  "Currency", "Action", "Quantity", "OrderType", _
        '                                       "Price", "Status", "WarningText", "ClientID", "Filled", "Remaining", "AvgFillPrice", "LastFillPrice", "PermID", _
        '                                      "ParentID", "WhyHeld"}

        ' find the correct row from OrderStatus
        With m_utils
            drow = .findRowInDatatable(.m_dataset, "OrderStatus", {"OrderID"}, {orderStatus.orderId})
        End With
        If drow IsNot Nothing Then
            ' update the order status
            'With m_utils.m_dataset.Tables("OrderStatus")
            '    .Rows(row)("Status") = orderStatus.status
            '    .Rows(row)("Filled") = orderStatus.filled
            '    .Rows(row)("AvgFillPrice") = orderStatus.avgFillPrice
            '    .Rows(row)("LastFillPrice") = orderStatus.lastFillPrice
            '    .Rows(row)("WhyHeld") = orderStatus.whyHeld
            'End With
            drow(0)("Status") = orderStatus.status
            drow(0)("Filled") = orderStatus.filled
            drow(0)("AvgFillPrice") = orderStatus.avgFillPrice
            drow(0)("LastFillPrice") = orderStatus.lastFillPrice
            drow(0)("WhyHeld") = orderStatus.whyHeld
            '
            ' display table
            m_utils.addToDatatable(m_utils.m_dataset, "OrderStatus", Nothing, Nothing)

        ElseIf drow Is Nothing Then
            ' row does not exist, create a new row
            ' find the correct row-data from OpenOrders
            With m_utils
                drow_or = .findRowInDatatable(.m_dataset, "OpenOrders", {"OrderID"}, {orderStatus.orderId})
            End With
            ' copy from row to OrderStatus Table
            If drow_or IsNot Nothing Then
                With orderStatus
                    colvalues = {drow_or(0)("Symbol").ToString, _
                                 drow_or(0)("SecType").ToString, _
                                 drow_or(0)("Exchange").ToString, _
                                 drow_or(0)("Currency").ToString, _
                                 drow_or(0)("Action").ToString, _
                                 drow_or(0)("Quantity").ToString, _
                                 drow_or(0)("OrderType").ToString, _
                                 drow_or(0)("Price").ToString, _
                                 .status, _
                                 .filled, _
                                 .remaining, _
                                 .avgFillPrice, _
                                 .lastFillPrice, _
                                 .permId, _
                                 .parentId, _
                                 .whyHeld, _
                                 .orderId, _
                                 .clientId}
                End With
                With m_utils
                    .m_dataset = .addToDatatable(.m_dataset, "OrderStatus", .m_orderstatusColumns, colvalues)
                End With

            ElseIf drow_or Is Nothing Then
                Call m_utils.addListItem(Utils.List_Types.ERRORS, " order ID = " & orderStatus.orderId & " not found in OpenOrders")
            End If

        Else
            Call m_utils.addListItem(Utils.List_Types.ERRORS, " unknown condition in IBData:updateOrderStatus")

        End If
        '

    End Sub

    '--------------------------------------------------------------------------------
    ' Update exeuctions
    '--------------------------------------------------------------------------------
    Public Sub updateExecutions(ByRef reqId As String, ByRef contract As IBApi.Contract, ByRef execution As IBApi.Execution)
        Dim drow() As DataRow
        Dim colvalues As String()
        colvalues = Nothing

        ' find row from executions table
        With m_utils
            'drow = .findRowInDatatable(.m_dataset, "Executions", {"ExecID"}, {execution.ExecId})
            drow = .findRowInDatatable(.m_dataset, "Executions", {"Symbol", "SecType", "Currency", "Side", "OrderID"}, _
                                         {contract.Symbol, contract.SecType, contract.Currency, execution.Side, execution.OrderId})
        End With

        If drow Is Nothing Then    ' no row exists
            colvalues = {contract.Symbol, _
                         contract.SecType, _
                         contract.Exchange, _
                         contract.Currency, _
                         execution.Side, _
                         execution.Shares, _
                         execution.Price, _
                         execution.CumQty, _
                         execution.AvgPrice, _
                         execution.OrderId, _
                         execution.ClientId, _
                         execution.AcctNumber, _
                         contract.ConId, _
                         execution.ExecId, _
                         reqId, _
                         execution.OrderRef, _
                         contract.Expiry, _
                         contract.Strike, _
                         contract.Right, _
                         contract.Multiplier, _
                         contract.PrimaryExch, _
                         contract.LocalSymbol, _
                         contract.TradingClass, _
                         execution.PermId, _
                         execution.Time, _
                         execution.Exchange, _
                         execution.Liquidation, _
                         execution.EvRule, _
                         execution.EvMultiplier}

            With m_utils
                .m_dataset = .addToDatatable(.m_dataset, "Executions", .m_execColumns, colvalues)
            End With

            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "ExecutionDetails received, reqId=" & reqId)

        ElseIf drow IsNot Nothing Then      ' row exists update the row
            ' update the execution statuses
            drow(0)("Shares") = execution.Shares
            drow(0)("Price") = execution.Price
            drow(0)("CumQty") = execution.CumQty
            drow(0)("AvgPrice") = execution.AvgPrice
            drow(0)("Time") = execution.Time

            ' display table
            m_utils.addToDatatable(m_utils.m_dataset, "Executions", Nothing, Nothing)

            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "ExecutionDetails received, reqId=" & reqId)

        Else
            Call m_utils.addListItem(Utils.List_Types.ERRORS, "Unknown case encountered in IBData:updateExecutions")

        End If
    End Sub

    '--------------------------------------------------------------------------------
    ' Clear executions
    '--------------------------------------------------------------------------------
    Public Sub clearExecutions()
        m_utils.m_dataset.Tables("Executions").Clear()
    End Sub
    '--------------------------------------------------------------------------------
    ' Clear Open Orders
    '--------------------------------------------------------------------------------
    Public Sub clearOpenOrders()
        m_utils.m_dataset.Tables("OpenOrders").Clear()
    End Sub
    '--------------------------------------------------------------------------------
    ' Clear executions
    '--------------------------------------------------------------------------------
    Public Sub clearOrderStatus()
        m_utils.m_dataset.Tables("OrderStatus").Clear()
    End Sub

    '================================================================================
    ' Server requests
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Requests the next avaliable order id for placing an order
    '--------------------------------------------------------------------------------
    Public Sub reqValidId()
        Call m_dlgMain.cmdReqIds()
    End Sub

    '--------------------------------------------------------------------------------
    ' Place order request
    '--------------------------------------------------------------------------------
    Public Sub placeOrderImpl(ByRef whatif As Boolean)
        Dim savedWhatIf As Boolean
        ' populate contract info
        getContractInfo()
        ' populate order info
        getOrderInfo()
        '
        savedWhatIf = m_orderInfo.WhatIf()
        m_orderInfo.WhatIf = whatif
        ' place order
        Call m_dlgMain.cmdPlaceOrderEx(i_orderID, m_contractInfo, m_orderInfo)
        m_orderInfo.WhatIf = savedWhatIf
        ' get a new order ID
        reqValidId()
        ' store the settings
        writeContractInfo()
        writeOrderInfo()
    End Sub

    '--------------------------------------------------------------------------------
    ' Request Exec info
    '--------------------------------------------------------------------------------
    Public Sub reqExecInfo(ByRef reqId As String, ByRef clientId As String, ByRef acct As String, ByRef time As String, ByRef sym As String, _
                            ByRef sectype As String, ByRef exch As String, ByRef side As String)
        Dim m_reqId As Integer

        m_reqId = Text2Int(reqId)
        With m_execFilter
            .ClientId = Text2Int(clientId)
            .AcctCode = acct
            .Time = time
            .Symbol = sym
            .SecType = sectype
            .Exchange = exch
            .Side = side
        End With

        Call m_dlgMain.cmdreqExections(m_reqId, m_execFilter)

    End Sub

    '================================================================================
    ' Read, Write or Get
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Load Contract Info from settings file
    '--------------------------------------------------------------------------------
    Public Sub loadContractInfo()
        Dim dTKrow As DataRow

        dTKrow = m_utils.m_IBsettings.Tables("TickerAttributes").Rows(0)

        With m_dlgMain
            ' Ticker attributes
            .txtboxType.Text = dTKrow("StkType").ToString
            .txtboxExchange.Text = dTKrow("Exchange").ToString
            .txtboxPrimExch.Text = dTKrow("PrimExchange").ToString
            .txtboxCurrency.Text = dTKrow("Currency").ToString

            ' Order Attributes
            '.txtboxAction.Text = dOrdrow("Action").ToString
            '.txtboxQuantity.Text = dOrdrow("Quantity").ToString
            '.txtboxOrderType.Text = dOrdrow("OrderType").ToString
            '.txtboxPrice.Text = dOrdrow("Price").ToString

        End With

        ' Extended Ticker attributes
        With m_dlgExtTickAtrr
            .txtboxContractId.Text = dTKrow("ContractID").ToString
            .txtboxExpiry.Text = dTKrow("Expiry").ToString
            .txtboxStrike.Text = dTKrow("Strike").ToString
            .txtboxRight.Text = dTKrow("Right").ToString
            .txtboxMultiplier.Text = dTKrow("Multiplier").ToString
            .txtboxLocalSymbol.Text = dTKrow("LocalSymbol").ToString
            .txtboxTradingClass.Text = dTKrow("TradingClass").ToString
            .cmbboxIncExpired.SelectedIndex = CInt(dTKrow("IncExpired").ToString)
            .txtboxSecIdType.Text = dTKrow("SecIDType").ToString
            .txtboxSecId.Text = dTKrow("SecID").ToString
        End With

    End Sub
    '--------------------------------------------------------------------------------
    ' Load Order Info from settings file
    '--------------------------------------------------------------------------------
    Public Sub loadOrderInfo()
        Dim dOrdrow As DataRow

        dOrdrow = m_utils.m_IBsettings.Tables("OrderAttributes").Rows(0)

        With m_dlgMain

            ' Order Attributes
            '.txtboxAction.Text = dOrdrow("Action").ToString
            '.txtboxQuantity.Text = dOrdrow("Quantity").ToString
            '.txtboxOrderType.Text = dOrdrow("OrderType").ToString
            '.txtboxPrice.Text = dOrdrow("Price").ToString

        End With

        ' Extended Order Attributes
        With m_dlgExtOrderAttr
            .txtboxAux.Text = dOrdrow("Aux").ToString
            .txtboxGoodAftTime.Text = dOrdrow("GoodAfterTime").ToString
            .txtboxGoodTillDate.Text = dOrdrow("GoodTillDate").ToString
            .txtboxTimeInForce.Text = dOrdrow("TimeInForce").ToString
            .txtboxOCAGroup.Text = dOrdrow("OCAGroup").ToString
            .txtboxOCAType.Text = dOrdrow("OCAType").ToString
            .txtboxOrderAccount.Text = dOrdrow("Account").ToString
            .txtboxSetFirm.Text = dOrdrow("SettlingFirm").ToString
            .txtboxClearingAcct.Text = dOrdrow("ClearingAccount").ToString
            .txtboxClearingInt.Text = dOrdrow("ClearingIntent").ToString
            .txtboxOpenClose.Text = dOrdrow("OpenClose").ToString
            .txtboxOrigin.Text = dOrdrow("Origin").ToString
            .txtboxOrderRef.Text = dOrdrow("OrderRef").ToString
            .txtboxParentID.Text = dOrdrow("ParentID").ToString
            .txtboxTransmit.Text = dOrdrow("Transmit").ToString
            .txtboxBlkOrder.Text = dOrdrow("BlockOrder").ToString
            .txtboxSweepFill.Text = dOrdrow("SweepToFill").ToString
            .txtboxDesLoc.Text = dOrdrow("DesignatedLocation").ToString
            .txtboxExmptCde.Text = dOrdrow("ExemptCode").ToString
            .txtboxRule80A.Text = dOrdrow("Rule80A").ToString
            .txtboxTrailStpPrice.Text = dOrdrow("TrailStopPrice").ToString
            .txtboxTrailPcnt.Text = dOrdrow("TrailingPercent").ToString
            .txtboxAllNone.Text = dOrdrow("AllNone").ToString
            .txtboxMinQnty.Text = dOrdrow("MinQuantity").ToString
            .txtboxPcntOffset.Text = dOrdrow("PercentOffset").ToString
            .txtboxEExchOnly.Text = dOrdrow("ElectronicXonly").ToString
            .txtboxFirmQuote.Text = dOrdrow("FirmQuote").ToString
            .txtboxNBBOCap.Text = dOrdrow("NBBOCap").ToString
            .txtboxOverridePcnt.Text = dOrdrow("OverridePercent").ToString
            .txtboxDisplay.Text = dOrdrow("Display").ToString
            .txtboxTrigger.Text = dOrdrow("TriggerMethod").ToString
            .txtboxOutRTH.Text = dOrdrow("OutsideRTH").ToString
            .txtboxHidden.Text = dOrdrow("Hidden").ToString
            .txtboxDiscAmt.Text = dOrdrow("DiscretionaryAmt").ToString
            .txtboxShortSalesSlot.Text = dOrdrow("ShortSalesSlot").ToString
            .txtboxActiveStartTime.Text = dOrdrow("ActiveStartTime").ToString
            .txtboxActiveStopTime.Text = dOrdrow("ActiveStopTime").ToString
            .chkboxOptOutSMART.Checked = CBool(dOrdrow("OptOutSmart").ToString)
        End With

    End Sub

    '--------------------------------------------------------------------------------
    ' Read Ticker/Contract info
    '--------------------------------------------------------------------------------
    Private Sub getContractInfo()
        With m_dlgMain
            m_contractInfo.Symbol = .txtboxTicker.Text
            m_contractInfo.SecType = .txtboxType.Text
            m_contractInfo.Exchange = .txtboxExchange.Text
            m_contractInfo.PrimaryExch = .txtboxPrimExch.Text
            m_contractInfo.Currency = .txtboxCurrency.Text
        End With
        With m_dlgExtTickAtrr
            m_contractInfo.ConId = ival(.txtboxContractId.Text)
            m_contractInfo.Expiry = .txtboxExpiry.Text
            m_contractInfo.Strike = CDbl(.txtboxStrike.Text)
            m_contractInfo.Right = .txtboxRight.Text
            m_contractInfo.Multiplier = .txtboxMultiplier.Text
            m_contractInfo.LocalSymbol = .txtboxLocalSymbol.Text
            m_contractInfo.TradingClass = .txtboxTradingClass.Text
            m_contractInfo.IncludeExpired = CBool(.cmbboxIncExpired.SelectedIndex)
            m_contractInfo.SecIdType = .txtboxSecIdType.Text
            m_contractInfo.SecId = .txtboxSecId.Text
        End With
    End Sub
    '--------------------------------------------------------------------------------
    ' Read Order info
    '--------------------------------------------------------------------------------
    Private Sub getOrderInfo()
        With m_dlgMain
            m_orderInfo.Action = .txtboxAction.Text
            If Not String.IsNullOrEmpty(.txtboxQuantity.Text) Then
                m_orderInfo.TotalQuantity = CInt(.txtboxQuantity.Text)
            Else
                m_orderInfo.TotalQuantity = 0
            End If
            m_orderInfo.OrderType = .txtboxOrderType.Text
            m_orderInfo.LmtPrice = dval(.txtboxPrice.Text)
        End With
        With m_dlgExtOrderAttr
            m_orderInfo.AuxPrice = dval(.txtboxAux.Text)
            m_orderInfo.GoodAfterTime = .txtboxGoodAftTime.Text
            m_orderInfo.GoodTillDate = .txtboxGoodTillDate.Text
            m_orderInfo.Tif = .txtboxTimeInForce.Text
            m_orderInfo.ActiveStartTime = .txtboxActiveStartTime.Text
            m_orderInfo.ActiveStopTime = .txtboxActiveStopTime.Text
            m_orderInfo.OcaGroup = .txtboxOCAGroup.Text
            m_orderInfo.OcaType = ival(.txtboxOCAType.Text)
            m_orderInfo.OrderRef = .txtboxOrderRef.Text
            m_orderInfo.Transmit = bval(.txtboxTransmit.Text)
            m_orderInfo.ParentId = ival(.txtboxParentID.Text)
            m_orderInfo.BlockOrder = bval(.txtboxBlkOrder.Text)
            m_orderInfo.SweepToFill = bval(.txtboxSweepFill.Text)
            m_orderInfo.DisplaySize = ival(.txtboxDisplay.Text)
            If Not String.IsNullOrEmpty(.txtboxTrigger.Text) Then
                m_orderInfo.TriggerMethod = CInt(.txtboxTrigger.Text)
            Else
                m_orderInfo.TriggerMethod = 0
            End If
            m_orderInfo.OutsideRth = bval(.txtboxOutRTH.Text)
            m_orderInfo.Hidden = .txtboxHidden.Text
            m_orderInfo.OverridePercentageConstraints = bval(.txtboxOverridePcnt.Text)
            m_orderInfo.Rule80A = .txtboxRule80A.Text
            m_orderInfo.AllOrNone = bval(.txtboxAllNone.Text)
            m_orderInfo.MinQty = ival(.txtboxMinQnty.Text)
            m_orderInfo.PercentOffset = dval(.txtboxPcntOffset.Text)
            m_orderInfo.TrailStopPrice = dval(.txtboxTrailStpPrice.Text)
            m_orderInfo.TrailingPercent = dval(.txtboxTrailPcnt.Text)

            ' Institutional orders only
            m_orderInfo.OpenClose = .txtboxOpenClose.Text
            m_orderInfo.Origin = ival(.txtboxOrigin.Text)
            m_orderInfo.ShortSaleSlot = ival(.txtboxShortSalesSlot.Text)
            m_orderInfo.DesignatedLocation = .txtboxDesLoc.Text
            If Not String.IsNullOrEmpty(.txtboxExmptCde.Text) Then
                m_orderInfo.ExemptCode = ival(.txtboxExmptCde.Text)
            Else
                m_orderInfo.ExemptCode = ival("-1")
            End If

            'SMART routing only
            m_orderInfo.DiscretionaryAmt = dval(.txtboxDiscAmt.Text)
            m_orderInfo.ETradeOnly = bval(.txtboxEExchOnly.Text)
            m_orderInfo.FirmQuoteOnly = bval(.txtboxFirmQuote.Text)
            m_orderInfo.NbboPriceCap = dval(.txtboxNBBOCap.Text)
            m_orderInfo.OptOutSmartRouting = .chkboxOptOutSMART.Checked

            ' Clearing info
            m_orderInfo.Account = .txtboxOrderAccount.Text
            m_orderInfo.SettlingFirm = .txtboxSetFirm.Text
            m_orderInfo.ClearingAccount = .txtboxClearingAcct.Text
            m_orderInfo.ClearingIntent = .txtboxClearingInt.Text
            'm_orderInfo.Solicited = cbSolicited.Checked
        End With
    End Sub

    '--------------------------------------------------------------------------------
    ' Write Contract Info to settings file
    '--------------------------------------------------------------------------------
    Private Sub writeContractInfo()
        Dim dTKrow As DataRow

        dTKrow = m_utils.m_IBsettings.Tables("TickerAttributes").Rows(0)
        m_utils.m_IBsettings.Tables("TickerAttributes").Rows(0).Delete()

        With m_dlgMain
            ' Ticker attributes
            dTKrow("StkType") = .txtboxType.Text
            dTKrow("Exchange") = .txtboxExchange.Text
            dTKrow("PrimExchange") = .txtboxPrimExch.Text
            dTKrow("Currency") = .txtboxCurrency.Text

            ' Order Attributes
            '.txtboxAction.Text = dOrdrow("Action").ToString
            '.txtboxQuantity.Text = dOrdrow("Quantity").ToString
            '.txtboxOrderType.Text = dOrdrow("OrderType").ToString
            '.txtboxPrice.Text = dOrdrow("Price").ToString

        End With

        ' Extended Ticker attributes
        With m_dlgExtTickAtrr
            dTKrow("ContractID") = .txtboxContractId.Text
            dTKrow("Expiry") = .txtboxExpiry.Text
            dTKrow("Strike") = .txtboxStrike.Text
            dTKrow("Right") = .txtboxRight.Text
            dTKrow("Multiplier") = .txtboxMultiplier.Text
            dTKrow("LocalSymbol") = .txtboxLocalSymbol.Text
            dTKrow("TradingClass") = .txtboxTradingClass.Text
            dTKrow("IncExpired") = .cmbboxIncExpired.SelectedIndex
            dTKrow("SecIDType") = .txtboxSecIdType.Text
            dTKrow("SecID") = .txtboxSecId.Text
        End With
        m_utils.m_IBsettings.Tables("TickerAttributes").Rows.Add(dTKrow)
        'write to XML
        With m_utils
            Try
                .m_IBsettings.WriteXml(.AppPath & .IBSETFILE)
            Catch ex As Exception
                Call .addListItem(Utils.List_Types.ERRORS, ex.Message)
            End Try
        End With
    End Sub

    '--------------------------------------------------------------------------------
    ' Write Order Info to settings file
    '--------------------------------------------------------------------------------
    Private Sub writeOrderInfo()
        Dim dOrdrow As DataRow

        dOrdrow = m_utils.m_IBsettings.Tables("OrderAttributes").Rows(0)
        m_utils.m_IBsettings.Tables("OrderAttributes").Rows(0).Delete()

        With m_dlgMain

            ' Order Attributes
            '.txtboxAction.Text = dOrdrow("Action").ToString
            '.txtboxQuantity.Text = dOrdrow("Quantity").ToString
            '.txtboxOrderType.Text = dOrdrow("OrderType").ToString
            '.txtboxPrice.Text = dOrdrow("Price").ToString

        End With

        ' Extended Order Attributes
        With m_dlgExtOrderAttr
            dOrdrow("Aux") = .txtboxAux.Text
            dOrdrow("GoodAfterTime") = .txtboxGoodAftTime.Text
            dOrdrow("GoodTillDate") = .txtboxGoodTillDate.Text
            dOrdrow("TimeInForce") = .txtboxTimeInForce.Text
            dOrdrow("OCAGroup") = .txtboxOCAGroup.Text
            dOrdrow("OCAType") = .txtboxOCAType.Text
            dOrdrow("Account") = .txtboxOrderAccount.Text
            dOrdrow("SettlingFirm") = .txtboxSetFirm.Text
            dOrdrow("ClearingAccount") = .txtboxClearingAcct.Text
            dOrdrow("ClearingIntent") = .txtboxClearingInt.Text
            dOrdrow("OpenClose") = .txtboxOpenClose.Text
            dOrdrow("Origin") = .txtboxOrigin.Text
            dOrdrow("OrderRef") = .txtboxOrderRef.Text
            dOrdrow("ParentID") = .txtboxParentID.Text
            dOrdrow("Transmit") = .txtboxTransmit.Text
            dOrdrow("BlockOrder") = .txtboxBlkOrder.Text
            dOrdrow("SweepToFill") = .txtboxSweepFill.Text
            dOrdrow("DesignatedLocation") = .txtboxDesLoc.Text
            dOrdrow("ExemptCode") = .txtboxExmptCde.Text
            dOrdrow("Rule80A") = .txtboxRule80A.Text
            dOrdrow("TrailStopPrice") = .txtboxTrailStpPrice.Text
            dOrdrow("TrailingPercent") = .txtboxTrailPcnt.Text
            dOrdrow("AllNone") = .txtboxAllNone.Text
            dOrdrow("MinQuantity") = .txtboxMinQnty.Text
            dOrdrow("PercentOffset") = .txtboxPcntOffset.Text
            dOrdrow("ElectronicXonly") = .txtboxEExchOnly.Text
            dOrdrow("FirmQuote") = .txtboxFirmQuote.Text
            dOrdrow("NBBOCap") = .txtboxNBBOCap.Text
            dOrdrow("OverridePercent") = .txtboxOverridePcnt.Text
            dOrdrow("Display") = .txtboxDisplay.Text
            dOrdrow("TriggerMethod") = .txtboxTrigger.Text
            dOrdrow("OutsideRTH") = .txtboxOutRTH.Text
            dOrdrow("Hidden") = .txtboxHidden.Text
            dOrdrow("DiscretionaryAmt") = .txtboxDiscAmt.Text
            dOrdrow("ShortSalesSlot") = .txtboxShortSalesSlot.Text
            dOrdrow("ActiveStartTime") = .txtboxActiveStartTime.Text
            dOrdrow("ActiveStopTime") = .txtboxActiveStopTime.Text
            dOrdrow("OptOutSmart") = .chkboxOptOutSMART.Checked
        End With
        m_utils.m_IBsettings.Tables("OrderAttributes").Rows.Add(dOrdrow)
        'write to XML
        With m_utils
            Try
                .m_IBsettings.WriteXml(.AppPath & .IBSETFILE)
            Catch ex As Exception
                Call .addListItem(Utils.List_Types.ERRORS, ex.Message)
            End Try
        End With
    End Sub



    '================================================================================
    ' Show DialogBoxes
    '================================================================================
    Public Sub showDlg(ByRef boxName As String)
        Select Case boxName
            Case "ExtTickerAttr"
                m_dlgExtTickAtrr.ShowDialog()
                If m_dlgExtTickAtrr.b_save Then
                    Call getContractInfo()
                    Call writeContractInfo()
                Else
                    Call loadContractInfo()
                End If

            Case "ExtOrderAttr"
                m_dlgExtOrderAttr.ShowDialog()
                If m_dlgExtOrderAttr.b_save Then
                    Call getOrderInfo()
                    Call writeOrderInfo()
                Else
                    Call loadOrderInfo()
                End If

            Case Else
                Call m_utils.addListItem(Utils.List_Types.ERRORS, "Unknown case in IBData:showDlg")
        End Select
    End Sub



    '================================================================================
    ' Methods
    '================================================================================
    Private Function ivalStr(ByVal val As Long) As String
        If val = Int32.MaxValue Then
            ivalStr = ""
        Else
            ivalStr = val
        End If
    End Function
    Private Function dvalStr(ByVal val As Double) As String
        If val = Double.MaxValue Then
            dvalStr = ""
        Else
            dvalStr = val
        End If
    End Function
    Private Function bval(ByVal text As String) As Boolean
        If Len(text) = 0 Then
            bval = False
        Else
            bval = CBool(text)
        End If
    End Function
    Private Function ival(ByVal text As String) As Integer
        If Len(text) = 0 Then
            ival = Int32.MaxValue
        Else
            ival = CInt(text)
        End If
    End Function
    Private Function dval(ByVal text As String) As Double
        If Len(text) = 0 Then
            dval = Double.MaxValue
        Else
            dval = CDbl(text)
        End If
    End Function
    Private Function IntMaxStr(ByRef intVal As Integer) As String
        If intVal = Integer.MaxValue Then
            IntMaxStr = ""
        Else
            IntMaxStr = CStr(intVal)
        End If
    End Function
    Private Function DblMaxStr(ByRef dblVal As Double) As String
        If dblVal = Double.MaxValue Then
            DblMaxStr = ""
        Else
            DblMaxStr = CStr(dblVal)
        End If
    End Function
    Private Function Text2Int(ByRef text As String) As Integer
        If Len(text) <= 0 Then
            Text2Int = 0
        Else
            Text2Int = text
        End If
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                m_dlgExtOrderAttr.Dispose()
                m_dlgExtTickAtrr.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
