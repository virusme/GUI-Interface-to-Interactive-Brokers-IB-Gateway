Option Strict Off

'-------------------------------------------------------------------------------------
'
' Update - Container for all functions and procedures related to IBDate.Update
'
'-------------------------------------------------------------------------------------
Partial Friend Class IBData

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
        drow = findRowInDatatable(m_dataset, "Account", {"Key", "Currency", "Account"}, _
                                       {key, curency, accountName}, {"=", "=", "="})
        '
        If drow Is Nothing Then ' row does not exist, hence add-row to datatable
            colvalues = {key, val_Renamed, curency, accountName}
            m_dataset = CType(addToDatatable(m_dataset, "Account", m_acctColumns, colvalues), System.Data.DataSet)

        Else
            If drow.Count = 0 Then  ' row does not exist, hence add-row to datatable
                colvalues = {key, val_Renamed, curency, accountName}
                m_dataset = CType(addToDatatable(m_dataset, "Account", m_acctColumns, colvalues), System.Data.DataSet)

            ElseIf drow.Count = 1 Then  ' row exist, update the value
                drow(0).Item("Value") = val_Renamed
                ' show the table
                m_dataset = CType(addToDatatable(m_dataset, "Account", Nothing, Nothing), System.Data.DataSet)

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
        drow = findRowInDatatable(m_dataset, "Portfolio", _
                                       {"Symbol", "SecType", "PrimaryExch", "Currency", "Account", "TradingClass", "ConId"}, _
                                       {contract.Symbol, contract.SecType, contract.PrimaryExch, contract.Currency, accountName, contract.TradingClass, CStr(contract.ConId)}, _
                                       {"=", "=", "=", "=", "=", "=", "="})
        '
        If Not IsNothing(contract.Currency) Then
            If drow Is Nothing Then ' row does not exist, add row
                With contract
                    colvalues = {.Symbol, .SecType, .PrimaryExch, .Currency, .LocalSymbol, CStr(position), CStr(marketPrice), _
                                  CStr(marketValue), CStr(averageCost), CStr(unrealizedPNL), _
                                  CStr(realizedPNL), .TradingClass, CStr(.ConId), .Expiry, CStr(.Strike), .Right, .Multiplier, accountName}
                End With
                m_dataset = CType(addToDatatable(m_dataset, "Portfolio", m_portfColumns, colvalues), System.Data.DataSet)

            Else
                If drow.Count = 0 Then ' row does not exist, add row
                    With contract
                        colvalues = {.Symbol, .SecType, .PrimaryExch, .Currency, .LocalSymbol, CStr(position), CStr(marketPrice), _
                                     CStr(marketValue), CStr(averageCost), CStr(unrealizedPNL), _
                                      CStr(realizedPNL), .TradingClass, CStr(.ConId), .Expiry, CStr(.Strike), .Right, .Multiplier, accountName}
                    End With
                    m_dataset = CType(addToDatatable(m_dataset, "Portfolio", m_portfColumns, colvalues), System.Data.DataSet)

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
                    m_dataset = CType(addToDatatable(m_dataset, "Portfolio", Nothing, Nothing), System.Data.DataSet)

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
        m_dlgNewOrder.txtboxOrderID.Text = CStr(i_orderID)
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
        'drow = .m_dataset.Tables("OpenOrders").Rows.Find(order.OrderId)
        drow = findRowInDatatable(m_dataset, "OpenOrders", {"OrderID"}, {CStr(order.OrderId)}, {"="})

        If drow Is Nothing Then
            ' row does not exist, add a new openOrder
            colvalues = {contract.Symbol, _
                     contract.SecType, _
                     contract.Exchange, _
                     contract.PrimaryExch, _
                     contract.Currency, _
                     order.Action, _
                     CStr(order.TotalQuantity), _
                     order.OrderType, _
                     DblMaxStr(order.LmtPrice), _
                     orderState.Status, _
                     orderState.WarningText, _
                     CStr(order.ClientId), _
                     CStr(order.OrderId), _
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

            m_dataset = CType(addToDatatable(m_dataset, "OpenOrders", m_openColumns, colvalues), System.Data.DataSet)

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
            m_dataset = CType(addToDatatable(m_dataset, "OpenOrders", Nothing, Nothing), System.Data.DataSet)

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
        drow = findRowInDatatable(m_dataset, "OrderStatus", {"OrderID"}, {orderStatus.orderId}, {"="})

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
            addToDatatable(m_dataset, "OrderStatus", Nothing, Nothing)

        ElseIf drow Is Nothing Then
            ' row does not exist, create a new row
            ' find the correct row-data from OpenOrders
            drow_or = findRowInDatatable(m_dataset, "OpenOrders", {"OrderID"}, {orderStatus.orderId}, {"="})

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
                m_dataset = CType(addToDatatable(m_dataset, "OrderStatus", m_orderstatusColumns, colvalues), System.Data.DataSet)

            ElseIf drow_or Is Nothing Then
                Call m_utils.addListItem(Utils.List_Types.ERRORS, " order ID = " & CStr(orderStatus.orderId) & " not found in OpenOrders")
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
        'drow = .findRowInDatatable(.m_dataset, "Executions", {"ExecID"}, {execution.ExecId})
        drow = findRowInDatatable(m_dataset, "Executions", {"Symbol", "SecType", "Currency", "Side", "OrderID"}, _
                                         {contract.Symbol, contract.SecType, contract.Currency, execution.Side, CStr(execution.OrderId)}, _
                                          {"=", "=", "=", "=", "="})

        If drow Is Nothing Then    ' no row exists
            colvalues = {contract.Symbol, _
                         contract.SecType, _
                         contract.Exchange, _
                         contract.Currency, _
                         execution.Side, _
                         CStr(execution.Shares), _
                         CStr(execution.Price), _
                         CStr(execution.CumQty), _
                         CStr(execution.AvgPrice), _
                         CStr(execution.OrderId), _
                         CStr(execution.ClientId), _
                         execution.AcctNumber, _
                         CStr(contract.ConId), _
                         execution.ExecId, _
                         reqId, _
                         execution.OrderRef, _
                         contract.Expiry, _
                         CStr(contract.Strike), _
                         contract.Right, _
                         contract.Multiplier, _
                         contract.PrimaryExch, _
                         contract.LocalSymbol, _
                         contract.TradingClass, _
                         CStr(execution.PermId), _
                         execution.Time, _
                         execution.Exchange, _
                         CStr(execution.Liquidation), _
                         execution.EvRule, _
                         CStr(execution.EvMultiplier)}

            m_dataset = CType(addToDatatable(m_dataset, "Executions", m_execColumns, colvalues), System.Data.DataSet)

            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "ExecutionDetails received, reqId=" & reqId)

        ElseIf drow IsNot Nothing Then      ' row exists update the row
            ' update the execution statuses
            drow(0)("Shares") = execution.Shares
            drow(0)("Price") = execution.Price
            drow(0)("CumQty") = execution.CumQty
            drow(0)("AvgPrice") = execution.AvgPrice
            drow(0)("Time") = execution.Time

            ' display table
            addToDatatable(m_dataset, "Executions", Nothing, Nothing)

            Call m_utils.addListItem(Utils.List_Types.SERVER_RESPONSES, "ExecutionDetails received, reqId=" & reqId)

        Else
            Call m_utils.addListItem(Utils.List_Types.ERRORS, "Unknown case encountered in IBData:updateExecutions")

        End If
    End Sub


End Class
