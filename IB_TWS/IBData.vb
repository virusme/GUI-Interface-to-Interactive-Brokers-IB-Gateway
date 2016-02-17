'=================================================================
'
' Author:  Whacky - The Portfolio Trader
' version: 1.0, 2014
'=================================================================

Option Explicit On
Option Strict On

Imports System.IO

Friend Class IBData
    Implements IDisposable

    '================================================================================
    ' Private Members
    '================================================================================
    ' data members
    Private m_utils As Utils
    Private m_dlgMain As dlgMain
    Private m_dlgNewOrder As New dlgNewOrder
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

        '
        AppPath = Directory.GetCurrentDirectory

        ' create datatables to hold data
        m_dataset = New DataSet
        m_dataset = CType(createDataTable(m_dataset, "Account", m_acctColumns), System.Data.DataSet)
        m_dataset = CType(createDataTable(m_dataset, "Portfolio", m_portfColumns), System.Data.DataSet)
        m_dataset = CType(createDataTable(m_dataset, "OpenOrders", m_openColumns), System.Data.DataSet)
        m_dataset = CType(createDataTable(m_dataset, "OrderStatus", m_orderstatusColumns), System.Data.DataSet)
        m_dataset = CType(createDataTable(m_dataset, "Executions", m_execColumns), System.Data.DataSet)
        ' read IB settings
        m_IBsettings = New DataSet
        Try
            If Not File.Exists(AppPath & IBSETFILE) Then
                ' load the default file from resource
                File.WriteAllText(AppPath & IBSETFILE, My.Resources.ibSet)
            End If
            m_IBsettings.ReadXml(AppPath & IBSETFILE)
        Catch Ex As Exception
            Call m_utils.addListItem(Utils.List_Types.ERRORS, Ex.Message)
        End Try
        '  

        m_dlgNewOrder.init(Me)
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
    ' Clear 
    '================================================================================

    '--------------------------------------------------------------------------------
    ' Clear executions
    '--------------------------------------------------------------------------------
    Public Sub clearExecutions()
        m_dataset.Tables("Executions").Clear()
    End Sub
    '--------------------------------------------------------------------------------
    ' Clear Open Orders
    '--------------------------------------------------------------------------------
    Public Sub clearOpenOrders()
        m_dataset.Tables("OpenOrders").Clear()
    End Sub
    '--------------------------------------------------------------------------------
    ' Clear executions
    '--------------------------------------------------------------------------------
    Public Sub clearOrderStatus()
        m_dataset.Tables("OrderStatus").Clear()
    End Sub

    '--------------------------------------------------------------------------------
    ' Clear accounts
    '--------------------------------------------------------------------------------
    Public Sub clearAccounts()
        m_dataset.Tables("Account").Clear()
        m_dataset.Tables("Portfolio").Clear()
    End Sub


    '================================================================================
    ' Read, Write or Get
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Load Contract Info from settings file
    '--------------------------------------------------------------------------------
    Public Sub loadContractInfo()
        Dim dTKrow As DataRow

        dTKrow = m_IBsettings.Tables("TickerAttributes").Rows(0)

        With m_dlgNewOrder
            ' Ticker attributes
            .txtboxTicker.Text = Nothing
            .txtboxType.Text = dTKrow("StkType").ToString
            .txtboxExchange.Text = dTKrow("Exchange").ToString
            .txtboxPrimExch.Text = dTKrow("PrimExchange").ToString
            .txtboxCurrency.Text = dTKrow("Currency").ToString
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

        dOrdrow = m_IBsettings.Tables("OrderAttributes").Rows(0)


        With m_dlgNewOrder
            ' Order Attributes
            .txtboxAction.Text = Nothing
            .txtboxQuantity.Text = Nothing
            .txtboxOrderType.Text = Nothing
            .txtboxPrice.Text = Nothing
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
        With m_dlgNewOrder
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
        With m_dlgNewOrder
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
            m_orderInfo.Hidden = CBool(.txtboxHidden.Text)
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

        dTKrow = m_IBsettings.Tables("TickerAttributes").Rows(0)
        m_IBsettings.Tables("TickerAttributes").Rows(0).Delete()

        With m_dlgNewOrder
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
        m_IBsettings.Tables("TickerAttributes").Rows.Add(dTKrow)
        'write to XML
        Try
            m_IBsettings.WriteXml(AppPath & IBSETFILE)
        Catch ex As Exception
            Call m_utils.addListItem(Utils.List_Types.ERRORS, ex.Message)
        End Try

    End Sub

    '--------------------------------------------------------------------------------
    ' Write Order Info to settings file
    '--------------------------------------------------------------------------------
    Private Sub writeOrderInfo()
        Dim dOrdrow As DataRow

        dOrdrow = m_IBsettings.Tables("OrderAttributes").Rows(0)
        m_IBsettings.Tables("OrderAttributes").Rows(0).Delete()

        With m_dlgNewOrder

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
        m_IBsettings.Tables("OrderAttributes").Rows.Add(dOrdrow)
        'write to XML
        Try
            m_IBsettings.WriteXml(AppPath & IBSETFILE)
        Catch ex As Exception
            Call m_utils.addListItem(Utils.List_Types.ERRORS, ex.Message)
        End Try

    End Sub



    '================================================================================
    ' Show DialogBoxes
    '================================================================================
    Public Sub showDlg(ByRef boxName As String, Optional ByVal b_fromPortfolio As Boolean = False, _
                                                Optional ByVal i_rowNum As Integer = -1)
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

            Case "NewOrder"
                ' check if data should be populated from portfolio
                If b_fromPortfolio And i_rowNum > -1 Then
                    ' load order attributes from portfolio using row-number from the right-click operation
                    Dim drow As DataRow
                    drow = m_dataset.Tables("Portfolio").Rows(i_rowNum)
                    ' populate the dialog
                    With m_dlgNewOrder
                        ' Ticker attributes
                        .txtboxTicker.Text = drow("Symbol").ToString
                        .txtboxType.Text = drow("SecType").ToString
                        '.txtboxExchange.Text = drow("Exchange").ToString
                        .txtboxPrimExch.Text = drow("PrimaryExch").ToString
                        .txtboxCurrency.Text = drow("Currency").ToString
                        ' Order Attributes
                        If dval(drow("Position").ToString) < 0 Then
                            .txtboxAction.Text = "BUY"
                        ElseIf dval(drow("Position").ToString) > 0 Then
                            .txtboxAction.Text = "SELL"
                        Else
                            ' do nothing
                        End If
                        .txtboxQuantity.Text = drow("Position").ToString
                        .txtboxOrderType.Text = Nothing
                        .txtboxPrice.Text = Nothing
                    End With

                Else
                    loadContractInfo()
                    loadOrderInfo()
                End If

                ' show dialog
                m_dlgNewOrder.ShowDialog()

            Case Else
                Call m_utils.addListItem(Utils.List_Types.ERRORS, "Unknown case in IBData:showDlg")
        End Select
    End Sub





#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                m_dlgExtOrderAttr.Dispose()
                m_dlgExtTickAtrr.Dispose()
                m_dlgNewOrder.Dispose()
                m_dataset.Dispose()
                m_IBsettings.Dispose()
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
