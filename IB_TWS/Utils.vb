' Copyright (C) 2013 Interactive Brokers LLC. All rights reserved. This code is subject to the terms
' and conditions of the IB API Non-Commercial License or the IB API Commercial License, as applicable.

Option Strict Off
Option Explicit On

Imports System.Data.OleDb
Imports System.Xml
Imports System.IO

Friend Class Utils
    Implements IDisposable
    ' Enums
    Public Enum TickType
        BID_SIZE = 0
        BID_PRICE
        ASK_PRICE
        ASK_SIZE
        LAST_PRICE
        LAST_SIZE
        HIGH
        LOW
        VOLUME
        CLOSE_PRICE
        BID_OPTION_COMPUTATION
        ASK_OPTION_COMPUTATION
        LAST_OPTION_COMPUTATION
        MODEL_OPTION
        OPEN_TICK
        LOW_13_WEEK
        HIGH_13_WEEK
        LOW_26_WEEK
        HIGH_26_WEEK
        LOW_52_WEEK
        HIGH_52_WEEK
        AVG_VOLUME
        OPEN_INTEREST
        OPTION_HISTORICAL_VOL
        OPTION_IMPLIED_VOL
        OPTION_BID_EXCH
        OPTION_ASK_EXCH
        OPTION_CALL_OPEN_INTEREST
        OPTION_PUT_OPEN_INTEREST
        OPTION_CALL_VOLUME
        OPTION_PUT_VOLUME
        INDEX_FUTURE_PREMIUM
        BID_EXCH
        ASK_EXCH
        AUCTION_VOLUME
        AUCTION_PRICE
        AUCTION_IMBALANCE
        MARK_PRICE
        BID_EFP_COMPUTATION
        ASK_EFP_COMPUTATION
        LAST_EFP_COMPUTATION
        OPEN_EFP_COMPUTATION
        HIGH_EFP_COMPUTATION
        LOW_EFP_COMPUTATION
        CLOSE_EFP_COMPUTATION
        LAST_TIMESTAMP
        SHORTABLE
        FUNDAMENTAL_RATIOS
        RT_VOLUME
        HALTED
        BID_YIELD
        ASK_YIELD
        LAST_YIELD
        CUST_OPTION_COMPUTATION
        TRADE_COUNT
        TRADE_RATE
        VOLUME_RATE
        LAST_RTH_TRADE
    End Enum

    Public Enum List_Types
        MKT_DATA = 0
        SERVER_RESPONSES
        ERRORS
        ACCOUNT_DATA
        PORTFOLIO_DATA
        DISPLAY_GROUPS_DATA
    End Enum

    Public Enum FA_Message_Type
        GROUPS = 1
        PROFILES = 2
        ALIASES = 3
    End Enum

    Public Shared ReadOnly CRSTR As String = Chr(13)
    Public Shared ReadOnly LFSTR As String = Chr(10)
    Public Shared ReadOnly CRLFSTR As String = CRSTR & LFSTR

    ' Win32 API functions
    Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    ' Constants
    Private Const LB_SETHORZEXTENT As Short = &H194S
    Public Const ERR_NOCONNECTION = 1100
    Public Const ERR_IDENT = " Error Code"


    Public m_dlgMain As dlgMain
    'Private m_dlgAcctData As dlgAcctData
    'Private m_dlgGroups As dlgGroups
    ' Datasets
    Public m_dataset As DataSet
    Public m_IBsettings As DataSet
    Public AppPath As String
    Public onTime_Hour As Integer
    Public onTime_Min As Integer
    Public TriggerTime As DateTime
    Public IBSETFILE As String = "/utilities/ibSet.xml"
    Public m_connColumns As String() = {"IPAddress", "Port", "ClientID", "ServerLogLevel"}
    Public m_acctColumns As String() = {"Key", "Value", "Currency", "Account"}
    Public m_portfColumns As String() = {"Symbol", "SecType", "PrimaryExch", "Currency", "LocalSymbol", _
                                          "Position", "MktPrice", "MktValue", "AvgCost", "UnrealizedPNL", _
                                         "RealizedPNL", "TradingClass", "ConId", "Expiry", "Strike", "Right", "Multiplier", "Account"}
    Public m_openColumns As String() = {"Symbol", "SecType", "Exchange", "PrimExchange", "Currency", "Action", "Quantity", "OrderType", _
                                        "Price", "Status", "WarningText", "ClientID", "OrderID", "LocalSymbol"}
    ''----
    '' currently not using below fields
    ''----
    '                                    {"PermID", "ConID", "TimeInForce", "GoodTillDate", "GoodAfterTime", _
    '                                    "OutsideRTH", "AuxPrice", "OCAGroup", "OCAType", "OrderRef", "ParentID", "BlockOrder", "SweepToFill", "Display", _
    '                                    "TriggerMethod", "Hidden", "OverridePercent", "Rule80A", "AllNone", "MinQuantity", "PercentOffset", _
    '                                    "TrailStopPrice", "TrailingPercent", "WhatIf", "NotHeld", "InitMargin", "MaintMargin", "EquityWithLoan", _
    '                                    "Commission", "MinCommission", "MaxCommission", "CommissionCurr", "Expiry", "Strike", "Right", "Multiplier", "TradingClass", _
    '                                    "ComboLegsDescr"}

    Public m_orderstatusColumns As String() = {"Symbol", "SecType", "Exchange", "Currency", "Action", "Quantity", "OrderType", _
                                               "Price", "Status", "Filled", "Remaining", "AvgFillPrice", "LastFillPrice", "PermID", _
                                               "ParentID", "WhyHeld", "OrderID", "ClientID"}

    Public m_execColumns As String() = {"Symbol", "SecType", "Exchange", "Currency", "Side", "Shares", "Price", _
                                        "CumQty", "AvgPrice", "OrderID", "ClientID", "Account", "ConID", "ExecID", "ReqID", "OrderRef", _
                                        "Expiry", "Strike", "Right", "Multiplier", "PrimExchange", "LocalSymbol", "TradingClass", "PermID", _
                                        "Time", "ExecExchange", "Liquidation", "evRule", "evMultiplier"}

    '================================================================================
    ' Public Helper Routines
    '================================================================================
    Public Sub init(ByVal dlgMainWnd As System.Windows.Forms.Form)
        m_dlgMain = dlgMainWnd
        '
        AppPath = Directory.GetCurrentDirectory

        ' create datatables to hold data
        m_dataset = New DataSet
        m_dataset = createDataTable(m_dataset, "Account", m_acctColumns)
        m_dataset = createDataTable(m_dataset, "Portfolio", m_portfColumns)
        m_dataset = createDataTable(m_dataset, "OpenOrders", m_openColumns)
        m_dataset = createDataTable(m_dataset, "OrderStatus", m_orderstatusColumns)
        m_dataset = createDataTable(m_dataset, "Executions", m_execColumns)

        ' read IB settings
        m_IBsettings = New DataSet
        Try
            m_IBsettings.ReadXml(AppPath & IBSETFILE)
        Catch Ex As Exception
            Call addListItem(Utils.List_Types.ERRORS, Ex.Message)
        End Try
        With m_dlgMain
            .txtboxIPAddress.Text = m_IBsettings.Tables("Connection").Rows(0)("IPAddress").ToString
            .txtboxPort.Text = m_IBsettings.Tables("Connection").Rows(0)("Port").ToString()
            .txtboxClientID.Text = m_IBsettings.Tables("Connection").Rows(0)("ClientID").ToString()
            .cmbboxLogLevel.SelectedIndex = m_IBsettings.Tables("Connection").Rows(0)("ServerLogLevel").ToString
        End With



    End Sub

    '--------------------------------------------------------------------------------
    ' Display an XML message
    '--------------------------------------------------------------------------------
    Public Sub displayMultiline(ByRef listType As List_Types, ByRef xml As String)
        Dim ctr As Object
        Dim FOUR_SPACES As Object
        Dim CRLFSTR As Object
        Dim CRSTR As Object
        Dim LFSTR As Object
        Dim TABSTR As Object
        TABSTR = Chr(9)
        LFSTR = Chr(10)
        CRSTR = Chr(13)
        CRLFSTR = CRSTR & LFSTR
        FOUR_SPACES = "    "
        Dim text As String
        Dim strArray() As String
        text = Replace(xml, TABSTR, FOUR_SPACES)
        strArray = Split(text, LFSTR) ' VB.NET must be removing the CR
        For ctr = 0 To UBound(strArray)
            Dim line As String
            line = strArray(ctr)
            If Len(line) > 1020 Then
                line = Left(line, 1020) & " ..."
            End If
            Call addListItem(listType, line)
        Next
    End Sub
    Public Function faMsgTypeName(ByVal faMsgType As FA_Message_Type) As Object
        Select Case faMsgType
            Case FA_Message_Type.GROUPS
                faMsgTypeName = "GROUPS"
            Case FA_Message_Type.PROFILES
                faMsgTypeName = "PROFILES"
            Case FA_Message_Type.ALIASES
                faMsgTypeName = "ALIASES"
            Case Else
                faMsgTypeName = "UNKNOWN"
        End Select
    End Function
    '--------------------------------------------------------------------------------
    ' Add an item to one of the display listboxes
    '--------------------------------------------------------------------------------
    Public Sub addListItem(ByRef listType As List_Types, ByRef data As String, Optional ByVal scrollList As Boolean = True, _
                                                                               Optional ByVal doclear As Boolean = False)
        Dim listBox As System.Windows.Forms.ListBox
        listBox = Nothing
        Select Case listType
            Case List_Types.MKT_DATA
                'listBox = m_dlgMain.lstMktData
            Case List_Types.SERVER_RESPONSES
                listBox = m_dlgMain.lstServerResponses
            Case List_Types.ERRORS
                listBox = m_dlgMain.lstErrors
            Case List_Types.ACCOUNT_DATA
                'listBox = m_dlgMain.lstKeyValueData
            Case List_Types.PORTFOLIO_DATA
                'listBox = m_dlgMain.lstPortfolioData
            Case List_Types.DISPLAY_GROUPS_DATA
                'listBox = m_dlgGroups.lstGroupMessages
            Case Else
                Return
        End Select
        If Not IsNothing(data) Then
            Dim lines() As String = data.Split(LFSTR.ToCharArray, 500)
            For ctr As Integer = 0 To lines.GetLength(0) - 1
                Dim theLine As String = lines(ctr).Replace(CRSTR, "")
                listBox.Items.Add(theLine)
            Next
        End If
        If doclear Then
            listBox.Items.Clear()
        End If
        If scrollList Then
            listBox.TopIndex = listBox.Items.Count - 1
        End If
    End Sub

    '--------------------------------------------------------------------------------
    ' Returns the tick display string used when adding tick data to the display list.
    '--------------------------------------------------------------------------------
    Public Function getField(ByRef field As TickType) As Object
        Select Case field
            Case TickType.BID_SIZE : getField = "bidSize"
            Case TickType.BID_PRICE : getField = "bidPrice"
            Case TickType.ASK_PRICE : getField = "askPrice"
            Case TickType.ASK_SIZE : getField = "askSize"
            Case TickType.LAST_PRICE : getField = "lastPrice"
            Case TickType.LAST_SIZE : getField = "lastSize"
            Case TickType.HIGH : getField = "high"
            Case TickType.LOW : getField = "low"
            Case TickType.VOLUME : getField = "volume"
            Case TickType.CLOSE_PRICE : getField = "close"
            Case TickType.BID_OPTION_COMPUTATION : getField = "bidOptComp"
            Case TickType.ASK_OPTION_COMPUTATION : getField = "askOptComp"
            Case TickType.LAST_OPTION_COMPUTATION : getField = "lastOptComp"
            Case TickType.MODEL_OPTION : getField = "modelOption"
            Case TickType.OPEN_TICK : getField = "open"
            Case TickType.LOW_13_WEEK : getField = "13WeekLow"
            Case TickType.HIGH_13_WEEK : getField = "13WeekHigh"
            Case TickType.LOW_26_WEEK : getField = "26WeekLow"
            Case TickType.HIGH_26_WEEK : getField = "26WeekHigh"
            Case TickType.LOW_52_WEEK : getField = "52WeekLow"
            Case TickType.HIGH_52_WEEK : getField = "52WeekHigh"
            Case TickType.AVG_VOLUME : getField = "AvgVolume"
            Case TickType.OPEN_INTEREST : getField = "OpenInterest"
            Case TickType.OPTION_HISTORICAL_VOL : getField = "OptionHistoricalVolatility"
            Case TickType.OPTION_IMPLIED_VOL : getField = "OptionImpliedVolatility"
            Case TickType.OPTION_BID_EXCH : getField = "OptionBidExchStr"
            Case TickType.OPTION_ASK_EXCH : getField = "OptionAskExchStr"
            Case TickType.OPTION_CALL_OPEN_INTEREST : getField = "OptionCallOpenInterest"
            Case TickType.OPTION_PUT_OPEN_INTEREST : getField = "OptionPutOpenInterest"
            Case TickType.OPTION_CALL_VOLUME : getField = "OptionCallVolume"
            Case TickType.OPTION_PUT_VOLUME : getField = "OptionPutVolume"
            Case TickType.INDEX_FUTURE_PREMIUM : getField = "IndexFuturePremium"
            Case TickType.BID_EXCH : getField = "bidExch"
            Case TickType.ASK_EXCH : getField = "askExch"
            Case TickType.AUCTION_VOLUME : getField = "auctionVolume"
            Case TickType.AUCTION_PRICE : getField = "auctionPrice"
            Case TickType.AUCTION_IMBALANCE : getField = "auctionImbalance"
            Case TickType.MARK_PRICE : getField = "markPrice"
            Case TickType.BID_EFP_COMPUTATION : getField = "bidEFP"
            Case TickType.ASK_EFP_COMPUTATION : getField = "askEFP"
            Case TickType.LAST_EFP_COMPUTATION : getField = "lastEFP"
            Case TickType.OPEN_EFP_COMPUTATION : getField = "openEFP"
            Case TickType.HIGH_EFP_COMPUTATION : getField = "highEFP"
            Case TickType.LOW_EFP_COMPUTATION : getField = "lowEFP"
            Case TickType.CLOSE_EFP_COMPUTATION : getField = "closeEFP"
            Case TickType.LAST_TIMESTAMP : getField = "lastTimestamp"
            Case TickType.SHORTABLE : getField = "shortable"
            Case TickType.FUNDAMENTAL_RATIOS : getField = "fundamentals"
            Case TickType.RT_VOLUME : getField = "RTVolume"
            Case TickType.HALTED : getField = "halted"
            Case TickType.BID_YIELD : getField = "bidYield"
            Case TickType.ASK_YIELD : getField = "askYield"
            Case TickType.LAST_YIELD : getField = "lastYield"
            Case TickType.CUST_OPTION_COMPUTATION : getField = "custOptComp"
            Case TickType.TRADE_COUNT : getField = "trades"
            Case TickType.TRADE_RATE : getField = "trades/min"
            Case TickType.VOLUME_RATE : getField = "volume/min"
            Case TickType.LAST_RTH_TRADE : getField = "lastRTHTrade"
            Case Else : getField = "unknown"
        End Select
    End Function

    '================================================================================
    ' Datasets and Datatables
    '================================================================================
    '--------------------------------------------------------------------------------
    ' Create a datatable
    '   - For creating a datatable within a dataset
    '       dataset = createDataTable(dataset, tablename, colnames)
    '   - For creating a datatable 
    '       datatable = createDataTable(Nothing, tablename, colnames)
    '        
    '--------------------------------------------------------------------------------
    Public Function createDataTable(ByRef dset As DataSet, ByRef tablename As String, ByRef colnames As String()) As Object

        Dim dtable As DataTable
        Dim colm As DataColumn
        Dim i As Integer

        dtable = New DataTable
        dtable.TableName = tablename
        For i = 0 To colnames.Length - 1
            colm = New DataColumn(colnames(i), Type.GetType("System.String"))
            dtable.Columns.Add(colm)
            colm = Nothing
        Next

        If dset IsNot Nothing Then
            dset.Tables.Add(dtable)
            Return dset
        Else
            Return dtable
        End If
    End Function

    '--------------------------------------------------------------------------------
    ' Add a new row to datatable and display table in respective DataGrid
    '   - For adding a new row to datatable within a dataset
    '        dataset = addToDatatable(dataset, tablename, colnames, colvalues)
    '   - For adding a new row to datatable
    '        datatable = addToDatatable(Nothing, tablename, colnames, datatable)
    '
    '--------------------------------------------------------------------------------
    Public Function addToDatatable(ByRef dset As DataSet, ByRef tablename As String, ByRef colnames As String(), _
                                   ByRef colvalues As String(), Optional ByRef dtable As DataTable = Nothing) As Object
        Dim drow As DataRow
        Dim id As Integer
        Dim i As Integer
        Dim type As Boolean

        ' check input
        If dset IsNot Nothing And dtable Is Nothing Then
            type = True ' dataset input
        Else
            type = False ' datatable input
        End If
        ' get the correct table 
        If type Then
            id = dset.Tables.IndexOf(tablename)
            dtable = dset.Tables(id)
        End If
        '
        If colnames IsNot Nothing Then
            drow = dtable.NewRow()
            '
            For i = 0 To (colnames.Length - 1)
                drow(colnames(i)) = colvalues(i)
            Next
            '
            dtable.Rows.Add(drow)
            '
            If type Then
                dset.Tables(id).Merge(dtable)
            End If
        End If
        ' show table
        Select Case tablename
            Case "Account"
                m_dlgMain.gridvwAcctSummary.DataSource = dtable
            Case "Portfolio"
                m_dlgMain.gridvwPortfolio.DataSource = dtable
            Case "Connection"
                ' do nothing
            Case "OpenOrders"
                m_dlgMain.gridvwOpenOrders.DataSource = dtable
            Case "OrderStatus"
                m_dlgMain.gridvwOrderStatus.DataSource = dtable
            Case "Executions"
                m_dlgMain.gridvwExecutionsReport.DataSource = dtable
            Case "freshOrders"
                'do nothing
            Case "Sym"
                'do nothing
            Case "settings"
                'do nothing
            Case Else
                Call addListItem(Utils.List_Types.ERRORS, "Unknown CASE: " & tablename & " inside Utils::addToDataTable")
        End Select
        '
        If type Then
            Return dset
        Else
            Return dtable
        End If

    End Function

    '--------------------------------------------------------------------------------
    ' Find row-index in a datatable 
    ' searches for UNIQUE values otherwise returns NOTHING if does not exist, throws error and returns DROW if more than one row)
    '
    '   - For finding row in datatable within dataset
    '       drow = findRowInDatatavle(dataset, tablename, colname, rowvalue)
    '   - For finding row in datatable
    '       drow = findRowInDatatable(nothing, tablename, colname, rowvalue, datatable)
    '
    '--------------------------------------------------------------------------------
    Public Function findRowInDatatable(ByRef dset As DataSet, ByRef tableName As String, _
                                       ByRef columnName As String(), ByRef rowValue() As String, Optional ByRef dtable As DataTable = Nothing) As DataRow()
        Dim drow() As DataRow
        Dim exp As String
        Dim id As Integer
        Dim i As Integer

        drow = Nothing
        exp = Nothing
        If dset IsNot Nothing And dtable Is Nothing Then
            ' get the correct table
            id = dset.Tables.IndexOf(tableName)
            dtable = dset.Tables(id)
        Else
            ' do nothing
        End If
        ' build an expression
        If columnName.Count <> rowValue.Count Then
            MsgBox("Error: Utils:findRowInDataTable: columnName and rowValue do not hae same number of elements")
        End If
        ' load the first search criteria
        exp = columnName(0) & " = '" & rowValue(0) & "'"
        ' if more criterias exists load them one-by-one
        If columnName.Count > 1 Then
            For i = 1 To columnName.Count - 1
                exp = exp & " AND " & columnName(i) & " = '" & rowValue(i) & "'"
            Next
        End If
        ' get the row
        drow = dtable.Select(exp)
        If drow IsNot Nothing Then
            If drow.Count = 1 Then

            ElseIf drow.Count > 1 Then
                Call addListItem(Utils.List_Types.ERRORS, "search = " & exp & " FOUND IN MORE THAN one-row in table = " & tableName)

            ElseIf drow.Count = 0 Then
                drow = Nothing
            End If
        End If
        Return drow

    End Function

    '--------------------------------------------------------------------------------
    ' Update Daily Run Time 
    '--------------------------------------------------------------------------------
    Public Sub updateDailyTime(ByVal timeStr As String)
        Dim dt As DateTime
        If String.IsNullOrEmpty(timeStr) Then
            onTime_Hour = Nothing
            onTime_Min = Nothing
            'm_myutils.runDaily = False
        Else
            dt = Convert.ToDateTime(timeStr)
            onTime_Hour = dt.Hour
            onTime_Min = dt.Minute
            'm_myUtils.runDaily = True
        End If
    End Sub

    '--------------------------------------------------------------------------------
    ' Check if IB Gateway is connected to IB server
    '--------------------------------------------------------------------------------
    Public Function checkConnection() As Boolean
        Dim str As String
        Dim parts() As String
        Dim subparts() As String
        Dim topindex As Integer
        Dim chk As Boolean

        chk = True
        topindex = m_dlgMain.lstErrors.Items.Count - 1
        If topindex > -1 Then
            str = m_dlgMain.lstErrors.Items(topindex)
            ' debugging string
            'str = "Time: 9/3/2015 @ 15:42:35.3653708 | id: -1 | Error Code: 1100 | Error Msg: Connectivity between IB and Trader Workstation has been lost."
            parts = str.Split("|")
            If parts IsNot Nothing Then
                If parts.Length > 3 Then
                    subparts = parts(2).Split(":")
                    If subparts(0) IsNot Nothing Then
                        If String.Compare(subparts(0), ERR_IDENT) = 0 Then
                            If subparts(1) IsNot Nothing Then
                                If CInt(subparts(1)) = ERR_NOCONNECTION Then
                                    ' IB Gateway is disconnected from IB Server
                                    chk = False
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return chk
    End Function

    
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
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
