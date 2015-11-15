Partial Friend Class IBData

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

    Public Const ERR_NOCONNECTION = 1100
    Public Const ERR_ID = "id"
    Public Const ERR_ID_MINUSONE = -1
    Public Const ERR_MSG = "Error Msg"
    Public Const ERR_IDENT = "Error Code"
    Public Const CUSHION = "Cushion"

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
        ' show datatable
        m_dlgMain.showDataTable(tablename, dtable)
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
                                       ByRef columnName As String(), ByRef rowValue() As String, Optional ByRef operatorValue() As String = Nothing, _
                                       Optional ByRef dtable As DataTable = Nothing) As DataRow()
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
        If IsNothing(operatorValue) Then
            exp = columnName(0) & " = '" & rowValue(0) & "'"
        Else
            exp = columnName(0) & " " & operatorValue(0) & " '" & rowValue(0) & "'"
        End If
        ' if more criterias exists load them one-by-one
        If columnName.Count > 1 Then
            For i = 1 To columnName.Count - 1
                If IsNothing(operatorValue) Then
                    exp = exp & " AND " & columnName(i) & " = '" & rowValue(i) & "'"
                Else
                    exp = exp & " AND " & columnName(i) & " " & operatorValue(i) & " '" & rowValue(i) & "'"
                End If
            Next
        End If
        ' get the row
        drow = dtable.Select(exp)
        If drow IsNot Nothing Then
            If drow.Count = 1 Then

            ElseIf drow.Count > 1 Then
                Call m_utils.addListItem(Utils.List_Types.ERRORS, "search = " & exp & " FOUND IN MORE THAN one-row in table = " & tableName)

            ElseIf drow.Count = 0 Then
                drow = Nothing
            End If
        End If
        Return drow

    End Function

    '--------------------------------------------------------------------------------
    ' Sort Datatable 
    '--------------------------------------------------------------------------------
    Public Function sortDatatable(ByVal dtable As DataTable, ByVal columnName As String, _
                                  Optional ByVal type As String = "ASC") As DataTable
        Dim dview As New DataView(dtable)
        dview.Sort = columnName & " " & type
        Return dview.ToTable()
    End Function

    '--------------------------------------------------------------------------------
    ' Sort Datatable using "Date"
    '--------------------------------------------------------------------------------
    Public Function sortUsingDate(ByVal dtable As DataTable, ByVal columnName As String) As DataTable
        ' check the datatable already has a column name "sortDt"
        If dtable.Columns.Contains("sortDt") Then
            'do nothing
        Else
            ' add a new column to datatable
            dtable.Columns.Add(New DataColumn("sortDt", System.Type.GetType("System.DateTime")))
        End If
        ' convert each row-date into DateTime
        For Each row As DataRow In dtable.Rows
            row.Item("sortDt") = obj2date(row.Item(columnName))
        Next
        ' sort and return
        If dtable IsNot Nothing AndAlso dtable.Rows.Count > 0 Then
            dtable = sortDatatable(dtable, "sortDt")

        End If
        ' check the datatable has column name "sortDt" in order to remove
        If dtable.Columns.Contains("sortDt") Then
            dtable.Columns.Remove("sortDt")
        Else
            'do nothing
        End If
        Return dtable
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
                    ' check for "id" value
                    subparts = parts(2).Split(":")
                    If subparts(0) IsNot Nothing Then
                        If String.Compare(subparts(0).Trim, ERR_IDENT) = 0 Then
                            If subparts(1) IsNot Nothing Then
                                If CInt(subparts(1).Trim) = ERR_NOCONNECTION Then
                                    ' IB Gateway is disconnected from IB server or something is amiss
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

    '--------------------------------------------------------------------------------
    ' Get Error Msg from the server
    '--------------------------------------------------------------------------------
    Public Function getErrorMsg() As String
        Dim str As String
        Dim parts() As String
        Dim subparts() As String
        Dim topindex As Integer
        Dim errmsg As String

        errmsg = Nothing
        topindex = m_dlgMain.lstErrors.Items.Count - 1
        If topindex > -1 Then
            str = m_dlgMain.lstErrors.Items(topindex)
            ' debugging string
            'str = "Time: 9/3/2015 @ 15:42:35.3653708 | id: -1 | Error Code: 1100 | Error Msg: Connectivity between IB and Trader Workstation has been lost."
            parts = str.Split("|")
            If parts IsNot Nothing Then
                If parts.Length > 3 Then
                    ' extract "Error Msg"
                    subparts = parts(3).Split(":")
                    If subparts(0) IsNot Nothing Then
                        If String.Compare(subparts(0).Trim, ERR_MSG) = 0 Then
                            errmsg = subparts(1).ToString
                        End If
                    End If
                End If
            End If
        End If
        Return errmsg
    End Function

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
    Public Function obj2str(ByVal obj As Object) As String
        If IsDBNull(obj) Then
            obj2str = Nothing
        Else
            obj2str = CStr(obj)
        End If
    End Function
    Public Function obj2int(ByVal obj As Object) As Integer
        If IsDBNull(obj) Then
            obj2int = 0
        Else
            obj2int = CInt(obj)
        End If
    End Function
    Public Function obj2dbl(ByVal obj As Object) As Double
        If IsDBNull(obj) Then
            obj2dbl = 0
        Else
            obj2dbl = CDbl(obj)
        End If
    End Function
    Public Function obj2date(ByVal obj As Object) As Date
        If IsDBNull(obj) Then
            obj2date = Date.Now.AddDays(1)  ' Add a future-date
        Else
            obj2date = CDate(obj)
        End If
    End Function


End Class
