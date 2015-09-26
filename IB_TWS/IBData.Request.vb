Option Strict On

'-------------------------------------------------------------------------------------
'
' Request - Container for all functions and procedures related to IBData.Request
'
'-------------------------------------------------------------------------------------

Partial Friend Class IBData

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

End Class
