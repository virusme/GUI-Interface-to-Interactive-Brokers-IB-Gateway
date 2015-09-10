'=================================================================
'
' Author:  Whacky - The Portfolio Trader
' version: 1.01, 2015
'=================================================================

Option Explicit On
Option Strict Off


Friend Class dlgNewOrder
    Public Sub New()
        MyBase.New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    '================================================================================
    ' Private Members
    '================================================================================
    Private m_IBdata As IBData

    '================================================================================
    ' Private Methods
    '================================================================================
    Private Sub dlgNewOrder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    '--------------------------------------------------------------------------------
    ' Initialisation
    '--------------------------------------------------------------------------------
    Public Sub init(ByRef myIBdata As IBData)
        m_IBdata = myIBdata

    End Sub

    '================================================================================
    ' Button Events
    '================================================================================
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
        Me.Hide()
    End Sub
    Private Sub btnPlaceModifyOrder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPlaceModifyOrder.Click
        m_IBdata.placeOrderImpl(False)
        Me.Hide()
    End Sub


End Class