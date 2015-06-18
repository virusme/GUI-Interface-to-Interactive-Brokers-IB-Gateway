'=================================================================
'
' Author:  Whacky - The Portfolio Trader
' version: 1.0, 2014
'=================================================================

Option Explicit On
Option Strict Off

Friend Class dlgExtTickerAttr
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
    Public b_save As Boolean

    '================================================================================
    ' Private Methods
    '================================================================================
    Private Sub dlgExtTickerAttr_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
    ' Apply
    '--------------------------------------------------------------------------------
    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        b_save = True
        Me.Hide()
    End Sub
    '--------------------------------------------------------------------------------
    ' Cancel
    '--------------------------------------------------------------------------------
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        b_save = False
        Me.Hide()
    End Sub
End Class