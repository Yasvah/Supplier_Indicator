Public Class Precalculatedvalue
    Private _ppm As Double
    Private _qnc_count As Integer
    Private _customer_claim_count As Integer
    Private _lnc_count As Double
    Private _logistic_rate As Double
    Private _delays_up_to_days_rate As Double
    Private _order_horizon_percentage_0_to_2 As Integer
    Private _order_horizon_percentage_3_to_4 As Integer
    Private _order_horizon_percentage_5_to_6 As Integer
    Private _order_horizon_percentage_7_to_8 As Integer
    Private _order_horizon_percentage_9_to_10 As Integer
    Private _order_horizon_percentage_11_to_12 As Integer
    Private _order_horizon_percentage_greater_than_12 As Integer

    Public ReadOnly Property PPM As Double
        Get
            Return _ppm
        End Get
    End Property
    Public ReadOnly Property QNC_COUNT As Integer
        Get
            Return _qnc_count
        End Get
    End Property
    Public ReadOnly Property CUSTOMER_CLAIM_COUNT As Integer
        Get
            Return _customer_claim_count
        End Get
    End Property
    Public ReadOnly Property LNC_COUNT As Double
        Get
            Return _lnc_count
        End Get
    End Property
    Public ReadOnly Property LOGISTIC_RATE As Double
        Get
            Return _logistic_rate
        End Get
    End Property
    Public ReadOnly Property DELAYS_UP_TO_DAYS_RATE As Double
        Get
            Return _delays_up_to_days_rate
        End Get
    End Property

    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_0_TO_2 As Integer
        Get
            Return _order_horizon_percentage_0_to_2
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_3_TO_4 As Integer
        Get
            Return _order_horizon_percentage_3_to_4
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_5_TO_6 As Integer
        Get
            Return _order_horizon_percentage_5_to_6
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_7_TO_8 As Integer
        Get
            Return _order_horizon_percentage_7_to_8
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_9_TO_10 As Integer
        Get
            Return _order_horizon_percentage_9_to_10
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_11_TO_12 As Integer
        Get
            Return _order_horizon_percentage_11_to_12
        End Get
    End Property
    Public ReadOnly Property ORDER_HORIZON_PERCENTAGE_GREATER_THAN_12 As Integer
        Get
            Return _order_horizon_percentage_greater_than_12
        End Get
    End Property


    Sub New(PPM As Double, QNC_COUNT As Integer, CUSTOMER_CLAIM_COUNT As Integer, LNC_COUNT As Double, LOGISTIC_RATE As Double, DELAYS_UP_TO_DAYS_RATE As Double, _
            ORDER_HORIZON_PERCENTAGE_0_TO_2 As Integer, ORDER_HORIZON_PERCENTAGE_3_TO_4 As Integer, ORDER_HORIZON_PERCENTAGE_5_TO_6 As Integer, ORDER_HORIZON_PERCENTAGE_7_TO_8 As Integer, _
            ORDER_HORIZON_PERCENTAGE_9_TO_10 As Integer, ORDER_HORIZON_PERCENTAGE_11_TO_12 As Integer, ORDER_HORIZON_PERCENTAGE_GREATHER_THAN_12 As Integer)
        Me._ppm = PPM
        Me._qnc_count = QNC_COUNT
        Me._customer_claim_count = CUSTOMER_CLAIM_COUNT
        Me._lnc_count = LNC_COUNT
        Me._logistic_rate = LOGISTIC_RATE
        Me._delays_up_to_days_rate = DELAYS_UP_TO_DAYS_RATE
        Me._order_horizon_percentage_0_to_2 = ORDER_HORIZON_PERCENTAGE_0_TO_2
        Me._order_horizon_percentage_3_to_4 = ORDER_HORIZON_PERCENTAGE_3_TO_4
        Me._order_horizon_percentage_5_to_6 = ORDER_HORIZON_PERCENTAGE_5_TO_6
        Me._order_horizon_percentage_7_to_8 = ORDER_HORIZON_PERCENTAGE_7_TO_8
        Me._order_horizon_percentage_9_to_10 = ORDER_HORIZON_PERCENTAGE_9_TO_10
        Me._order_horizon_percentage_11_to_12 = ORDER_HORIZON_PERCENTAGE_11_TO_12
        Me._order_horizon_percentage_greater_than_12 = ORDER_HORIZON_PERCENTAGE_GREATHER_THAN_12
    End Sub
End Class