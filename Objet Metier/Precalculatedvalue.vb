Public Class Precalculatedvalue
    Private _ppm As Double
    Private _qnc_count As Integer
    Private _customer_claim_count As Integer
    Private _lnc_count As Double
    Private _logistic_rate As Double
    Private _delays_up_to_days_rate As Double

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

    Sub New(PPM As Double, QNC_COUNT As Integer, CUSTOMER_CLAIM_COUNT As Integer, LNC_COUNT As Double, LOGISTIC_RATE As Double, DELAYS_UP_TO_DAYS_RATE As Double)
        Me._ppm = PPM
        Me._qnc_count = QNC_COUNT
        Me._customer_claim_count = CUSTOMER_CLAIM_COUNT
        Me._lnc_count = LNC_COUNT
        Me._logistic_rate = LOGISTIC_RATE
        Me._delays_up_to_days_rate = DELAYS_UP_TO_DAYS_RATE
    End Sub
End Class