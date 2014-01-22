Public Class AssessmentSERTA
    Inherits Assessment

    Public Sub New(idSupplier As Integer, quarter As Integer)
        MyBase.New(idSupplier, quarter)
    End Sub
    Public Sub New(id As Integer, supplier As Integer, quarter As String, form As Integer, indicePPMValue As Integer, _
                   indicePPMPoint As Integer, sinNBValue As Integer, sinNBPoint As Integer, customerClaimNBValue As Integer, customerClaimNBPoint As Integer, _
                   actionPlanReactivityPoint As Integer, bonus500PPMPoint As Integer, logisticRateTarget95Value As Integer, logisticRateTarget95Point As Integer, _
                   flexibilityPoint As Integer, deliveryDelaysLevelValue As Integer, deliveryDelaysLevelPoint As Integer, deliveryQualityValue As Integer, deliveryQualityPoint As Integer, _
                   priceCompetitivenessValue As Integer, priceCompetitivenessPoint As Integer, improvmentPlanPoint As Integer, businessRelationshipPoint As Integer, _
                   financialSituationPoint As Integer, OffersReactivityPoint As Integer, technicalAnswerQualityPoint As Integer, isoCertificationPoint As Integer, comment As String, _
                   commentDetail As String, commentClassification As String, commentGlobal As String, totalPoint As Integer, trend As Integer)
        MyBase.New(id, supplier, quarter, indicePPMValue, indicePPMPoint, sinNBValue, sinNBPoint, customerClaimNBValue, customerClaimNBPoint, _
                   actionPlanReactivityPoint, bonus500PPMPoint, logisticRateTarget95Value, logisticRateTarget95Point, _
                   flexibilityPoint, deliveryDelaysLevelValue, deliveryDelaysLevelPoint, deliveryQualityValue, deliveryQualityPoint, _
                   priceCompetitivenessValue, priceCompetitivenessPoint, improvmentPlanPoint, businessRelationshipPoint, _
                   financialSituationPoint, OffersReactivityPoint, technicalAnswerQualityPoint, isoCertificationPoint, comment, _
                   commentDetail, commentClassification, commentGlobal, totalPoint, trend)

    End Sub
    Private _PrecalculedValue As Precalculatedvalue
    Public ReadOnly Property PrecalculedValue As Precalculatedvalue
        Get
            Try
                If IsNothing(_PrecalculedValue) Then
                    Dim db As dbSupplierIndicatorDataContext = New dbSupplierIndicatorDataContext
                    For Each uneLigne In db.P_Assessment_Values_SERTA(Me.idSupplier, Me.quarter)
                        Dim PPM = If(IsNothing(uneLigne.PPM), 0, uneLigne.PPM)
                        Dim QNC_COUNT = uneLigne.QNC_COUNT
                        Dim CUSTOMER_CLAIN8COUNT = uneLigne.CUSTOMER_CLAIM_COUNT
                        Dim LNC_COUNT = uneLigne.LNC_COUNT
                        Dim LOGISTIC_RATE = IIf(IsNothing(uneLigne.LOGISTIC_RATE), 0, uneLigne.LOGISTIC_RATE)
                        Dim DELAY_UP_TO_DAYS_RATE = IIf(IsNothing(uneLigne.DELAYS_UP_TO_10_DAYS_RATE), 0, uneLigne.DELAYS_UP_TO_10_DAYS_RATE)
                        Dim _order_horizon_percentage_0_to_2 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_0_TO_2), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_0_TO_2))
                        Dim _order_horizon_percentage_3_to_4 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_3_TO_4), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_3_TO_4))
                        Dim _order_horizon_percentage_5_to_6 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_5_TO_6), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_5_TO_6))
                        Dim _order_horizon_percentage_7_to_8 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_7_TO_8), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_7_TO_8))
                        Dim _order_horizon_percentage_9_to_10 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_9_TO_10), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_9_TO_10))
                        Dim _order_horizon_percentage_11_to_12 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_11_TO_12), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_11_TO_12))
                        Dim _order_horizon_percentage_greather_than_12 As Integer = CInt(IIf(IsNothing(uneLigne.ORDER_HORIZON_PERCENTAGE_GREATER_THAN_12), 0, uneLigne.ORDER_HORIZON_PERCENTAGE_GREATER_THAN_12))
                        _PrecalculedValue = New Precalculatedvalue(PPM, QNC_COUNT, CUSTOMER_CLAIN8COUNT, LNC_COUNT, LOGISTIC_RATE, DELAY_UP_TO_DAYS_RATE, _
                                                                   _order_horizon_percentage_0_to_2, _order_horizon_percentage_3_to_4, _order_horizon_percentage_5_to_6, _order_horizon_percentage_7_to_8, _
                                                                   _order_horizon_percentage_9_to_10, _order_horizon_percentage_11_to_12, _order_horizon_percentage_greather_than_12)
                    Next
                End If
                Return _PrecalculedValue
            Catch ex As Exception
                Throw ex
            End Try
        End Get
    End Property

End Class
