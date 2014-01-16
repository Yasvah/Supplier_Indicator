Public Class AssessmentCOMMUN
    Inherits Assessment

    Public Sub New(idSupplier As Integer, quarter As Integer)
        MyBase.New(idSupplier, quarter)
    End Sub
    Public Sub New(id As Integer, supplier As Integer, quarter As String, indicePPMValue As Integer, _
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
    Public ReadOnly Property Assessment_Values_Commun() As Precalculatedvalue
        Get
            Dim db As dbSupplierIndicatorDataContext = New dbSupplierIndicatorDataContext
            Return db.P_Assessment_Values_Commun(Me.idSupplier, Me.quarter)
        End Get
    End Property
    Private _PrecalculedValue As Precalculatedvalue
    Public ReadOnly Property PrecalculedValue As Precalculatedvalue
        Get
            If IsNothing(_PrecalculedValue) Then
                Dim db As dbSupplierIndicatorDataContext = New dbSupplierIndicatorDataContext
                For Each uneLigne In db.P_Assessment_Values_Commun(Me.idSupplier, Me.quarter)
                    Dim PPM = IIf(IsNothing(uneLigne.PPM_GROUP), 0, uneLigne.PPM_GROUP)
                    Dim QNC_COUNT = uneLigne.QNC_COUNT_GROUP
                    Dim CUSTOMER_CLAIN8COUNT = uneLigne.CUSTOMER_CLAIM_COUNT_GROUP
                    Dim LNC_COUNT = uneLigne.LNC_COUNT_GROUP
                    Dim LOGISTIC_RATE = IIf(IsNothing(uneLigne.LOGISTIC_RATE_GROUP), 0, uneLigne.LOGISTIC_RATE_GROUP)
                    Dim DELAY_UP_TO_DAYS_RATE = IIf(IsNothing(uneLigne.DELAYS_UP_TO_10_DAYS_RATE_GROUP), 0, uneLigne.DELAYS_UP_TO_10_DAYS_RATE_GROUP)
                    _PrecalculedValue = New Precalculatedvalue(PPM, QNC_COUNT, CUSTOMER_CLAIN8COUNT, LNC_COUNT, LOGISTIC_RATE, DELAY_UP_TO_DAYS_RATE, 0, 0, 0, 0, 0, 0, 0)
                Next
            End If
            Return _PrecalculedValue
        End Get
    End Property
End Class

