﻿Public Class AssessmentPNS
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


    Private _PrecalculedValue As Precalculatedvalue
    ''' <summary>
    ''' Retourne les valeur pré_calculer dans la procédure stocker
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrecalculedValue As Precalculatedvalue
        Get
            If IsNothing(_PrecalculedValue) Then
                Dim db As dbSupplierIndicatorDataContext = New dbSupplierIndicatorDataContext
                For Each uneLigne In db.P_Assessment_Values_PNS(Me.idSupplier, Me.quarter)
                    Dim PPM = If(IsNothing(uneLigne.PPM), 0, uneLigne.PPM)
                    Dim QNC_COUNT = uneLigne.QNC_COUNT
                    Dim CUSTOMER_CLAIN8COUNT = uneLigne.CUSTOMER_CLAIM_COUNT
                    Dim LNC_COUNT = uneLigne.LNC_COUNT
                    Dim LOGISTIC_RATE = IIf(IsNothing(uneLigne.LOGISTIC_RATE), 0, uneLigne.LOGISTIC_RATE)
                    Dim DELAY_UP_TO_DAYS_RATE = IIf(IsNothing(uneLigne.DELAYS_UP_TO_10_DAYS_RATE), 0, uneLigne.DELAYS_UP_TO_10_DAYS_RATE)
                    _PrecalculedValue = New Precalculatedvalue(PPM, QNC_COUNT, CUSTOMER_CLAIN8COUNT, LNC_COUNT, LOGISTIC_RATE, DELAY_UP_TO_DAYS_RATE, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    'TODO: A modifier Pour Moving-Ahead 2014
                Next
            End If
            Return _PrecalculedValue
        End Get
    End Property
End Class
