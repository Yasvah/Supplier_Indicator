﻿''' <summary>
''' Cette classe permet de gérer (Manager) les scores pour SERTA
''' </summary>
''' <remarks></remarks>
Public Class MgtSupplierIndicatorSERTA
#Region "Propriété"
    ''' <summary>
    ''' Variable utiliser pour le data context
    ''' </summary>
    ''' <remarks></remarks>
    Private BaseSupplierAssessment As dbSupplierIndicatorDataContext = New dbSupplierIndicatorDataContext()

    Dim _listAssessment As List(Of AssessmentSERTA) = New List(Of AssessmentSERTA)
    ''' <summary>
    ''' Liste des score enregistré dans la base de donnée
    ''' </summary>
    ''' <value></value>
    ''' <returns>Retourne une liste de type assessment</returns>
    ''' <remarks></remarks>
    ReadOnly Property ListAssessment As List(Of AssessmentSERTA)
        Get
            _listAssessment.Clear()
            Dim query = From listeAssessment In BaseSupplierAssessment.SUP_ASSESSMENT_SERTA
                   Select listeAssessment

            For Each item In query
                _listAssessment.Add(CAsessement(item))
            Next
            Return _listAssessment
        End Get
    End Property
    ''' <summary>
    ''' Liste des fournisseurs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Retourne le resultat de la procédure stocker Supplier_list</remarks>
    Public ReadOnly Property ListSupplier As System.Data.Linq.ISingleResult(Of P_SUPPLIER_LISTResult1)
        Get
            Try
                Return BaseSupplierAssessment.P_Supplier_list_SERTA()
            Catch ex As Exception
                Throw ex
            End Try
        End Get
    End Property
#End Region
#Region "constructeur"
    Private Shared _instance As New MgtSupplierIndicatorSERTA()

    Private Sub New()
    End Sub

    Public Shared Function getInstance() As MgtSupplierIndicatorSERTA
        Return _instance
    End Function
#End Region
#Region "Méthode"
    ''' <summary>
    ''' Converti la classe SUP_ASSESSMENT_??? en Assessment
    ''' </summary>
    ''' <param name="Assessment"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CAsessement(Assessment As SUP_ASSESSMENT_SERTA) As AssessmentSERTA

        Return New AssessmentSERTA(IIf(IsDBNull(Assessment.SUA_ID), 0, Assessment.SUA_ID), _
                              (IIf(IsDBNull(Assessment.SUP_ID), 0, Assessment.SUP_ID)), _
                              (IIf(IsDBNull(Assessment.SUA_QUARTER), 0, Assessment.SUA_QUARTER)), _
                              (IIf(IsDBNull(Assessment.SUA_FORM), 0, Assessment.SUA_FORM)), _
                              (IIf(IsDBNull(Assessment.SUA_INDICE_PPM_VALUE), 0, Assessment.SUA_INDICE_PPM_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_INDICE_PPM_POINT), 0, Assessment.SUA_INDICE_PPM_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_SIN_NB_VALUE), 0, Assessment.SUA_SIN_NB_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_SIN_NB_POINT), 0, Assessment.SUA_SIN_NB_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_CUSTOMER_CLAIM_NB_VALUE), 0, Assessment.SUA_CUSTOMER_CLAIM_NB_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_CUSTOMER_CLAIM_NB_POINT), 0, Assessment.SUA_CUSTOMER_CLAIM_NB_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_ACTION_PLAN_REACTIVITY_POINT), 0, Assessment.SUA_ACTION_PLAN_REACTIVITY_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_BONUS_500_PPM_POINT), 0, Assessment.SUA_BONUS_500_PPM_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_LOGISTIC_RATE_TARGET_95_VALUE), 0, Assessment.SUA_LOGISTIC_RATE_TARGET_95_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_LOGISTIC_RATE_TARGET_95_POINT), 0, Assessment.SUA_LOGISTIC_RATE_TARGET_95_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_FLEXIBILITY_POINT), 0, Assessment.SUA_FLEXIBILITY_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_DELIVERY_DELAYS_LEVEL_VALUE), 0, Assessment.SUA_DELIVERY_DELAYS_LEVEL_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_DELIVERY_DELAYS_LEVEL_POINT), 0, Assessment.SUA_DELIVERY_DELAYS_LEVEL_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_DELIVERY_QUALITY_VALUE), 0, Assessment.SUA_DELIVERY_QUALITY_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_DELIVERY_QUALITY_POINT), 0, Assessment.SUA_DELIVERY_QUALITY_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_PRICE_COMPETITIVENESS_VALUE), 0, Assessment.SUA_PRICE_COMPETITIVENESS_VALUE)), _
                              (IIf(IsDBNull(Assessment.SUA_PRICE_COMPETITIVENESS_POINT), 0, Assessment.SUA_PRICE_COMPETITIVENESS_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_IMPROVMENT_PLAN_POINT), 0, Assessment.SUA_IMPROVMENT_PLAN_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_BUSINESS_RELATIONSHIP_POINT), 0, Assessment.SUA_BUSINESS_RELATIONSHIP_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_FINANCIAL_SITUATION_POINT), 0, Assessment.SUA_FINANCIAL_SITUATION_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_OFFERS_REACTIVITY_POINT), 0, Assessment.SUA_OFFERS_REACTIVITY_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_TECHNICAL_ANSWER_QUALITY_POINT), 0, Assessment.SUA_TECHNICAL_ANSWER_QUALITY_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_ISO_CERTFICATION_POINT), 0, Assessment.SUA_ISO_CERTFICATION_POINT)), _
                              (IIf(IsDBNull(Assessment.SUA_COMMENT), "", Assessment.SUA_COMMENT)), _
                              (IIf(IsDBNull(Assessment.SUA_COMMENT_DETAIL), "", Assessment.SUA_COMMENT_DETAIL)), _
                              (IIf(IsDBNull(Assessment.SUA_COMMENT_CLASSIFICATION), "", Assessment.SUA_COMMENT_CLASSIFICATION)), _
                              (IIf(IsDBNull(Assessment.SUA_COMMENT_GLOBAL), "", Assessment.SUA_COMMENT_GLOBAL)), _
                              (IIf(IsDBNull(Assessment.SUA_TOTAL_POINT), 0, Assessment.SUA_TOTAL_POINT)), _
                              CInt((IIf(IsDBNull(Assessment.SUA_TREND), 0, Assessment.SUA_TREND.ToString))))
    End Function
    ''' <summary>
    ''' Function qui sauvegarde un assessment
    ''' </summary>
    ''' <param name="assessment"></param>
    ''' <remarks>Met à jour le score si il exist, sinon crée une insertion</remarks>
    Public Sub Save(assessment As AssessmentSERTA)
        If Not IsNothing(recherche(assessment.idSupplier, assessment.quarter)) Then
            Update(assessment)
        Else
            Insert(assessment)
        End If
    End Sub
    ''' <summary>
    ''' Insertion d'un nouveau score dans la base de donnée
    ''' </summary>
    ''' <param name="Assessment"></param>
    ''' <remarks></remarks>
    Public Sub Insert(Assessment As AssessmentSERTA)
        Dim nouveauAssessment As SUP_ASSESSMENT_SERTA = New SUP_ASSESSMENT_SERTA
        With nouveauAssessment
            .SUP_ID = Assessment.idSupplier
            .SUA_QUARTER = Assessment.quarter
            .SUA_FORM = Assessment.form
            .SUA_INDICE_PPM_VALUE = Assessment.indicePPMValue
            .SUA_INDICE_PPM_POINT = Assessment.indicePPMPoint
            .SUA_SIN_NB_VALUE = Assessment.sinNBValue
            .SUA_SIN_NB_POINT = Assessment.sinNBPoint
            .SUA_CUSTOMER_CLAIM_NB_VALUE = Assessment.customerClaimNBValue
            .SUA_CUSTOMER_CLAIM_NB_POINT = Assessment.customerClaimNBPoint
            .SUA_ACTION_PLAN_REACTIVITY_POINT = Assessment.actionPlanReactivityPoint
            .SUA_BONUS_500_PPM_POINT = Assessment.bonus500PPMPoint
            .SUA_LOGISTIC_RATE_TARGET_95_VALUE = Assessment.logisticRateTarget95Value
            .SUA_LOGISTIC_RATE_TARGET_95_POINT = Assessment.logisticRateTarget95Point
            .SUA_FLEXIBILITY_POINT = Assessment.flexibilityPoint
            .SUA_DELIVERY_DELAYS_LEVEL_VALUE = Assessment.deliveryDelaysLevelValue
            .SUA_DELIVERY_DELAYS_LEVEL_POINT = Assessment.deliveryDelaysLevelPoint
            .SUA_DELIVERY_QUALITY_VALUE = Assessment.deliveryQualityValue
            .SUA_DELIVERY_QUALITY_POINT = Assessment.deliveryQualityPoint
            .SUA_PRICE_COMPETITIVENESS_VALUE = Assessment.priceCompetitivenessValue
            .SUA_PRICE_COMPETITIVENESS_POINT = Assessment.priceCompetitivenessPoint
            .SUA_IMPROVMENT_PLAN_POINT = Assessment.improvmentPlanPoint
            .SUA_BUSINESS_RELATIONSHIP_POINT = Assessment.businessRelationshipPoint
            .SUA_FINANCIAL_SITUATION_POINT = Assessment.financialSituationPoint
            .SUA_OFFERS_REACTIVITY_POINT = Assessment.offersReactivityPoint
            .SUA_TECHNICAL_ANSWER_QUALITY_POINT = Assessment.technicalAnswerQualityPoint
            .SUA_ISO_CERTFICATION_POINT = Assessment.isoCertificationPoint
            .SUA_COMMENT = Assessment.comment
            .SUA_COMMENT_CLASSIFICATION = Assessment.commentClassification
            .SUA_COMMENT_DETAIL = Assessment.commentDetail
            .SUA_COMMENT_GLOBAL = Assessment.commentGlobal
            .SUA_TOTAL_POINT = Assessment.totalPoint
            .SUA_TREND = CStr(Assessment.trend)
        End With
        BaseSupplierAssessment.SUP_ASSESSMENT_SERTA.InsertOnSubmit(nouveauAssessment)
        BaseSupplierAssessment.SubmitChanges()
    End Sub
    ''' <summary>
    ''' Mise à jour d'un score dans la base de donnée
    ''' </summary>
    ''' <param name="Assessment"></param>
    ''' <remarks></remarks>
    Public Sub Update(Assessment As AssessmentSERTA)
        Dim modificationAssessment = From unAssessment In BaseSupplierAssessment.SUP_ASSESSMENT_SERTA
                                                             Where unAssessment.SUA_ID = Assessment.id
                                                             Select unAssessment

        For Each unAssessment In modificationAssessment
            With unAssessment
                .SUA_FORM = Assessment.form
                .SUA_INDICE_PPM_VALUE = Assessment.indicePPMValue
                .SUA_INDICE_PPM_POINT = Assessment.indicePPMPoint
                .SUA_SIN_NB_VALUE = Assessment.sinNBValue
                .SUA_SIN_NB_POINT = Assessment.sinNBPoint
                .SUA_CUSTOMER_CLAIM_NB_VALUE = Assessment.customerClaimNBValue
                .SUA_CUSTOMER_CLAIM_NB_POINT = Assessment.customerClaimNBPoint
                .SUA_ACTION_PLAN_REACTIVITY_POINT = Assessment.actionPlanReactivityPoint
                .SUA_BONUS_500_PPM_POINT = Assessment.bonus500PPMPoint
                .SUA_LOGISTIC_RATE_TARGET_95_VALUE = Assessment.logisticRateTarget95Value
                .SUA_LOGISTIC_RATE_TARGET_95_POINT = Assessment.logisticRateTarget95Point
                .SUA_FLEXIBILITY_POINT = Assessment.flexibilityPoint
                .SUA_DELIVERY_DELAYS_LEVEL_VALUE = Assessment.deliveryDelaysLevelValue
                .SUA_DELIVERY_DELAYS_LEVEL_POINT = Assessment.deliveryDelaysLevelPoint
                .SUA_DELIVERY_QUALITY_VALUE = Assessment.deliveryQualityValue
                .SUA_DELIVERY_QUALITY_POINT = Assessment.deliveryQualityPoint
                .SUA_PRICE_COMPETITIVENESS_VALUE = Assessment.priceCompetitivenessValue
                .SUA_PRICE_COMPETITIVENESS_POINT = Assessment.priceCompetitivenessPoint
                .SUA_IMPROVMENT_PLAN_POINT = Assessment.improvmentPlanPoint
                .SUA_BUSINESS_RELATIONSHIP_POINT = Assessment.businessRelationshipPoint
                .SUA_FINANCIAL_SITUATION_POINT = Assessment.financialSituationPoint
                .SUA_OFFERS_REACTIVITY_POINT = Assessment.offersReactivityPoint
                .SUA_TECHNICAL_ANSWER_QUALITY_POINT = Assessment.technicalAnswerQualityPoint
                .SUA_ISO_CERTFICATION_POINT = Assessment.isoCertificationPoint
                .SUA_COMMENT = Assessment.comment
                .SUA_COMMENT_CLASSIFICATION = Assessment.commentClassification
                .SUA_COMMENT_DETAIL = Assessment.commentDetail
                .SUA_COMMENT_GLOBAL = Assessment.commentGlobal
                .SUA_TOTAL_POINT = Assessment.totalPoint
                .SUA_TREND = CStr(Assessment.trend)
            End With
        Next
        BaseSupplierAssessment.SubmitChanges()

    End Sub
    ''' <summary>
    ''' Recherche un score dans la liste des scores
    ''' </summary>
    ''' <param name="Id">identifient du score</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function recherche(Id As Integer) As AssessmentSERTA
        Dim trouverAssessment = From unAssessment In BaseSupplierAssessment.SUP_ASSESSMENT_SERTA
        Where unAssessment.SUA_ID = Id
                          Select unAssessment
        If trouverAssessment.Any Then
            Return CAsessement(trouverAssessment.First)
        Else
            Return Nothing
        End If
    End Function
    ''' <summary>
    ''' Recherche un score das la liste des scores
    ''' </summary>
    ''' <param name="IdSupplier">Identificateur du fournisseur</param>
    ''' <param name="quarter">Numéro du trimestre</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function recherche(IdSupplier As Integer, quarter As Integer) As AssessmentSERTA
        Dim trouverAssessment = From unAssessment In BaseSupplierAssessment.SUP_ASSESSMENT_SERTA
                            Where unAssessment.SUP_ID = IdSupplier And unAssessment.SUA_QUARTER = quarter
                           Select unAssessment
        If trouverAssessment.Any Then
            Return CAsessement(trouverAssessment.First)
        Else
            Return Nothing
        End If
    End Function
#End Region
End Class
