' Copyright 2011 ESRI
' 
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
' 
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
' 
' See the use restrictions at http://resourcesbeta.arcgis.com/en/help/arcobjects-net/usagerestrictions.htm
' 
Imports System
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem

Public Class GPCalculateAreaFunction : Inherits ECLGPBaseFunction

    ' Set the name of the function tool. 
    ' This name appears when executing the tool at the command line or in scripting. 
    ' This name should be unique to each toolbox and must not contain spaces.
    Friend Overrides ReadOnly Property ToolName() As String
        Get
            Return "CalculateArea"
        End Get
    End Property

    Friend Overrides ReadOnly Property ToolDisplayName() As String
        Get
            Return "Calculate Area"
        End Get
    End Property

    ' This is the location where the parameters to the Function Tool are defined. 
    ' This property returns an IArray of parameter objects (IGPParameter). 
    ' These objects define the characteristics of the input and output parameters. 
    Public Overrides ReadOnly Property ParameterInfo() As IArray
        Get
            'Array to the hold the parameters
            Dim pParameters As IArray = New ArrayClass()

            'Input DataType is GPFeatureLayerType
            Dim inputParameter As IGPParameterEdit3 = New GPParameterClass()
            inputParameter.DataType = New GPFeatureLayerTypeClass()

            ' Default Value object is DEFeatureClass
            inputParameter.Value = New GPFeatureLayerClass()

            ' Set Input Parameter properties
            inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput
            inputParameter.DisplayName = "Input Features"
            inputParameter.Name = "input_features"
            inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired
            pParameters.Add(inputParameter)

            ' Area field parameter
            inputParameter = New GPParameterClass()
            inputParameter.DataType = New GPStringTypeClass()

            ' Value object is GPString
            Dim gpStringValue As IGPString = New GPStringClass()
            gpStringValue.Value = "Area"
            inputParameter.Value = CType(gpStringValue, IGPValue)

            ' Set field name parameter properties
            inputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionInput
            inputParameter.DisplayName = "Area Field Name"
            inputParameter.Name = "field_name"
            inputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeRequired

            pParameters.Add(inputParameter)

            ' Output parameter (Derived) and data type is DEFeatureClass
            Dim outputParameter As IGPParameterEdit3 = New GPParameterClass()
            outputParameter.DataType = New GPFeatureLayerTypeClass()

            ' Value object is DEFeatureClass
            outputParameter.Value = CType(New DEFeatureClass(), IGPValue)

            'Create a new feature schema object
            Dim featureSchema As IGPFeatureSchema
            featureSchema = CType(New GPFeatureSchema, IGPFeatureSchema)
            Dim schema As IGPSchema
            schema = CType(featureSchema, IGPSchema)

            'Clone the dependency
            schema.CloneDependency = True

            ' Set output parameter properties
            outputParameter.Direction = esriGPParameterDirection.esriGPParameterDirectionOutput
            outputParameter.DisplayName = "Output FeatureClass"
            outputParameter.Name = "out_featureclass"
            outputParameter.ParameterType = esriGPParameterType.esriGPParameterTypeDerived
            outputParameter.Schema = schema
            outputParameter.AddDependency("input_features")
            pParameters.Add(outputParameter)

            Return pParameters
        End Get
    End Property

    ' Called after returning from the internal validation routine. You can examine the messages created from internal validation and change them if desired. 
    Public Overrides Sub UpdateParameters(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager)
        Parameters = paramvalues

        ' Retrieve the input parameter value
        Dim parameterValue As IGPValue
        parameterValue = GpUtilities.UnpackGPValue(paramvalues.Element(ByName("input_features")))

        ' Get the derived output feature class schema and empty the additional fields.
        ' This will ensure you don't get duplicate entries
        Dim derivedFeatures As IGPParameter3
        derivedFeatures = CType(paramvalues.Element(ByName("out_featureclass")), IGPParameter3)

        Dim schema As IGPFeatureSchema
        schema = CType(derivedFeatures.Schema, IGPFeatureSchema)
        schema.AdditionalFields = Nothing

        ' If we have an input value, create a new field based on the field name the user entered
        If parameterValue.IsEmpty() = False Then
            Dim fieldNameParameter As IGPParameter3
            fieldNameParameter = CType(paramvalues.Element(ByName("field_name")), IGPParameter3)

            Dim fieldName As String
            fieldName = fieldNameParameter.Value.GetAsText()

            ' Check if the user's entered value already exists
            Dim areaField As IField
            areaField = GpUtilities.FindField(parameterValue, fieldName)

            If areaField Is Nothing Then
                Dim fieldsEdit As IFieldsEdit
                fieldsEdit = CType(New Fields, IFieldsEdit)
                Dim fieldEdit As IFieldEdit
                fieldEdit = CType(New Field, IFieldEdit)

                fieldEdit.Name_2 = fieldName
                fieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble

                fieldsEdit.AddField(fieldEdit)

                ' Add an additional field for the area values to the derived output
                Dim pFields As IFields
                pFields = fieldsEdit
                schema.AdditionalFields = pFields
            End If
        End If


    End Sub

    ' Called after returning from the internal validation routine. You can examine the ListMessages created from internal validation and change them if desired. 
    Public Overrides Sub UpdateMessages(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager, ByVal messages As IGPMessages)

        ' Check for error messages
        Dim msg As IGPMessage
        msg = CType(messages, IGPMessage)

        If msg.IsError() Then
            Return
        End If

        ' Get the first input parameter
        Dim parameter As IGPParameter
        parameter = CType(paramvalues.Element(0), IGPParameter)

        ' UnPackGPValue. This ensures you get the value either from 
        ' the DataElement or from GPVaraible (ModelBuilder)
        Dim parameterValue As IGPValue
        parameterValue = GpUtilities.UnpackGPValue(parameter)

        ' Open the Input Dataset - use DecodeFeatureLayer as the input might be
        ' a layer file or a feature layer from ArcMap
        Dim inputFeatureClass As IFeatureClass = Nothing
        Dim qf As IQueryFilter = Nothing
        GpUtilities.DecodeFeatureLayer(parameterValue, inputFeatureClass, qf)

        Dim fieldParameter As IGPParameter3
        fieldParameter = CType(paramvalues.Element(1), IGPParameter3)
        Dim fieldName As String
        fieldName = fieldParameter.Value.GetAsText()

        ' Check if the field already exists and provide a warning
        Dim indexA As Integer
        indexA = inputFeatureClass.FindField(fieldName)

        If indexA > 0 Then
            messages.ReplaceWarning(1, "Field already exists. It will be overwritten.")
        End If

        Return

    End Sub

    ' Execute: Execute the function given the array of the parameters
    Public Overrides Sub Execute(ByVal paramvalues As IArray, ByVal trackcancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages)

        ' Get the first Input Parameter
        Dim parameter As IGPParameter = CType(paramvalues.Element(0), IGPParameter)

        ' UnPackGPValue. This ensures you get the value either form the data element or GpVariable (ModelBuilder)
        Dim parameterValue As IGPValue = GpUtilities.UnpackGPValue(parameter)

        ' Open the Input Dataset - use DecodeFeatureLayer as the input might be
        ' a layer file or a feature layer from ArcMap
        Dim inputFeatureClass As IFeatureClass = Nothing
        Dim qf As IQueryFilter = Nothing
        GpUtilities.DecodeFeatureLayer(parameterValue, inputFeatureClass, qf)

        If inputFeatureClass Is Nothing Then
            message.AddError(2, "Could not open input dataset.")
            Return
        End If

        ' Add the field if it does not exist.
        Dim indexA As Integer
        parameter = CType(paramvalues.Element(1), IGPParameter)
        Dim sField As String = parameter.Value.GetAsText()

        indexA = inputFeatureClass.FindField(sField)
        If indexA < 0 Then
            Dim fieldEdit As IFieldEdit = New FieldClass()
            fieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            fieldEdit.Name_2 = sField
            message.AddMessage(sField)
            inputFeatureClass.AddField(fieldEdit)
        End If

        Dim featcount As Integer
        featcount = inputFeatureClass.FeatureCount(Nothing)

        ' Set the properties of the Step Progressing
        Dim pStepPro As IStepProgressor = Nothing
        If Not trackcancel Is Nothing Then
            pStepPro = CType(trackcancel, IStepProgressor)
            pStepPro.MinRange = 0
            pStepPro.MaxRange = featcount
            pStepPro.StepValue = 1
            pStepPro.Message = "Calculate Area"
            pStepPro.Position = 0
            pStepPro.Show()
        End If

        ' Create an Update Cursor
        indexA = inputFeatureClass.FindField(sField)
        Dim updateCursor As IFeatureCursor = inputFeatureClass.Update(Nothing, False)
        Dim updateFeature As IFeature = updateCursor.NextFeature()
        Dim geometry As IGeometry
        Dim area As IArea
        Dim dArea As Double

        Do While Not updateFeature Is Nothing
            geometry = updateFeature.Shape
            area = CType(geometry, IArea)
            dArea = area.Area
            updateFeature.Value(indexA) = dArea
            updateCursor.UpdateFeature(updateFeature)
            updateFeature.Store()
            updateFeature = updateCursor.NextFeature()
            If Not pStepPro Is Nothing Then pStepPro.Step()
        Loop

        If Not pStepPro Is Nothing Then pStepPro.Hide()

        ' Release the update cursor to remove the lock on the input data.
        Runtime.InteropServices.Marshal.ReleaseComObject(updateCursor)

    End Sub

End Class

