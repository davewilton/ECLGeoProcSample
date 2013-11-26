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
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ADF.CATIDs

<Guid("4e03330f-ddf4-48ad-ad84-38a002f1c57c"), ComVisible(True)> _
Public Class ECLGPFunctionFactory : Implements IGPFunctionFactory

    '=========================================================================================================
    'NOTE Developer does not need to edit this file. Add your implementation of new functions to GPItems class
    '=========================================================================================================

#Region "Component Category Registration"
    ' Register the Function Factory with the ESRI Geoprocessor Function Factory Component Category.
    <ComRegisterFunction()> _
    Private Shared Sub Reg(ByVal regKey As String)
        GPFunctionFactories.Register(regKey)
    End Sub

    <ComUnregisterFunction()> _
    Private Shared Sub Unreg(ByVal regKey As String)
        GPFunctionFactories.Unregister(regKey)
    End Sub
#End Region

#Region "IGPFunctionFactory Members"
    ' Implementation of the Function Factory

    Private _mGPFunction As IGPFunction
    Private ReadOnly _mItems As New GPAllItems

    ' Utility Function added to create the function names.
    Private Function CreateGPFunctionNames(ByVal index As Long) As IGPFunctionName

        Dim item As GPAllItems.GPItem = _mItems.GetItem(CInt(index))

        Dim functionName As IGPFunctionName = New GPFunctionNameClass()
        Dim gprName As IGPName

        gprName = CType(functionName, IGPName)
        gprName.Category = item.Catagory 'The sub group of tools within the toolbox. a backslash adds another tool set
        gprName.Description = item.Description
        gprName.DisplayName = item.DisplayName
        gprName.Name = item.Name
        gprName.Factory = Me

        Return functionName

    End Function

    ' This is the name of the function factory. 
    ' This is used when generating the Toolbox containing the function tools of the factory.
    Public ReadOnly Property Name() As String Implements IGPFunctionFactory.Name
        Get
            Return GPToolBoxName
        End Get
    End Property

    ' This is the alias name of the factory.
    Public ReadOnly Property [Alias]() As String Implements IGPFunctionFactory.Alias
        Get
            Return GPToolBoxAlias
        End Get
    End Property

    ' This is the class id of the factory. 
    Public ReadOnly Property CLSID() As UID Implements IGPFunctionFactory.CLSID
        Get
            Dim id As UID = New UIDClass()
            id.Value = Me.GetType().GUID.ToString("B")
            Return id
        End Get
    End Property

    ' This method will create and return a function object based upon the input name.
    Public Function GetFunction(ByVal sName As String) As IGPFunction Implements IGPFunctionFactory.GetFunction
        _mGPFunction = _mItems.GetItem(_mItems.GetIndex(sName)).Type

        Return _mGPFunction
    End Function

    ' This method will create and return a function name object based upon the input name.
    Public Function GetFunctionName(ByVal sName As String) As IGPName Implements IGPFunctionFactory.GetFunctionName
        Return CType(CreateGPFunctionNames(_mItems.GetIndex(sName)), IGPName)
    End Function

    ' This method will create and return an enumeration of function names that the factory supports.
    Public Function GetFunctionNames() As IEnumGPName Implements IGPFunctionFactory.GetFunctionNames
        Dim iToolCount As Integer = _mItems.Count
        Dim nameArray As IArray = New EnumGPNameClass()
        Dim i As Integer
        For i = 0 To iToolCount - 1
            nameArray.Add(CreateGPFunctionNames(i))
        Next
        Return CType(nameArray, IEnumGPName)
    End Function

    ' This method will create and return an enumeration of GPEnvironment objects. 
    ' If tools published by this function factory required new environment settings, 
    'then you would define the additional environment settings here. 
    ' This would be similar to how parameters are defined. 
    Public Function GetFunctionEnvironments() As IEnumGPEnvironment Implements IGPFunctionFactory.GetFunctionEnvironments
        Return Nothing
    End Function

#End Region

End Class