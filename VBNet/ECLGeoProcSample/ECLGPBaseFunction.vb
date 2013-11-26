Imports System.Diagnostics
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem
Imports System.IO
Imports System.Reflection

Public MustInherit Class ECLGPBaseFunction : Implements IGPFunction2

    ' Local members for GP
    Friend MustOverride ReadOnly Property ToolName As String
    Friend MustOverride ReadOnly Property ToolDisplayName As String
    Friend Parameters As IArray ' Array of Parameters
    Friend ReadOnly GpUtilities As New GPUtilities ' GPUtilities object


#Region "IGPFunction2 Members"

    ' Set the name of the function tool. 
    ' This name appears when executing the tool at the command line or in scripting. 
    ' This name should be unique to each toolbox and must not contain spaces.
    Public ReadOnly Property Name() As String Implements IGPFunction2.Name
        Get
            Return ToolName
        End Get
    End Property

    ' Set the function tool Display Name as seen in ArcToolbox.
    Public ReadOnly Property DisplayName() As String Implements IGPFunction2.DisplayName
        Get
            Return ToolDisplayName
        End Get
    End Property

    Friend ReadOnly Property MetaDataFileName As String
        Get
            Return ToolName & ".xml"
        End Get
    End Property


    ''' <summary>
    ''' Helper function Returns a parameter by name
    ''' </summary>
    ''' <param name="sName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()> _
    Friend Function ByName(sName As String) As Integer
        Dim i As Integer
        For i = 0 To Parameters.Count - 1
            Dim inputParameter As IGPParameter = CType(Parameters.Element(i), IGPParameter)
            If inputParameter.Name.ToUpper() = sName.ToUpper() Then Return i
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Helper function checks if a parameter exists
    ''' </summary>
    ''' <param name="sName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function CheckPerameterExists(sName As String) As Boolean
        Dim i As Integer
        For i = 0 To Parameters.Count - 1
            Dim inputParameter As IGPParameter = CType(Parameters.Element(i), IGPParameter)
            If inputParameter.Name.ToUpper() = sName.ToUpper() Then Return True
        Next
        Return False
    End Function


    ''' <summary>
    ''' location where the parameters to the Function Tool are defined. 
    ''' </summary>
    ''' <value></value>
    ''' <returns>
    ''' This property returns an IArray of parameter objects (IGPParameter). 
    ''' These objects define the characteristics of the input and output parameters. 
    ''' </returns>
    ''' <remarks></remarks>
    MustOverride ReadOnly Property ParameterInfo() As IArray Implements IGPFunction2.ParameterInfo

    ''' <summary>
    ''' Verify parameters
    ''' </summary>
    ''' <param name="paramvalues"></param>
    ''' <param name="pEnvMgr"></param>
    ''' <remarks>
    ''' Called each time a parameter is changed in the tool dialog box or command line. This updates the output data of the tool, 
    ''' which is useful for building models. After returning from UpdateParameters(), geoprocessing calls its internal validation 
    ''' routine, checking that a given set of parameter values are of the appropriate number, data type, and value.
    '''</remarks>
    Public MustOverride Sub UpdateParameters(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal pEnvMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction2.UpdateParameters


    ''' <summary>
    ''' Checks for for errors and adds custom ListMessages
    ''' </summary>
    ''' <param name="paramvalues"></param>
    ''' <param name="pEnvMgr"></param>
    ''' <param name="messages"></param>
    ''' <remarks>
    ''' Called after returning from the internal validation routine. You can examine the ListMessages created from internal validation and 
    ''' change them if desired. 
    ''' </remarks>
    Public MustOverride Sub UpdateMessages(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal pEnvMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager, ByVal messages As ESRI.ArcGIS.Geodatabase.IGPMessages) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction2.UpdateMessages
    ''' <summary>
    ''' Execute the function given the array of the parameters
    ''' </summary>
    ''' <param name="paramvalues"></param>
    ''' <param name="trackcancel"></param>
    ''' <param name="envMgr"></param>
    ''' <param name="message"></param>
    ''' <remarks>
    ''' DEBUG: To debug background processing must be disabled or attach to the process. See
    ''' http://help.arcgis.com/EN/sdk/10.0/ArcObjects_NET/conceptualhelp/index.html#//000100000mqr000000
    ''' </remarks>
    Public MustOverride Sub Execute(ByVal paramvalues As IArray, ByVal trackcancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction2.Execute

    ' This is the function name object for the Geoprocessing Function Tool. 
    ' This name object is created and returned by the Function Factory.
    ' The Function Factory must first be created before implementing this property.
    Public ReadOnly Property FullName() As IName Implements IGPFunction2.FullName
        Get
            ' Add CalculateArea.FullName getter implementation
            Dim functionFactory As IGPFunctionFactory = New ECLGPFunctionFactory()
            Return CType(functionFactory.GetFunctionName(ToolName), IName)
        End Get
    End Property

    ' This is used to set a custom renderer for the output of the Function Tool.
    Public Function GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction2.GetRenderer
        Return Nothing
    End Function


    ''' <summary>
    ''' Required legacy code
    ''' </summary>
    ''' <param name="paramvalues"></param>
    ''' <param name="updateValues"></param>
    ''' <param name="envMgr"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Validate is an IGPFunction method, and we need to implement it in case there
    ''' is legacy code that queries for the IGPFunction interface instead of the IGPFunction2 interface.  
    ''' This Validate code is boilerplate - copy and insert into any IGPFunction2 code.
    ''' This is the calling sequence that the gp framework now uses when it QI's for IGPFunction2..
    ''' </remarks>
    Public Function Validate(ByVal paramvalues As IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction2.Validate
        Return Nothing
    End Function

    ' This is the unique context identifier in a [MAP] file (.h). 
    ' ESRI Knowledge Base article #27680 provides more information about creating a [MAP] file. 
    Public ReadOnly Property HelpContext() As Integer Implements IGPFunction2.HelpContext
        Get
            Return 0
        End Get
    End Property

    ' This is the path to a .chm file which is used to describe and explain the function and its operation. 
    Public ReadOnly Property HelpFile() As String Implements IGPFunction2.HelpFile
        Get
            Return ""
        End Get
    End Property

    ' This is used to return whether the function tool is licensed to execute.
    Public Function IsLicensed() As Boolean Implements IGPFunction2.IsLicensed
        Return True
    End Function

    ' This is the name of the (.xml) file containing the default metadata for this function tool. 
    ' The metadata file is used to supply the parameter descriptions in the help panel in the dialog. 
    ' If no (.chm) file is provided, the help is based on the metadata file. 
    ' ESRI Knowledge Base article #27000 provides more information about creating a metadata file.
    Public ReadOnly Property MetadataFile() As String Implements IGPFunction2.MetadataFile

        ' if you just return the name of an *.xml file as follows:
        ' Get
        '   return m_MetaDataFile
        ' End Get
        ' then the metadata file will be created 
        ' in the default location - <install directory>\help\gp

        ' alternatively, you can send the *.xml file to the location of the DLL.
        ' 
        Get
            Dim filePath As String, fileLocation As String
            fileLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
            'TODO
            filePath = System.IO.Path.Combine(fileLocation, "Resources\GeoprocessingHelp")
            If Not Directory.Exists(filePath) Then
                Directory.CreateDirectory(filePath)
            End If
            filePath = System.IO.Path.Combine(filePath, MetaDataFileName)
            Return filePath
        End Get
    End Property

    ' This is the class id used to override the default dialog for a tool. 
    ' By default, the Toolbox will create a dialog based upon the parameters returned 
    ' by the ParameterInfo property.
    Public ReadOnly Property DialogCLSID() As UID Implements IGPFunction2.DialogCLSID
        Get
            Return Nothing
        End Get
    End Property
#End Region

#Region "IGPFunction Members"

    Public Function GetRenderer1(ByVal pParam As ESRI.ArcGIS.Geoprocessing.IGPParameter) As Object Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.GetRenderer
        Return Nothing
    End Function

    Public ReadOnly Property ParameterInfo1() As ESRI.ArcGIS.esriSystem.IArray Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.ParameterInfo
        Get
            Return ParameterInfo()
        End Get
    End Property

    Public ReadOnly Property DialogCLSID1() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.DialogCLSID
        Get
            Return DialogCLSID
        End Get
    End Property

    Public ReadOnly Property DisplayName1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.DisplayName
        Get
            Return DisplayName
        End Get
    End Property

    Public Sub Execute1(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal trackcancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager, ByVal message As ESRI.ArcGIS.Geodatabase.IGPMessages) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Execute
        Call Execute(paramvalues, trackcancel, envMgr, message)
    End Sub

    Public ReadOnly Property FullName1() As ESRI.ArcGIS.esriSystem.IName Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.FullName
        Get
            FullName1 = FullName
        End Get
    End Property

    Public ReadOnly Property HelpContext1() As Integer Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.HelpContext
        Get
            Return HelpContext
        End Get
    End Property

    Public ReadOnly Property HelpFile1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.HelpFile
        Get
            Return HelpFile
        End Get
    End Property

    Public Function IsLicensed1() As Boolean Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.IsLicensed
        Return IsLicensed()
    End Function

    Public ReadOnly Property MetadataFile1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.MetadataFile
        Get
            Return MetadataFile
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Name
        Get
            Return Name
        End Get
    End Property

    Public Function Validate1(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager) As ESRI.ArcGIS.Geodatabase.IGPMessages Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Validate
        Return Validate(paramvalues, updateValues, envMgr)
    End Function
#End Region


End Class


