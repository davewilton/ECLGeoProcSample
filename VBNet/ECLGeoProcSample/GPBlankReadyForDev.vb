Imports System
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Geodatabase

Public Class GPBlankReadyForDev : Inherits ECLGPBaseFunction

    Friend Overrides ReadOnly Property ToolName() As String
        Get
            Return "blankGPreadyForDev"
        End Get
    End Property

    Friend Overrides ReadOnly Property ToolDisplayName() As String
        Get
            Return "Develop this tool"
        End Get
    End Property

    Public Overrides ReadOnly Property ParameterInfo() As IArray
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides Sub Execute(ByVal paramvalues As IArray, ByVal trackcancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages)
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub UpdateParameters(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager)
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub UpdateMessages(ByVal paramvalues As IArray, ByVal pEnvMgr As IGPEnvironmentManager, ByVal messages As IGPMessages)
        Throw New NotImplementedException()
    End Sub

End Class
