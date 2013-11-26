Public Class ECLGeoProExtension
    Inherits ESRI.ArcGIS.Desktop.AddIns.Extension

    Public Sub New()

    End Sub

    Protected Overrides Sub OnStartup()

        'Create the toolbox. WARNING this is slow, don't call on every start up in production
        CreateToolbox()

    End Sub

    Protected Overrides Sub OnShutdown()

    End Sub

End Class
