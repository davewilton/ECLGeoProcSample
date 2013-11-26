Imports ECLUtility

Public Class ECLGeoProExtension
    Inherits ESRI.ArcGIS.Desktop.AddIns.Extension

    Public Sub New()

    End Sub

    Protected Overrides Sub OnStartup()
        '
        ' TODO: Uncomment to start listening to document events
        'WireDocumentEvents()

        'Create the toolbox. WARNING this is slow, don't call on every start up in production
        CreateToolbox()

    End Sub

    Protected Overrides Sub OnShutdown()

    End Sub

    Private Sub WireDocumentEvents()
        '
        ' TODO: Sample document event wiring code. Change as needed.
        '
        AddHandler My.ArcMap.Events.NewDocument, AddressOf ArcMapNewDocument

    End Sub

    Private Sub ArcMapNewDocument()

    End Sub
End Class
