Imports ESRI.ArcGIS.Geoprocessing
Imports System.Collections.Generic


''' <summary>
''' To add new items to the toolbox just add them to the constructor of this class.
''' </summary>
''' <remarks></remarks>
Public Class GPItems

    Private ReadOnly _mItems As New List(Of GPItem)

    ''' <summary>
    ''' Add new items here
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        Dim item As IGPFunction2

        'TODO Add your list of items here
        item = New CalculateAreaFunction()
        _mItems.Add(New GPItem(item.Name, "Folder\SubFolder", "", item.DisplayName, item))
        item = New BlankGPreadyForDev()
        _mItems.Add(New GPItem(item.Name, "Folder\SubFolder", "", item.DisplayName, item))

    End Sub

#Region "Helper methods for factory, these can be ignored"
    Public Function GetIndex(sname As String) As Integer
        Dim i As Integer = _mItems.FindIndex(Function(gpItem) (gpItem.Name = sname))
        Return i
    End Function

    Public Function GetItem(index As Integer) As GPItem
        Return _mItems.Item(index)
    End Function

    Public ReadOnly Property Count As Integer
        Get
            Return _mItems.Count
        End Get
    End Property

    Public Class GPItem
        Public ReadOnly Name As String
        Public ReadOnly Catagory As String
        Public ReadOnly Description As String
        Public ReadOnly DisplayName As String
        Public ReadOnly Type As IGPFunction2
        Sub New(sName As String, sCat As String, sDesc As String, sDisp As String, gpType As IGPFunction2)
            Name = sName
            Catagory = sCat
            Description = sDesc
            DisplayName = sDisp
            Type = gpType
        End Sub
    End Class
#End Region

End Class


