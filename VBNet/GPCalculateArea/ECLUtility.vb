Imports System
Imports ESRI.ArcGIS.Geoprocessing
Imports System.IO
Imports Microsoft.Win32

Public Module ECLUtility
    Public Const GPToolBoxName As String = "ECL Geo Proc Sample" 'The name the user sees
    Public Const GPToolBoxAlias As String = "ECLGeoProcSample" 'the abbreviation, used in python

    Public Sub CreateToolbox()

        'WARNING this is slow. Don't run on every start-up. Just regenerate the toolbox and include in version
        'DONT forget to update the version in the toolbox description (can't find a way to do this pragmatically)

        'Dll must be registered with ESRIregasm for this to work

        'this path will ensure it appears in 'My Toolboxes'.
        Dim folder As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\ESRI\Desktop" + KBtargetArcGisVersion + "\ArcToolbox\My Toolboxes\"
        Dim toolBoxLocation As String = Path.Combine(folder, GPToolBoxName & ".tbx")

        Dim pGPUtilities As IGPUtilities3
        pGPUtilities = New GPUtilitiesClass
        Dim fileInfo As New FileInfo(toolBoxLocation)
        Dim dir As New DirectoryInfo(folder)
        If dir.Exists = False Then
            dir.Create()
        End If

        Dim sError As String
        Try
            If fileInfo.Exists Then fileInfo.Delete()
            pGPUtilities.CreateToolboxFromFactory(GPToolBoxAlias, folder, GPToolBoxName)
        Catch ex As Exception
            sError = ex.Message
            Console.WriteLine(sError)
        End Try

    End Sub


    ''' <summary>
    ''' The version of ArcGIS. Used to find the registry key and the app directory path
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property KBtargetArcGisVersion() As String
        Get

            Dim pRegKey As RegistryKey
            Dim agSversion As String = ""
            If Environment.GetEnvironmentVariable("ProgramW6432") = "" Then
                pRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\ESRI", False)
            Else
                pRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\ESRI", False)
            End If
            Dim keys As String() = pRegKey.GetSubKeyNames()
            For Each key As String In keys
                If key.StartsWith("Desktop") Then
                    pRegKey = pRegKey.OpenSubKey(key)
                    agSversion = pRegKey.GetValue("RealVersion").ToString()
                    Exit For
                End If
            Next
            Return agSversion
        End Get
    End Property
End Module
