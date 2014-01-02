Imports System
Imports ESRI.ArcGIS.Geoprocessing
Imports System.IO
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Reflection
Imports System.Threading

Public Module ECLUtility
    Public Const GPToolBoxName As String = "ECL Geo Proc Sample" 'The name the user sees
    Public Const GPToolBoxAlias As String = "ECLGeoProcSample" 'the abbreviation, used in python

    Public Sub CreateToolbox()
        'WARNING this is slow. Don't run on every start-up if not needed.  Consider running in a thread
        'WARNING Dll must be registered with ESRIregasm for this to work
        Try

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

            If fileInfo.Exists Then fileInfo.Delete()
            pGPUtilities.CreateToolboxFromFactory(GPToolBoxAlias, folder, GPToolBoxName)

            RegisterWithArcPy()

        Catch ex As Exception
            Console.WriteLine(ex.Message) 'silent, this is run on start-up
        End Try

    End Sub

    ''' <summary>
    ''' Registers tool with ArcPy so that it appears in the users toolbox
    ''' </summary>
    ''' <remarks>Slow to run</remarks>
    Private Sub RegisterWithArcPy()
        'This is Slow to run, consider when it should be run. It will not work if run in a background thread on start-up.
        'Run when the user interacts with another part of application

        Dim location As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)
        Dim regTb As String = IO.Path.Combine(location, "Resources") & "\regarcpy.tbx"

        Dim desTbx As String
        'this path will ensure it appears in 'My Toolboxes'.
        Dim folder As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\ESRI\Desktop" + KBtargetArcGisVersion + "\ArcToolbox\My Toolboxes\"
        desTbx = Path.Combine(folder, GPToolBoxName & ".tbx")

        Dim gp As IGeoProcessor = New GeoProcessor
        gp.AddToolbox(regTb)
        Dim iv As IVariantArray = New VarArrayClass()
        iv.Add(desTbx)
        gp.Execute("registerTbx", iv, Nothing)
        gp.RemoveToolbox(regTb)

        Marshal.ReleaseComObject(gp)

    End Sub

    ''' <summary> The version of ArcGIS. Used to find the registry key and the app directory path </summary>
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
