Imports HSCF
Imports HomeSeerAPI
Imports Scheduler
Imports HSCF.Communication.ScsServices.Service
Imports System.Reflection
Imports System.Text
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Threading
Imports System.Globalization

'
' Declare Global Variables
'
Module SonosV3SpecificGlobalDefinitions
    Public hs As HomeSeerAPI.IHSApplication
    Public callback As HomeSeerAPI.IAppCallbackAPI

    Public WithEvents MyZonePlayer As HSPI
    Public IFACE_NAME As String = "Sonos"                   ' This is the plugin's name
    Public tIFACE_NAME As String = "Sonos"
    Public sIFACE_NAME As String = "Sonos" '"SONOSCONTROLLER"
    Public ShortIfaceName As String = "Sonos"
    Public MyINIFile As String = tIFACE_NAME & ".ini"    ' Configuration File
    Public Const ConfigPage = "SonosConfig"
    Public Const PlayerControlPage = "PlayerControl"
    Public Const UPnPViewPage As String = "SonosViewer"
End Module


Public Class HSPI
    Inherits ScsService
    Implements IPlugInAPI


    Public instance As String = ""
    Public isRoot As Boolean = True

    Private ConfigurationPage As SonosConfig  ' a jquery web page
    Public MyHSDeviceLinkedList As New LinkedList(Of MyUPnpDeviceInfo)

    Friend WithEvents MyControllerTimer As Timers.Timer
    Friend WithEvents MyAnnouncementTimer As Timers.Timer
    Friend WithEvents MyDBCreationTimer As Timers.Timer

    Const MaxTOActionArray = 10
    Private MyTimeoutActionArray(MaxTOActionArray) As Integer
    '
    ' timeout indexes
    ' 
    Private Const SQLFileName As String = "System.Data.SQLite.dll"
    Private Const LinuxSQLFileName As String = "LinuxSystem.Data.SQLite.dll"
    Private Const WindowsSQLFileName As String = "WindowsSystem.Data.SQLite.dll"
    Const TrialPhase As Boolean = False
    Const TrialLastDate = "Mar 31, 2015 12:00:00 AM" '"#12-31-2013#"
    Const TORediscover = 0
    Const TORediscoverValue = 600 ' 10 minutes  changed in v3.1.0.27 from 5 min to 10 because discovery would kick in when init was still ongoing
    Const TOCheckChange = 1
    Const TOCheckChangeValue = 600 ' 10 minutes
    Const TOCheckAnnouncement = 2
    Const TOCheckAnnouncementValue = 1
    Const TOAddNewDevice = 3
    Const TOAddNewDeviceValue = 10   ' seconds

    Private AutoUpdate As Boolean = False
    Private AutoUpdateTime As String = ""
    Private DBZoneName As String = ""
    Private MyLinkState As Boolean = False
    Private NbrOfSonosPlayers As Integer = 0
    Private ZoneCount As Integer = 0
    Private ActAsSpeakerProxy As Boolean = False
    Private ProxySpeakerActive As Boolean = False
    Private MyLinkgroupArray As New LinkGroupInfoArray
    Private CapabilitiesCalledFlag As Boolean = False
    Private MyPingAddressLinkedList As New LinkedList(Of PingArrayElement)
    Private MyPingReEntry As Boolean = False
    Private MyPostAnnouncementAction As PostAnnouncementAction = PostAnnouncementAction.paaForwardNoMatch
    Private MyNewDiscoveredDeviceQueue As Queue(Of String) = New Queue(Of String)()
    Private NewDeviceHandlerReEntryFlag As Boolean = False
    'Private MissedNewDeviceNotificationHandlerFlag As Boolean = False
    'Private AddDeviceFlag As Boolean = False
    Private UPnPViewerPage As UPnPDebugWindow
    Private MediaAPIEnabled As Boolean = False
    Private ConditionSetFlag As Boolean = False

    Const LinkGroupButtonOffset As Integer = 1006

    Private mvarActionAdvanced As Boolean
    Private MyConfigDevice As PlayerControl
    Private InitDeviceFlag As Boolean = False

    Private MyPlayerControlWebPage As PlayerControl

    Const TriggersPageName As String = "Events"
    Const ActionsPageName As String = "Events"



#Region "Common Interface"

    Public Function Search(ByVal SearchString As String, ByVal RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements IPlugInAPI.Search
        ' Not yet implemented in the Sample
        '
        ' Normally we would do a search on plug-in actions, triggers, devices, etc. for the string provided, using
        '   the string as a regular expression if RegEx is True.
        '
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Search called for instance = " & instance & " and SearchString = " & SearchString & " and RegEx = " & RegEx.ToString, LogType.LOG_TYPE_INFO)
        Return Nothing
    End Function
    Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction
        ' a custom call to call a specific procedure in the plugin
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginFunction called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception

        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As MethodInfo = ty.GetMethod(proc)
            If mi Is Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Method " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
                Return Nothing
            End If
            Return (mi.Invoke(Me, parms))
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PluginProc for instance = " & instance & " : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return Nothing
    End Function
    Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginPropertyGet called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Method " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
                Return Nothing
            End If
            Return mi.GetValue(Me, Nothing)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PluginProc for instance = " & instance & " : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return Nothing
    End Function
    Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet
        Try
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PluginPropertySet called for instance = " & instance & " and proc = " & proc.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Property " & proc & " does not exist in this plugin.", LogType.LOG_TYPE_ERROR)
            End If
            mi.SetValue(Me, value, Nothing)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PluginPropertySet for instance = " & instance & " : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Name Called for Instance = " & instance & " and returned = " & tIFACE_NAME, LogType.LOG_TYPE_INFO)
            Return sIFACE_NAME
        End Get
    End Property

    Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
        If instance = "" Then   ' rewritten 5/27/2020 to prevent error at startup when HS calls this function before any init
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Capabilities called for Instance = " & instance & " Capabilities are IO", LogType.LOG_TYPE_INFO)
            Return HomeSeerAPI.Enums.eCapabilities.CA_IO
        Else
            If (GetStringIniFile(UDN, DeviceInfoIndex.diSonosPlayerType.ToString, "").ToUpper <> "SUB") And GetBooleanIniFile("Options", "MediaAPIEnabled", False) Then
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Capabilities called for Instance = " & instance & " Capabilities are IO and Music", LogType.LOG_TYPE_INFO)
                Return HomeSeerAPI.Enums.eCapabilities.CA_IO + HomeSeerAPI.Enums.eCapabilities.CA_Music
            Else
                Return HomeSeerAPI.Enums.eCapabilities.CA_IO
            End If
        End If
    End Function

    Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
        ' return the access level for this plugin
        ' 1=everyone can access, no protection
        ' 2=level 2 plugin. Level 2 license required to run this plugin
        Return 1
    End Function

    Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
        Get
            Return False  'We do not use a COM port, or we will get it from the user and save it using code in this plug-in.
        End Get
    End Property

    Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
        Return True
    End Function

    Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements IPlugInAPI.SupportsMultipleInstancesSingleEXE
        Return True
    End Function

    Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InstanceFriendlyName called for instance = " & instance, LogType.LOG_TYPE_INFO)
        If instance <> "" And (Not isRoot) Then
            Return ZoneName 'GetZoneByUDN(instance)
        Else
            Return ""
        End If
    End Function

    Public Function InitIO(ByVal port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO

        InitIO = ""

        If TrialPhase Then
            Try
                Dim d As DateTime
                d = DateTime.Parse(TrialLastDate) 'date assignment
                Dim CurrentDate As Date = Now.Date
                If Date.Compare(CurrentDate, d) > 0 Then
                    InitIO = "Time expired, get latest version"
                    Exit Function
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO for Instance = " & instance & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        Try
            CurrentAppPath = Environment.CurrentDirectory
            ' if root directory, the currentAppPath will be ended by a | or / like in C:\, else no slash, so to make it consistent, I must remove this
            If CurrentAppPath <> "" Then
                If CurrentAppPath(CurrentAppPath.Length - 1) = "/" Or CurrentAppPath(CurrentAppPath.Length - 1) = "\" Then
                    CurrentAppPath.Remove(CurrentAppPath.Length - 1, 1)
                End If
            End If
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found CurrentAppPath = " & CurrentAppPath, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine the current directory path this plugin is running in.", LogType.LOG_TYPE_ERROR)
            CurrentAppPath = hs.GetAppPath
        End Try
        Try
            HSisRunningOnLinux = (hs.GetOSType() = eOSType.linux)
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found HS running on Linux = " & HSisRunningOnLinux.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine what OS HS is running on.", LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Log("InitIO Called for Instance = " & instance & " and running on OS = " & Environment.OSVersion.Platform.ToString, LogType.LOG_TYPE_INFO)
            ImRunningOnLinux = Type.GetType("Mono.Runtime") IsNot Nothing
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO for Instance = " & instance & " found this plugin running on Linux = " & ImRunningOnLinux.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in InitIO Called for Instance = " & instance & ". Unable to determine what OS this plugin is running on.", LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If isRoot Then DebugLogFileName = CurrentAppPath & "/html" & DebugLogFileName
        Catch ex As Exception
            DebugLogFileName = ""
        End Try
        Try ' added 9/7/2019 in v3.1.0.37 because this is causing an exception in HS4 and was already fixed in the MediaController PI
            PlugInIPAddress = hs.GetIPAddress
            PluginIPPort = hs.GetINISetting("Settings", "gWebSvrPort", "")
            'Dim Ethernetports As Dictionary(Of String, String) = GetEthernetPorts() ' test dcor
            Dim HSServerIPBinding = hs.GetINISetting("Settings", "gServerAddressBind", "")
            If HSServerIPBinding <> "" Then ' added in v3.1.0.30
                If HSServerIPBinding.ToLower <> "(no binding)" Then
                    ' HS has a non default setting
                    If HSServerIPBinding = PlugInIPAddress Then
                        ' all cool here
                    Else
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in InitIO for Instance = " & instance & " received (" & PlugInIPAddress & "), which is a different IP adress from it's server binding (" & HSServerIPBinding & ")", LogType.LOG_TYPE_WARNING)
                    End If
                End If
            End If
            If ServerIPAddress <> "" Then
                ImRunningLocal = CheckLocalIPv4Address(hs.GetIPAddress)
                If Not ImRunningLocal Then
                    PlugInIPAddress = GetLocalIPv4Address()
                End If
            End If
        Catch ex As Exception
            Log("Error in InitIO Called for Instance = " & instance & ". Unable to retrieve IP address info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitIO Called for Instance = " & instance, LogType.LOG_TYPE_INFO)

        If Not isRoot Then
            gIOEnabled = True
            gInterfaceStatus = ERR_NONE
            bShutDown = False
            InitMusicAPI()
            Exit Function
        End If

        If gIOEnabled Then Exit Function

        If HSisRunningOnLinux Then
            FileArtWorkPath = tIFACE_NAME & "/Artwork/"
        End If

        If ImRunningOnLinux Then
            MusicDBPath = "/html/" & tIFACE_NAME & "/MusicDB/SonosDB.sdb"
            RadioStationsDBPath = "/html/" & tIFACE_NAME & "/MusicDB/SonosRadioStationsDB.sdb"
            DockedPlayersDBPath = "/html/" & tIFACE_NAME & "/MusicDB/"
            FileArtWorkPath = tIFACE_NAME & "/Artwork/"
            AnnouncementPath = "/" & tIFACE_NAME & "/Announcements/"
            'DebugLogFileName = CurrentAppPath & "/html/" & tIFACE_NAME & "/Logs/SonosDebug.txt"
            Try
                If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME) Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME)
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & " directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/MusicDB") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/MusicDB")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/MusicDB directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/Logs") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/Logs")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/Logs directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/Announcements") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "/html/" & tIFACE_NAME & "/Announcements")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create /html/" & tIFACE_NAME & "/Announcements directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not File.Exists(CurrentAppPath & "/Config/" & tIFACE_NAME & ".ini") Then
                    Try
                        WriteIntegerIniFile("Options", "piDebuglevel", DebugLevel.dlErrorsOnly) ' fixed error here where "piDebuglevel > DebugLevel.dlEvents=False" was written in V52
                        WriteIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly)
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create /Config/" & tIFACE_NAME & ".ini file with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try ' this is new to support the different SQL file for Linux
                If File.Exists(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName) Then
                    Try
                        File.Delete(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & SQLFileName)
                    Catch ex As Exception
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO. Unable to delete " & CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & SQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        File.Move(CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName, CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & SQLFileName)
                    Catch ex As Exception
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO. Unable to rename " & CurrentAppPath & "/html/" & tIFACE_NAME & "/bin/" & LinuxSQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
        Else
            Try
                If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME) Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME)
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & " directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\MusicDB") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\MusicDB")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\MusicDB directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Logs") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Logs")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Logs directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not Directory.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\Announcements") Then
                    Try
                        Directory.CreateDirectory(CurrentAppPath & "\html\" & tIFACE_NAME & "\Announcements")
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create \html\" & tIFACE_NAME & "\Announcements directory: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try
                If Not File.Exists(CurrentAppPath & "\Config\" & tIFACE_NAME & ".ini") Then
                    Try
                        'WriteBooleanIniFile("Options", "Debug", False)
                        WriteIntegerIniFile("Options", "piDebuglevel", DebugLevel.dlErrorsOnly) ' fixed on 9/27/2020 in v3.1.0.54
                        WriteIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly) ' fixed on 9/27/2020 in v3.1.0.54
                    Catch ex As Exception
                        Log("Error in InitIO. Unable to create \Config\" & tIFACE_NAME & ".ini file with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
            Try ' this is new to support the different SQL file for Linux
                If File.Exists(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName) Then
                    Try
                        File.Delete(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName)
                    Catch ex As Exception
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO. Unable to delete " & CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        File.Move(CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName, CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & SQLFileName)
                    Catch ex As Exception
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in InitIO. Unable to rename " & CurrentAppPath & "\html\" & tIFACE_NAME & "\bin\" & WindowsSQLFileName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Catch ex As Exception
            End Try
        End If

        gLogToDisk = GetBooleanIniFile("Options", "LogToDisk", False)
        If gLogToDisk Then OpenLogFile(DebugLogFileName)

        ReadIniFile()

        Try
            ' create our jquery web page
            If MainInstance <> "" Then
                ConfigurationPage = New SonosConfig(ConfigPage & ":" & MainInstance)
            Else
                ConfigurationPage = New SonosConfig(ConfigPage)
            End If
            ConfigurationPage.RefToPlugIn = Me
            ' register the page with the HS web server, HS will post back to the WebPage class
            ' "pluginpage" is the URL to access this page
            ' comment this out if you are going to use the GenPage/PutPage API istead
            hs.RegisterPage(ConfigPage, sIFACE_NAME, MainInstance)

            ' register a configuration link that will appear on the interfaces page
            Dim wpd As New WebPageDesc With {
                .link = ConfigPage
            }
            If MainInstance <> "" Then
                wpd.linktext = "Config for instance " & MainInstance
            Else
                wpd.linktext = "Config"
            End If

            wpd.page_title = "Sonos Config"
            wpd.plugInName = sIFACE_NAME
            wpd.plugInInstance = MainInstance
            callback.RegisterConfigLink(wpd)

            ' register a normal page to appear in the HomeSeer menu
            wpd = New WebPageDesc With {
                .link = ConfigPage
            }
            If MainInstance <> "" Then
                wpd.linktext = "Config for instance " & MainInstance
            Else
                wpd.linktext = "Config"
            End If
            wpd.page_title = "Sonos Config"
            wpd.plugInName = sIFACE_NAME
            wpd.plugInInstance = MainInstance
            hs.RegisterLinkEx(wpd)

        Catch ex As Exception
            bShutDown = True
            Return "Error on InitIO: " & ex.Message
        End Try

        Try
            ' register a normal page to appear in the HomeSeer menu
            Dim Helpwpd As New WebPageDesc With {
                .link = sIFACE_NAME & "/Help/Help.htm",
                .linktext = "Sonos help",
                .page_title = "Sonos help",
                .plugInName = sIFACE_NAME,
                .plugInInstance = MainInstance
            }
            hs.RegisterHelpLink(Helpwpd)
        Catch ex As Exception
            Log("Error in InitIO intializing the help link with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If MainInstance <> "" Then
                UPnPViewerPage = New UPnPDebugWindow(UPnPViewPage & ":" & MainInstance)
            Else
                UPnPViewerPage = New UPnPDebugWindow(UPnPViewPage)
            End If
            'UPnPViewerPage.RefToPlugIn = Me
            ' register the page with the HS web server, HS will post back to the WebPage class
            ' "pluginpage" is the URL to access this page
            ' comment this out if you are going to use the GenPage/PutPage API istead
            hs.RegisterPage(UPnPViewPage, sIFACE_NAME, MainInstance)
        Catch ex As Exception
            Log("Error in InitIO intializing the UPnP Viewer link with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try


        Dim Index As Integer
        For Index = 0 To MaxTOActionArray
            MyTimeoutActionArray(Index) = 0
        Next
        MyTimeoutActionArray(TORediscover) = TORediscoverValue
        MyTimeoutActionArray(TOCheckChange) = TOCheckChangeValue

        MyControllerTimer = New Timers.Timer With {
            .Interval = 1000,
            .AutoReset = True,
            .Enabled = True
        }

        Try
            'MyPingTimer = New Timers.Timer
            'MyPingTimer.Interval = 10000
            'MyPingTimer.AutoReset = True
            'MyPingTimer.Enabled = True
        Catch ex As Exception
            Log("Error in InitIO. Unable to create the Ping timer with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MyAnnouncementTimer = New Timers.Timer With {
                .Interval = 500,
                .AutoReset = True,
                .Enabled = True
            }
        Catch ex As Exception
            Log("Error in InitIO. Unable to create the Announcement timer with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        InitDeviceFlag = True
        Log("Sonos Plugin Initialized", LogType.LOG_TYPE_INFO)
        gIOEnabled = True
        gInterfaceStatus = ERR_NONE
        bShutDown = False
        Return ""       ' return no error, or an error message

    End Function

    Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO
        ' shutdown the I/O interface
        ' called when HS exits
        Log("ShutdownIO called for Instance = " & instance & " and isRoot = " & isRoot.ToString, LogType.LOG_TYPE_INFO)

        If isRoot Then
            hs.SetDeviceValueByRef(MasterHSDeviceRef, msDisconnected, True)
            If MyControllerTimer IsNot Nothing Then MyControllerTimer.Enabled = False
            'If MyPingTimer IsNot Nothing Then MyPingTimer.Enabled = False
            If MyAnnouncementTimer IsNot Nothing Then MyAnnouncementTimer.Enabled = False
            MyLinkgroupArray = Nothing
            Try
                If ProxySpeakerActive Then
                    callback.UnRegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                    ProxySpeakerActive = False
                End If
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " unregistering SpeakerProxy with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If MySSDPDevice IsNot Nothing Then
                Try
                    MySSDPDevice.Dispose()
                    MySSDPDevice = Nothing
                Catch ex As Exception
                    Log("Error in ShutdownIO for Instance = " & instance & " destroying the SSDP device with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            If UPnPViewerPage IsNot Nothing Then
                Try
                    UPnPViewerPage.Dispose()
                Catch ex As Exception
                End Try
                UPnPViewerPage = Nothing
            End If
            Try
                RemoveHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
            Catch ex As Exception
            End Try
            Try
                RemoveHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
            Catch ex As Exception
            End Try
            Try
                DestroySonosControllers()
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " destroying controllers with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                If MyHSDeviceLinkedList IsNot Nothing Then
                    If MyHSDeviceLinkedList.Count > 0 Then
                        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                            If Not HSDevice Is Nothing Then
                                HSDevice.Close()
                            End If
                        Next
                        MyHSDeviceLinkedList.Clear()
                        NbrOfSonosPlayers = 0
                    End If
                End If
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " deleting weblinks with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            MyControllerTimer = Nothing
            'MyPingTimer = Nothing
            MyAnnouncementTimer = Nothing
            Try
                hs.UnRegisterHelpLinks(sIFACE_NAME, instance)
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " unregistering Help links with error ", LogType.LOG_TYPE_ERROR)
            End Try
            Try
                ' register a normal page to appear in the HomeSeer menu
                Dim wpd = New WebPageDesc With {
                    .link = ConfigPage
                }
                If MainInstance <> "" Then
                    wpd.linktext = "Config for instance " & MainInstance
                Else
                    wpd.linktext = "Config"
                End If
                wpd.page_title = "Sonos Config"
                wpd.plugInName = sIFACE_NAME
                wpd.plugInInstance = MainInstance
                hs.UnRegisterLinkEx(wpd)
            Catch ex As Exception
                Log("Error in ShutdownIO for Instance = " & instance & " unregistering links with error ", LogType.LOG_TYPE_ERROR)
            End Try
            CloseLogFile()
            gIOEnabled = False  ' moved here in v3.1.0.25 on 9/17/2018
            bShutDown = True    ' moved here in v3.1.0.25 on 9/17/2018
        Else
            Try
                Disconnect(True)
            Catch ex As Exception
            End Try
            Try
                DeleteWebLink(UDN, ZoneName)
            Catch ex As Exception
            End Try
            Try
                DestroyPlayer(True)
            Catch ex As Exception
            End Try
        End If


        'gIOEnabled = False     ' removed here in v3.1.0.25 on 9/17/2018
        GC.Collect()
        ' bShutDown = True      ' removed here in v3.1.0.25 on 9/17/2018 caused all instances to terminate as opposed to just a single instance

    End Sub

    Public Function ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String Implements HomeSeerAPI.IPlugInAPI.ConfigDevice
        Log("ConfigDevice called for instance = " & MainInstance & " and zoneplayer = " & ZoneName & " with Ref = " & ref.ToString & " and user = " & user & " and userRights = " & userRights.ToString & " and newDevice = " & newDevice.ToString, LogType.LOG_TYPE_INFO)
        ConfigDevice = ""
        Exit Function
    End Function

    Public Function ConfigDevicePost(ref As Integer, data As String, user As String, userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost
        Log("ConfigDevicePost called for instance = " & MainInstance & " and Ref = " & ref.ToString & " and data = " & data.ToString & " and user = " & user & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
        ConfigDevicePost = Enums.ConfigDevicePostReturn.CallbackOnce
        Exit Function
    End Function


    ' Web Page Generation - OLD METHODS
    ' ================================================================================================
    Public Function GenPage(ByVal link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GenPage called for instance = " & instance & " with link = " & link.ToString, LogType.LOG_TYPE_INFO)
        Return ""
    End Function
    Public Function PagePut(ByVal data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PagePut called for instance = " & instance & "  with data = " & data.ToString, LogType.LOG_TYPE_INFO)
        PagePut = ""
        Try
            'Return MyPlayerControlWebPage.postBackProc("", data, "", "")
        Catch ex As Exception
            Log("Error in hspi.PagePut called with data = " & data.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return ""
    End Function
    ' ================================================================================================

    ' Web Page Generation - NEW METHODS
    ' ================================================================================================
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin
        'If you have more than one web page, use pageName to route it to the proper GetPagePlugin
        GetPagePlugin = ""
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("hspi.GetPagePlugin called for instance = " & instance & " and pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)
        Try
            If pageName.IndexOf(ConfigPage) = 0 Then
                Return ConfigurationPage.GetPagePlugin(pageName, user, userRights, queryString)
            ElseIf pageName.IndexOf(PlayerControlPage) = 0 Then
                If instance = "" Then
                    ' this is a problem
                    Log("Error in hspi.GetPagePlugin missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString & ". Unable to retrieve UDN = " & instance, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerControlWebPage.GetPagePlugin(pageName, user, userRights, queryString, True)
            ElseIf pageName.IndexOf(UPnPViewPage) = 0 Then
                Return UPnPViewerPage.GetPagePlugin(pageName, user, userRights, queryString)
            Else
                Return ""
            End If
        Catch ex As Exception
            Log("Error in hspi.GetPagePlugin called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
        'If you have more than one web page, use pageName to route it to the proper postBackProc
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("hspi.PostBackProc called for instance = " & instance & " with pageName = " & pageName.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
        PostBackProc = ""
        Try
            If pageName.IndexOf(ConfigPage) = 0 Then
                Return ConfigurationPage.postBackProc(pageName, data, user, userRights)
            ElseIf pageName.IndexOf(PlayerControlPage) = 0 Then
                Dim ZoneUDN As String = ""
                Dim UDNParts As String() = pageName.Split(":")
                If UBound(UDNParts) > 0 Then
                    ZoneUDN = UDNParts(1) ' the structure is PlayerControl:RINCON_000E5859008Axxxxx
                Else
                    ' this is a problem
                    Log("Error in hspi.PostBackProc missing UDN part called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                If instance = "" Then
                    ' this is the root
                    Log("Error in hspi.PostBackProc called for instance = " & instance & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & ". Unable to retrieve UDN = " & ZoneUDN, LogType.LOG_TYPE_ERROR)
                    Return ""
                End If
                Return MyPlayerControlWebPage.postBackProc(pageName, data, user, userRights)
            ElseIf pageName.IndexOf(UPnPViewPage) = 0 Then
                Return UPnPViewerPage.postBackProc(pageName, data, user, userRights)
            Else
                Return ""
            End If
        Catch ex As Exception
            Log("Error in hspi.PostBackProc called for instance = " & instance & " with pageName = " & pageName.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function
    ' ================================================================================================

    Public Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent
        Log("HSEvent for instance = " & MainInstance & " : " & EventType.ToString, LogType.LOG_TYPE_INFO)
        Select Case EventType
            Case Enums.HSEvent.VALUE_CHANGE
        End Select
    End Sub



    Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus
        'Log("InterfaceStatus called for instance " & instance, LogType.LOG_TYPE_INFO)
        Dim es As New IPlugInAPI.strInterfaceStatus With {
            .intStatus = IPlugInAPI.enumInterfaceStatus.OK
        }

        If TrialPhase Then
            Try
                Dim d As DateTime
                d = DateTime.Parse(TrialLastDate) 'date assignment
                Dim CurrentDate As Date = Now.Date
                Dim DaysLeft As Integer = Date.Compare(d, CurrentDate)
                If DaysLeft < 0 Then
                    es.sStatus = "Time expired"
                    es.intStatus = IPlugInAPI.enumInterfaceStatus.FATAL
                Else
                    es.sStatus = "Expires: " & TrialLastDate
                    es.intStatus = IPlugInAPI.enumInterfaceStatus.WARNING
                End If
            Catch ex As Exception
                Log("Error in InterfaceStatus for Instance = " & MainInstance, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        Return es
    End Function

    Public Function PollDevice(ByVal dvref As Integer) As IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice
        'Console.WriteLine("PollDevice")
        Dim ri As IPlugInAPI.PollResultInfo
        ri.Result = IPlugInAPI.enumPollResult.OK
        Return ri
    End Function

    Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("RaisesGenericCallbacks called.", LogType.LOG_TYPE_INFO)
        Return True
    End Function

    Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOMulti called", LogType.LOG_TYPE_INFO)
        Dim CC As CAPIControl
        For Each CC In colSend
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOMulti set value: " & CC.ControlValue.ToString & "->ref:" & CC.Ref.ToString, LogType.LOG_TYPE_INFO)
            SetIOEx(CC)
        Next
    End Sub

    Public Sub SetIOEx(CC As CAPIControl)

        If Not CC Is Nothing Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetIOEx called for Ref = " & CC.Ref.ToString & ", Index " & CC.CCIndex.ToString & ", controlFlag = " & CC.ControlFlag.ToString &
                 ", ControlString" & CC.ControlString.ToString & ", ControlType = " & CC.ControlType.ToString & ", ControlValue = " & CC.ControlValue.ToString &
                  ", Label = " & CC.Label.ToString, LogType.LOG_TYPE_INFO)
        Else
            Exit Sub    ' Not ours.

        End If

        Dim SonosPlayer As HSPI
        Dim DestinationSonosPlayer As HSPI = Nothing
        Dim SourceSonosPlayer As HSPI = Nothing
        If CC.Ref = MasterHSDeviceRef Then
            ' Check which action is required
            Select Case CC.ControlValue
                Case msPauseAll ' All Zones Pause
                    AllZonesPause()
                Case msPlayAll ' All Zones Play
                    AllZonesOn()
                Case msMuteAll ' All Zones Mute On"
                    AllZonesMuteOn()
                Case msUnmuteAll ' All Zones Mute Off"
                    AllZonesMuteOff()
                Case msBuildDB ' BuildDB
                    ' start a very small timer, so the PI can return from this SetIOMulti call and create the DB in the timer thread
                    Try
                        MyDBCreationTimer = New Timers.Timer With {
                            .Interval = 500,
                            .AutoReset = False,
                            .Enabled = True
                        }
                    Catch ex As Exception
                        Log("Error in SetIOEx. Unable to create the CreateDBtimer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Case Else
                    ' these are the Link/Unlink buttons. The buttons start with index 6 and are pairs. If Even it is an On button else an Odd button
                    Dim LinkgroupIndex As Integer
                    LinkgroupIndex = CC.ControlValue - LinkGroupButtonOffset
                    Dim TempIndex As Integer
                    Dim LinkgroupZones() As String = Nothing
                    Dim LinkgroupZone As String
                    Dim ButtonName As String = ""
                    LinkgroupZone = GetStringIniFile("LinkgroupNames", "Names", "")
                    LinkgroupZones = Split(LinkgroupZone, "|")
                    If LinkgroupZones Is Nothing Then
                        Log("Error in SetIOEx. No Linkgroups found in the .ini file", LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End If
                    TempIndex = LinkgroupIndex
                    ' if the index is odd, the command is to unlink, if even it is to link.
                    TempIndex = TempIndex \ 2 ' this will create the right offset in the LinkgroupZones
                    Dim IsEven As Boolean = False
                    If (LinkgroupIndex Mod 2 = 0) Then IsEven = True
                    If TempIndex > UBound(LinkgroupZones) Then
                        Log("Error in SetIOEx. Button index larger then defined link buttons in .ini file. Received index = " & TempIndex.ToString & " and max index = " & UBound(LinkgroupZones).ToString, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End If
                    LinkgroupZone = LinkgroupZones(TempIndex)
                    If IsEven Then ButtonName = "Link-" & LinkgroupZone Else ButtonName = "Unlink-" & LinkgroupZone
                    If LinkgroupZone <> "" Then HandleLinkEvents(ButtonName) Else Log("Error in SetIOEx. Couldn't find Button", LogType.LOG_TYPE_ERROR)
                    LinkgroupZone = Nothing
                    LinkgroupZones = Nothing
            End Select
            Exit Sub
        End If
        ' [UPnP HSRef to UDN]
        Dim ccUDN As String = GetStringIniFile("UPnP HSRef to UDN", CC.Ref, "")
        If ccUDN = "" Then
            ' not found
            Log("ERROR in SetIOEx: Zoneplayer not found for received event. Event = " & CC.ControlValue.ToString & " DeviceRef = " & CC.Ref.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        SonosPlayer = GetAPIByUDN(ccUDN)
        If SonosPlayer Is Nothing Then
            ' not found
            Log("ERROR in SetIOEx: Zoneplayer not found for received event. Event = " & CC.ControlValue.ToString & " DeviceRef = " & CC.Ref.ToString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        SonosPlayer.TreatSetIOEx(CC)
        SonosPlayer = Nothing
    End Sub



    Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
        Return True
    End Function

    Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
        Return False
    End Function

    Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
        Return False
    End Function


#End Region

#Region "Actions Interface"

    Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionCount called", LogType.LOG_TYPE_INFO)
        If Not isRoot Then Return 0
        Return 1
    End Function

    Public ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionName called with ActionNumber = " & ActionNumber.ToString, LogType.LOG_TYPE_INFO)
            Select Case ActionNumber
                Case 1
                    If MainInstance <> "" Then
                        Return "Sonos Instance " & MainInstance & " Actions"
                    Else
                        Return "Sonos Actions"
                    End If
            End Select
            Return ""
        End Get
    End Property

    Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
        Set(ByVal value As Boolean)
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionAdvancedMode Set called with Value = " & value.ToString, LogType.LOG_TYPE_INFO)
            mvarActionAdvanced = value
        End Set
        Get
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionAdvancedMode Get called and returned = " & mvarActionAdvanced, LogType.LOG_TYPE_INFO)
            Return mvarActionAdvanced
        End Get
    End Property

    Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
        Dim stb As New StringBuilder
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim PlayerList As New clsJQuery.jqDropList("PlayerListAction" & sUnique, ActionsPageName, True)
        Dim CommandList As New clsJQuery.jqDropList("CommandListAction" & sUnique, ActionsPageName, True)
        Dim sKey As String
        Dim action As New action

        Dim PlayerIndex As String = "" ' this is the selected UDN??
        Dim LinkList As String = ""
        Dim CommandIndex As String = ""
        Dim InputIndex As String = ""
        Dim MuteIndex As String = ""
        Dim ShuffleIndex As String = ""
        Dim RepeatIndex As String = ""
        Dim LoudnessIndex As String = ""
        Dim SelectionIndex As String = ""
        Dim GenreSelection As String = ""
        Dim ArtistSelection As String = ""
        Dim AlbumSelection As String = ""
        Dim TrackSelection As String = ""
        Dim FavoriteSelection As String = ""
        Dim AddInfoSelection As String = ""
        Dim ClearQueueSelection As String = ""
        Dim PlayNowSelection As String = ""
        Dim AudioInputPlayer As String = ""
        Dim InputString As String = ""


        Try
            CommandList.autoPostBack = True
            CommandList.AddItem("--Please Select--", "", False)

            PlayerList.autoPostBack = True
            PlayerList.AddItem("--Please Select--", "", False)


            If Not (ActInfo.DataIn Is Nothing) Then
                DeSerializeObject(ActInfo.DataIn, action)
            Else 'new event, so clean out the trigger object
                action = New action
            End If

            For Each sKey In action.Keys
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found skey = " & sKey.ToString & " and PlayerUDN = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerIndex = action(sKey)
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI found PlayerIndex with Actioninfo = " & PlayerIndex.ToString, LogType.LOG_TYPE_INFO)
                    Case InStr(sKey, "LinkPlayerAction") > 0
                        LinkList = action(sKey)
                    Case InStr(sKey, "AudioInputPlayerAction") > 0
                        AudioInputPlayer = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        CommandIndex = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputIndex = action(sKey)
                    Case InStr(sKey, "MuteAction") > 0
                        MuteIndex = action(sKey)
                    Case InStr(sKey, "ShuffleAction") > 0
                        ShuffleIndex = action(sKey)
                    Case InStr(sKey, "RepeatAction") > 0
                        RepeatIndex = action(sKey)
                    Case InStr(sKey, "LoudnessAction") > 0
                        LoudnessIndex = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ArtistAction") > 0
                        ArtistSelection = action(sKey)
                    Case InStr(sKey, "AlbumAction") > 0
                        AlbumSelection = action(sKey)
                    Case InStr(sKey, "TrackAction") > 0
                        TrackSelection = action(sKey)
                    Case InStr(sKey, "FavoriteAction") > 0
                        FavoriteSelection = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        ClearQueueSelection = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI found Command = " & CommandIndex & " and PlayerUDN = " & PlayerIndex & " and Text = " & InputIndex & " and Linklist = " & LinkList & " and Mute = " & MuteIndex & " and Shuffle = " & ShuffleIndex & " and Repeat = " & RepeatIndex & " and Loudness = " & LoudnessIndex, LogType.LOG_TYPE_INFO)

            Dim InputBox As New clsJQuery.jqTextBox("InputBoxAction" & sUnique, "text", InputIndex, ActionsPageName, 40, True)
            CommandList.AddItem("Play Track", "Play Track", CommandIndex = "Play Track")
            CommandList.AddItem("Play Artist", "Play Artist", CommandIndex = "Play Artist")
            CommandList.AddItem("Play Album", "Play Album", CommandIndex = "Play Album")
            CommandList.AddItem("Play Playlist", "Play Playlist", CommandIndex = "Play Playlist")
            CommandList.AddItem("Play Radiostation", "Play Radiostation", CommandIndex = "Play Radiostation")
            CommandList.AddItem("Play Favorite", "Play Favorite", CommandIndex = "Play Favorite")
            'CommandList.AddItem("Play Audiobook", "Play Audiobook", CommandIndex = "Play Audiobook")
            'CommandList.AddItem("Play Podcast", "Play Podcast", CommandIndex = "Play Podcast")
            CommandList.AddItem("Play AudioInput", "Play AudioInput", CommandIndex = "Play AudioInput")
            CommandList.AddItem("Play TV", "Play TV", CommandIndex = "Play TV")
            CommandList.AddItem("Play URL", "Play URL", CommandIndex = "Play URL")
            CommandList.AddItem("Pause if Playing", "Pause if Playing", CommandIndex = "Pause if Playing")
            CommandList.AddItem("Resume if Paused", "Resume if Paused", CommandIndex = "Resume if Paused")
            CommandList.AddItem("Stop Playing", "Stop Playing", CommandIndex = "Stop Playing")
            CommandList.AddItem("Next Track", "Next Track", CommandIndex = "Next Track")
            CommandList.AddItem("Previous Track", "Previous Track", CommandIndex = "Previous Track")
            CommandList.AddItem("Next RadioStation", "Next RadioStation", CommandIndex = "Next RadioStation")
            CommandList.AddItem("Previous RadioStation", "Previous RadioStation", CommandIndex = "Previous RadioStation")
            CommandList.AddItem("Next Playlist", "Next Playlist", CommandIndex = "Next Playlist")
            CommandList.AddItem("Previous Playlist", "Previous Playlist", CommandIndex = "Previous Playlist")
            CommandList.AddItem("Link", "Link", CommandIndex = "Link")
            CommandList.AddItem("Unlink", "Unlink", CommandIndex = "Unlink")
            CommandList.AddItem("AddGroup", "AddGroup", CommandIndex = "AddGroup")
            CommandList.AddItem("Clear Queue", "Clear Queue", CommandIndex = "Clear Queue")
            CommandList.AddItem("Save State All Players", "Save State All Players", CommandIndex = "Save State All Players")
            CommandList.AddItem("Restore State All Players", "Restore State All Players", CommandIndex = "Restore State All Players")
            CommandList.AddItem("Set Volume", "Set Volume", CommandIndex = "Set Volume")
            CommandList.AddItem("Set Mute", "Set Mute", CommandIndex = "Set Mute")
            CommandList.AddItem("Set Loudness", "Set Loudness", CommandIndex = "Set Loudness")
            CommandList.AddItem("Set Repeat", "Set Repeat", CommandIndex = "Set Repeat")
            CommandList.AddItem("Set Shuffle", "Set Shuffle", CommandIndex = "Set Shuffle")
            CommandList.AddItem("Set Track Position", "Set Track Position", CommandIndex = "Set Track Position")


            stb.Append("Select Command:")
            stb.Append(CommandList.Build)

            Select Case CommandIndex
                Case "Save State All Players", "Restore State All Players"
                    ' no Additional parameters
                    Return stb.ToString
                Case "" ' start building, this is the first time called
                    Return stb.ToString
            End Select


            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    ' special case, just find the first player with a reference
                    Dim Playername As String = HSDevice.ZonePlayerControllerRef.ZoneName
                    Dim PlayerUDN As String = HSDevice.ZonePlayerControllerRef.UDN
                    If HSDevice.ZonePlayerControllerRef.ZoneModel.ToUpper <> "SUB" Then
                        If CommandIndex = "Play TV" Then
                            If CheckPlayerCanPlayTV(HSDevice.ZonePlayerControllerRef.ZoneModel) Then  ' changed 7/12/2019. Moved check into ONE place
                                PlayerList.AddItem(Playername, PlayerUDN, PlayerIndex = PlayerUDN)
                            End If
                        Else
                            PlayerList.AddItem(Playername, PlayerUDN, PlayerIndex = PlayerUDN)
                        End If
                    End If
                End If
            Next

            stb.Append("Select Player:")
            stb.Append(PlayerList.Build)

            Select Case CommandIndex
                Case "Play Track"
                    Dim ArtistList As New clsJQuery.jqDropList("ArtistAction" & sUnique, ActionsPageName, True)
                    Dim AlbumList As New clsJQuery.jqDropList("AlbumAction" & sUnique, ActionsPageName, True)
                    Dim TrackList As New clsJQuery.jqDropList("TrackAction" & sUnique, ActionsPageName, True)
                    stb.Append("</br>")
                    If CommandIndex = "Play Track" Then
                        Dim ClearQueueSelectionList As New clsJQuery.jqDropList("ClearQueueAction" & sUnique, ActionsPageName, True)
                        ClearQueueSelectionList.AddItem("No", "No", "No" = ClearQueueSelection)
                        ClearQueueSelectionList.AddItem("Yes", "Yes", "Yes" = ClearQueueSelection)
                        Dim PlayNowSelectionList As New clsJQuery.jqDropList("PlayNowAction" & sUnique, ActionsPageName, True)
                        PlayNowSelectionList.AddItem("Last", "Last", "Last" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Now", "Now", "Now" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Next", "Next", "Next" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("No", "No", "No" = PlayNowSelection)
                        stb.Append("Clear Queue:" & ClearQueueSelectionList.Build)
                        stb.Append("Play Last/Now/Next/No:" & PlayNowSelectionList.Build)
                        stb.Append("</br>")
                    End If
                    stb.Append("Select Artist:")
                    If PlayerIndex = "" Then
                        ArtistList.AddItem("Select Player First!", "", True)
                        stb.Append(ArtistList.Build)
                        Return stb.ToString
                    End If
                    Dim MusicApi As HSPI = Nothing
                    Try
                        MusicApi = MyHSPIControllerRef.GetAPIByUDN(PlayerIndex)
                    Catch ex As Exception
                    End Try
                    If MusicApi Is Nothing Then
                        ArtistList.AddItem("Player not Found!", "", True)
                        stb.Append(ArtistList.Build)
                        Return stb.ToString
                    End If
                    ArtistList.AddItem("--Please Select--", "", False)
                    Dim ar() As String
                    ar = MusicApi.GetArtists("", "")
                    If ar IsNot Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI for Player = " & ZoneName & "  has " & (UBound(ar) + 1).ToString & " entries in the ArtistList", LogType.LOG_TYPE_INFO)
                        For Each MediaItem In ar
                            If MediaItem <> "" Then
                                ArtistList.AddItem(MediaItem, MediaItem, MediaItem = ArtistSelection)
                            End If
                        Next
                    End If
                    stb.Append(ArtistList.Build)
                    If ArtistSelection = "" Then Return stb.ToString

                    stb.Append("</br>")
                    stb.Append("Select Album:")
                    AlbumList.AddItem("--Please Select--", "", False)
                    ar = MusicApi.GetAlbums(ArtistSelection, "")
                    If ar IsNot Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI for Player = " & ZoneName & "  has " & (UBound(ar) + 1).ToString & " entries in the AlbumList", LogType.LOG_TYPE_INFO)
                        For Each MediaItem In ar
                            If MediaItem <> "" Then
                                AlbumList.AddItem(MediaItem, MediaItem, MediaItem = AlbumSelection)
                            End If
                        Next
                    End If
                    stb.Append(AlbumList.Build)
                    If AlbumSelection = "" Then Return stb.ToString

                    stb.Append("</br>")
                    stb.Append("Select Track:")
                    TrackList.AddItem("--Please Select--", "", False)
                    ar = MusicApi.DBGetTracks(ArtistSelection, AlbumSelection, "")
                    If ar IsNot Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI for Player = " & ZoneName & "  has " & (UBound(ar) + 1).ToString & " entries in the TrackList", LogType.LOG_TYPE_INFO)
                        For Each MediaItem In ar
                            If MediaItem <> "" Then
                                'TrackList.AddItem(EncodeTags(MediaItem), EncodeTags(MediaItem), MediaItem = TrackSelection)
                                ' System.Web.HttpUtility.HtmlDecode
                                TrackList.AddItem(MediaItem, MediaItem, MediaItem = TrackSelection)
                            End If
                        Next
                    End If
                    stb.Append(TrackList.Build)

                    Return stb.ToString


                Case "Play Album", "Play Artist", "Play Playlist", "Play Radiostation", "Play Audiobook", "Play Podcast"
                    Dim AdditionalList As New clsJQuery.jqDropList("AddInfoAction" & sUnique, ActionsPageName, True)
                    stb.Append("</br>")
                    If CommandIndex = "Play Artist" Or CommandIndex = "Play Album" Or CommandIndex = "Play Playlist" Then
                        Dim ClearQueueSelectionList As New clsJQuery.jqDropList("ClearQueueAction" & sUnique, ActionsPageName, True)
                        ClearQueueSelectionList.AddItem("No", "No", "No" = ClearQueueSelection)
                        ClearQueueSelectionList.AddItem("Yes", "Yes", "Yes" = ClearQueueSelection)
                        Dim PlayNowSelectionList As New clsJQuery.jqDropList("PlayNowAction" & sUnique, ActionsPageName, True)
                        PlayNowSelectionList.AddItem("Last", "Last", "Last" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Now", "Now", "Now" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("Next", "Next", "Next" = PlayNowSelection)
                        PlayNowSelectionList.AddItem("No", "No", "No" = PlayNowSelection)
                        stb.Append("Clear Queue:" & ClearQueueSelectionList.Build)
                        stb.Append("Play Last/Now/Next/No:" & PlayNowSelectionList.Build)
                        stb.Append("</br>")
                    End If
                    Select Case CommandIndex
                        Case "Play Artist"
                            stb.Append("Select Artist:")
                        Case "Play Album"
                            stb.Append("Select Album:")
                        Case "Play Playlist"
                            stb.Append("Select Playlist:")
                        Case "Play Radiostation"
                            stb.Append("Select Radiostation:")
                        Case "Play Audiobook"
                            stb.Append("Select Audiobook:")
                        Case "Play Podcast"
                            stb.Append("Select Podcast:")
                    End Select
                    If PlayerIndex = "" Then
                        AdditionalList.AddItem("Select Player First!", "", True)
                        stb.Append(AdditionalList.Build)
                        Return stb.ToString
                    End If
                    Dim MusicApi As HSPI = Nothing
                    Try
                        MusicApi = MyHSPIControllerRef.GetAPIByUDN(PlayerIndex)
                    Catch ex As Exception
                    End Try
                    If MusicApi Is Nothing Then
                        AdditionalList.AddItem("Player not Found!", "", True)
                        stb.Append(AdditionalList.Build)
                        Return stb.ToString
                    End If
                    AdditionalList.AddItem("--Please Select--", "", False)
                    Dim ar() As String = Nothing
                    Select Case CommandIndex
                        Case "Play Artist"
                            ar = MusicApi.GetArtists("", "")
                        Case "Play Album"
                            ar = MusicApi.GetAlbums("", "")
                        Case "Play Playlist"
                            ar = MusicApi.GetPlaylists("", False)
                        Case "Play Radiostation"
                            ar = MusicApi.LibGetRadioStationlists(True)
                        Case "Play Audiobook"
                            ar = MusicApi.LibGetAudiobookslists("")
                        Case "Play Podcast"
                            ar = MusicApi.LibGetPodcastlists("")
                    End Select
                    If ar IsNot Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI for Player = " & ZoneName & "  has " & (UBound(ar) + 1).ToString & " entries in the AdditionalList", LogType.LOG_TYPE_INFO)
                        For Each MediaItem In ar
                            If MediaItem <> "" Then
                                AdditionalList.AddItem(MediaItem, MediaItem, MediaItem = AddInfoSelection)
                            End If
                        Next
                    End If
                    stb.Append(AdditionalList.Build)
                Case "Play Favorite"
                    Dim AdditionalList As New clsJQuery.jqDropList("AddInfoAction" & sUnique, ActionsPageName, True)
                    stb.Append("</br>")
                    stb.Append("Select Favorite:")
                    If PlayerIndex = "" Then
                        AdditionalList.AddItem("Select Player First!", "", True)
                        stb.Append(AdditionalList.Build)
                        Return stb.ToString
                    End If
                    Dim MusicApi As HSPI = Nothing
                    Try
                        MusicApi = MyHSPIControllerRef.GetAPIByUDN(PlayerIndex)
                    Catch ex As Exception
                    End Try
                    If MusicApi Is Nothing Then
                        AdditionalList.AddItem("Player not Found!", "", True)
                        stb.Append(AdditionalList.Build)
                        Return stb.ToString
                    End If
                    AdditionalList.AddItem("--Please Select--", "", False)
                    Dim ar() As String = Nothing
                    ar = MusicApi.LibGetObjectslist("FV:2")
                    If ar IsNot Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionBuildUI for Player = " & ZoneName & "  has " & (UBound(ar) + 1).ToString & " entries in the AdditionalList", LogType.LOG_TYPE_INFO)
                        For Each MediaItem In ar
                            If MediaItem <> "" Then
                                AdditionalList.AddItem(MediaItem, MediaItem, MediaItem = AddInfoSelection)
                            End If
                        Next
                    End If
                    stb.Append(AdditionalList.Build)
                Case "Play AudioInput"
                    Dim AudioInputPlayerList As New clsJQuery.jqDropList("AudioInputPlayerAction" & sUnique, ActionsPageName, True)
                    AudioInputPlayerList.AddItem("--Please Select--", "", False)
                    For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                        If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                            ' special case, just find the first player with a reference
                            Dim Playername As String = HSDevice.ZonePlayerControllerRef.ZoneName
                            Dim PlayerUDN As String = HSDevice.ZonePlayerControllerRef.UDN
                            If CheckPlayerHasAudioInput(HSDevice.ZonePlayerControllerRef.ZoneModel.ToUpper) Then ' Changed on 7/12/2019 in v3.1.0.31 
                                AudioInputPlayerList.AddItem(Playername, PlayerUDN, AudioInputPlayer = PlayerUDN)
                            End If
                        End If
                    Next
                    stb.Append("Select Player to Link to Audio Source:")
                    stb.Append(AudioInputPlayerList.Build)
                    Return stb.ToString
                Case "Play TV"
                    Return stb.ToString
                Case "Pause if Playing", "Resume if Paused", "Stop Playing", "Next Track", "Previous Track", "Next RadioStation", "Previous RadioStation", "Next Playlist", "Previous Playlist", "Clear Queue"
                    Return stb.ToString
                Case "Link"
                    Dim LinkPlayerList As New clsJQuery.jqSelector("LinkPlayerAction" & sUnique, ActionsPageName, True)
                    LinkPlayerList.IncludeValues = True ' added 2/17/2020 in v.52 so we can have the value (UDN) returned. Name only gives issues with international characters
                    For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                        If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                            ' special case, just find the first player with a reference
                            Dim Playername As String = HSDevice.ZonePlayerControllerRef.ZoneName
                            Dim PlayerUDN As String = HSDevice.ZonePlayerControllerRef.UDN
                            If HSDevice.ZonePlayerControllerRef.ZoneModel.ToUpper <> "SUB" Then
                                LinkPlayerList.AddItem(Playername, PlayerUDN, InStr(LinkList, PlayerUDN) > 0)
                            End If
                        End If
                    Next
                    stb.Append("Select Player to Link:")
                    stb.Append(LinkPlayerList.Build)
                    Return stb.ToString
                Case "Unlink"
                    Return stb.ToString
                Case "AddGroup"
                    Dim LinkPlayerList As New clsJQuery.jqSelector("LinkPlayerAction" & sUnique, ActionsPageName, True)
                    LinkPlayerList.IncludeValues = True ' added 2/17/2020 in v.52 so we can have the value (UDN) returned. Name only gives issues with international characters
                    For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                        If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                            ' special case, just find the first player with a reference
                            Dim Playername As String = HSDevice.ZonePlayerControllerRef.ZoneName
                            Dim PlayerUDN As String = HSDevice.ZonePlayerControllerRef.UDN
                            If HSDevice.ZonePlayerControllerRef.ZoneModel.ToUpper <> "SUB" Then
                                LinkPlayerList.AddItem(Playername, PlayerUDN, InStr(LinkList, PlayerUDN) > 0)
                            End If
                        End If
                    Next
                    stb.Append("Select Player to join grouping:")
                    stb.Append(LinkPlayerList.Build)
                    Return stb.ToString
                Case "Set Volume"
                    InputBox.size = 4
                    stb.Append("Set Volume:")
                    stb.Append(InputBox.Build)
                    Return stb.ToString
                Case "Set Mute"
                    Dim MuteList As New clsJQuery.jqDropList("MuteAction" & sUnique, ActionsPageName, True) With {
                        .autoPostBack = True
                    }
                    MuteList.AddItem("--Please Select--", "", False)
                    MuteList.AddItem("Mute Off", "Mute Off", MuteIndex = "Mute Off")
                    MuteList.AddItem("Mute On", "Mute On", MuteIndex = "Mute On")
                    stb.Append("Set Mute:")
                    stb.Append(MuteList.Build)
                    Return stb.ToString
                Case "Set Loudness"
                    Dim LoudnessList As New clsJQuery.jqDropList("LoudnessAction" & sUnique, ActionsPageName, True) With {
                        .autoPostBack = True
                    }
                    LoudnessList.AddItem("--Please Select--", "", False)
                    LoudnessList.AddItem("Loudness Off", "Loudness Off", LoudnessIndex = "Loudness Off")
                    LoudnessList.AddItem("Loudness On", "Loudness On", LoudnessIndex = "Loudness On")
                    stb.Append("Set Loudness:")
                    stb.Append(LoudnessList.Build)
                    Return stb.ToString
                Case "Set Repeat"
                    Dim RepeatList As New clsJQuery.jqDropList("RepeatAction" & sUnique, ActionsPageName, True) With {
                        .autoPostBack = True
                    }
                    RepeatList.AddItem("--Please Select--", "", False)
                    RepeatList.AddItem("Repeat Off", "Repeat Off", RepeatIndex = "Repeat Off")
                    RepeatList.AddItem("Repeat On", "Repeat On", RepeatIndex = "Repeat On")
                    stb.Append("Set Repeat:")
                    stb.Append(RepeatList.Build)
                    Return stb.ToString
                Case "Set Shuffle"
                    Dim ShuffleList As New clsJQuery.jqDropList("ShuffleAction" & sUnique, ActionsPageName, True) With {
                        .autoPostBack = True
                    }
                    ShuffleList.AddItem("--Please Select--", "", False)
                    ShuffleList.AddItem("Shuffle Off", "Shuffle Off", ShuffleIndex = "Shuffle Off")
                    ShuffleList.AddItem("Shuffle On", "Shuffle On", ShuffleIndex = "Shuffle On")
                    stb.Append("Set Shuffle:")
                    stb.Append(ShuffleList.Build)
                    Return stb.ToString
                Case "Set Track Position"
                    stb.Append("Set Track Position:")
                    stb.Append(InputBox.Build)
                    Return stb.ToString
                Case "Play URL"
                    Dim InputEditBox As New clsJQuery.jqTextBox("InputEditAction" & sUnique, "text", "", ActionsPageName, 1000, True)
                    stb.Append("Enter URL:")
                    stb.Append(InputEditBox.Build)
                    Return stb.ToString
            End Select

        Catch ex As Exception
            Log("Error in ActionBuildUI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return stb.ToString
    End Function

    Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim Configured As Boolean = False
        Dim sKey As String
        Dim itemsConfigured As Integer = 0
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return False
        End If
        Dim action As New action
        Dim PlayerUDN As String = ""
        Dim LinkList As String = ""
        Dim Command As String = ""
        Dim InputBox As String = ""
        Dim ShuffleIndex As String = ""
        Dim RepeatIndex As String = ""
        Dim LoudnessIndex As String = ""
        Dim MuteIndex As String = ""
        Dim SelectionIndex As String = ""
        Dim ArtistSelection As String = ""
        Dim AlbumSelection As String = ""
        Dim TrackSelection As String = ""
        Dim FavoriteSelection As String = ""
        Dim AddInfoSelection As String = ""
        Dim ClearQueueSelection As String = ""
        Dim PlayNowSelection As String = ""
        Dim AudioInputPlayer As String = ""
        Dim InputString As String = ""

        Try
            DeSerializeObject(ActInfo.DataIn, action)
            For Each sKey In action.Keys
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured found sKey = " & sKey.ToString & " and Value = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        PlayerUDN = action(sKey)
                    Case InStr(sKey, "LinkPlayerAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        LinkList = action(sKey)
                    Case InStr(sKey, "AudioInputPlayerAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        AudioInputPlayer = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0 AndAlso action(sKey) <> ""
                        itemsConfigured += 1
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0 AndAlso action(sKey) <> ""
                        Select Case Command
                            Case "Set Volume", "Set Track Position"
                                itemsConfigured += 1 ' only copy for correct command                                
                        End Select
                        InputBox = action(sKey)
                    Case InStr(sKey, "MuteAction") > 0 AndAlso action(sKey) <> ""
                        If Command = "Set Mute" Then itemsConfigured += 1
                        MuteIndex = action(sKey)
                    Case InStr(sKey, "ShuffleAction") > 0 AndAlso action(sKey) <> ""
                        If Command = "Set Shuffle" Then itemsConfigured += 1
                        ShuffleIndex = action(sKey)
                    Case InStr(sKey, "RepeatAction") > 0 AndAlso action(sKey) <> ""
                        If Command = "Set Repeat" Then itemsConfigured += 1
                        RepeatIndex = action(sKey)
                    Case InStr(sKey, "LoudnessAction") > 0 AndAlso action(sKey) <> ""
                        If Command = "Set Loudness" Then itemsConfigured += 1
                        LoudnessIndex = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track") Then itemsConfigured += 1
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ArtistAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track") Then itemsConfigured += 1
                        ArtistSelection = action(sKey)
                    Case InStr(sKey, "AlbumAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track") Then itemsConfigured += 1
                        AlbumSelection = action(sKey)
                    Case InStr(sKey, "TrackAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track") Then itemsConfigured += 1
                        TrackSelection = action(sKey)
                    Case InStr(sKey, "FavoriteAction") > 0 AndAlso action(sKey) <> ""
                        If Command = "Play Favorite" Then itemsConfigured += 1
                        FavoriteSelection = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Album" Or Command = "Play Artist" Or Command = "Play Playlist" Or Command = "Play Radiostation" Or Command = "Play Audiobook" Or Command = "Play Podcast" Or Command = "Play Favorite") Then itemsConfigured += 1
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track" Or Command = "Play Album" Or Command = "Play Artist" Or Command = "Play Playlist") Then itemsConfigured += 1
                        ClearQueueSelection = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0 AndAlso action(sKey) <> ""
                        If (Command = "Play Track" Or Command = "Play Album" Or Command = "Play Artist" Or Command = "Play Playlist") Then itemsConfigured += 1
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0 AndAlso action(sKey) <> ""
                        Select Case Command
                            Case "Play URL"
                                itemsConfigured += 1 ' only copy for correct command                                
                        End Select
                        InputString = action(sKey)
                End Select
            Next
            Select Case Command
                Case "Play Track"
                    If itemsConfigured = 7 Then Configured = True
                Case "Play Album", "Play Artist", "Play Playlist"
                    If itemsConfigured = 5 Then Configured = True
                Case "Play Radiostation", "Play Audiobook", "Play Podcast", "Play Favorite"
                    If itemsConfigured = 3 Then Configured = True
                Case "Pause if Playing", "Resume if Paused", "Stop Playing", "Next Track", "Previous Track", "Next RadioStation", "Previous RadioStation", "Next Playlist", "Previous Playlist", "Clear Queue", "Unlink", "Play TV"
                    If itemsConfigured = 2 Then Configured = True
                Case "Link", "AddGroup"
                    If itemsConfigured = 3 Then Configured = True
                Case "Play AudioInput"
                    If itemsConfigured = 3 Then Configured = True
                Case "Save State All Players", "Restore State All Players"
                    ' no Additional parameters
                    If itemsConfigured = 1 Then Configured = True
                Case "Set Volume"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Mute"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Loudness"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Repeat"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Shuffle"
                    If itemsConfigured = 3 Then Configured = True
                Case "Set Track Position"
                    If itemsConfigured <> 3 Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns False", LogType.LOG_TYPE_INFO)
                        Return False
                    End If
                    If Val(InputBox) <> 0 Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns True", LogType.LOG_TYPE_INFO)
                        Return True ' valid integer
                    End If
                    Dim Index As Integer
                    Dim Counter As Integer = 0
                    For Index = 0 To InputBox.Count - 1
                        If InputBox(Index) = ":" Then Counter += 1
                    Next
                    If Counter <> 2 Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns False", LogType.LOG_TYPE_INFO)
                        Return False
                    Else
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns True", LogType.LOG_TYPE_INFO)
                        Return True
                    End If
                Case "Play URL"
                    If itemsConfigured = 3 Then Configured = True
            End Select

        Catch ex As Exception
            Log("Error in ActionConfigured with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionConfigured returns " & Configured.ToString, LogType.LOG_TYPE_INFO)
        Return Configured
    End Function

    Public Function ActionReferencesDevice(ByVal ActInfo As IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionReferencesDevice called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString & " and dvRef = " & dvRef.ToString, LogType.LOG_TYPE_INFO)
        '
        ' Actions in the sample plug-in do not reference devices, but for demonstration purposes we will pretend they do, 
        '   and that ALL actions reference our sample devices.
        '
        'If ActInfo.DataIn Is Nothing Then
        ' no info, can't be good
        'Return False
        'End If
        ' look in ini file for [UPnP HSRef to UDN]
        'Dim action As New action
        'Try

        'If Not (ActInfo.DataIn Is Nothing) Then
        'DeSerializeObject(ActInfo.DataIn, action)
        'Else 'new event, so clean out the trigger object
        'action = New action
        'End If
        ' I suspect that I need to compare the dvRef and check whether it has anything to do with my devices. Then I could ??? check the data-in and lift which player (s) are involved and see if the dvref belongs to any of the players referenced
        'For Each sKey In action.Keys
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionReferencesDevice found skey = " & sKey.ToString & " and action = " & action(sKey), LogType.LOG_TYPE_INFO)
        'Next


        'Catch ex As Exception
        'Log("Error in ActionReferencesDevice with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        'End Try
        'If dvRef = -1 Then Return True
        Return False
    End Function

    Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI
        Dim stb As New StringBuilder
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return ""
        End If
        Dim PlayerUDN As String = ""
        Dim action As New action
        Try

            If Not (ActInfo.DataIn Is Nothing) Then
                DeSerializeObject(ActInfo.DataIn, action)
            Else 'new event, so clean out the trigger object
                action = New action
            End If

            Dim LinkList As String = ""
            Dim PlayerName As String = ""
            Dim Command As String = ""
            Dim InputBox As String = ""
            Dim ShuffleIndex As String = ""
            Dim RepeatIndex As String = ""
            Dim LoudnessIndex As String = ""
            Dim MuteIndex As String = ""
            Dim SelectionIndex As String = ""
            Dim ArtistSelection As String = ""
            Dim AlbumSelection As String = ""
            Dim TrackSelection As String = ""
            Dim FavoriteSelection As String = ""
            Dim AddInfoSelection As String = ""
            Dim ClearQueueSelection As String = ""
            Dim PlayNowSelection As String = ""
            Dim AudioInputPlayerUDN As String = ""
            Dim AudioInputPlayerName As String = ""
            Dim InputString As String = ""


            For Each sKey In action.Keys
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found skey = " & sKey.ToString & " and PlayerUDN = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerUDN = action(sKey)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found PlayerIndex with Actioninfo = " & PlayerUDN.ToString, LogType.LOG_TYPE_INFO)
                        If PlayerUDN <> "" Then
                            PlayerName = GetZoneByUDN(PlayerUDN)
                        End If
                    Case InStr(sKey, "LinkPlayerAction") > 0
                        LinkList = action(sKey)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionFormatUI found Linklist = " & LinkList.ToString, LogType.LOG_TYPE_INFO)
                    Case InStr(sKey, "AudioInputPlayerAction") > 0
                        AudioInputPlayerUDN = action(sKey)
                        If AudioInputPlayerUDN <> "" Then
                            AudioInputPlayerName = GetZoneByUDN(AudioInputPlayerUDN)
                        End If
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputBox = action(sKey)
                    Case InStr(sKey, "MuteAction") > 0
                        MuteIndex = action(sKey)
                    Case InStr(sKey, "ShuffleAction") > 0
                        ShuffleIndex = action(sKey)
                    Case InStr(sKey, "RepeatAction") > 0
                        RepeatIndex = action(sKey)
                    Case InStr(sKey, "LoudnessAction") > 0
                        LoudnessIndex = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ArtistAction") > 0
                        ArtistSelection = action(sKey)
                    Case InStr(sKey, "AlbumAction") > 0
                        AlbumSelection = action(sKey)
                    Case InStr(sKey, "TrackAction") > 0
                        TrackSelection = action(sKey)
                    Case InStr(sKey, "FavoriteAction") > 0
                        FavoriteSelection = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        ClearQueueSelection = action(sKey)
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next

            Dim CommandPrefix As String = ""
            If MainInstance <> "" Then
                CommandPrefix = "Sonos Instance " & MainInstance
            Else
                CommandPrefix = "Sonos"
            End If
            Select Case Command

                Case "Play Track"
                    stb.Append(CommandPrefix & " Action Play Track - " & " for player - " & PlayerName & ". Artist: " & ArtistSelection & ", Album: " & AlbumSelection & ", Track: " & TrackSelection & ", Clear Queue: " & ClearQueueSelection & ", Play: " & PlayNowSelection)
                Case "Play Artist"
                    stb.Append(CommandPrefix & " Action Play Artist - " & " for player - " & PlayerName & ". Artist: " & AddInfoSelection & ", Clear Queue: " & ClearQueueSelection & ", Play: " & PlayNowSelection)
                Case "Play Album"
                    stb.Append(CommandPrefix & " Action Play Album - " & " for player - " & PlayerName & ". Album: " & AddInfoSelection & ", Clear Queue: " & ClearQueueSelection & ", Play: " & PlayNowSelection)
                Case "Play Playlist"
                    stb.Append(CommandPrefix & " Action Play Playlist - " & " for player - " & PlayerName & ". Playlist: " & AddInfoSelection & ", Clear Queue: " & ClearQueueSelection & ", Play: " & PlayNowSelection)
                Case "Play Radiostation"
                    stb.Append(CommandPrefix & " Action Play Radiostation - " & " for player - " & PlayerName & ". Radiostation: " & AddInfoSelection)
                Case "Play Audiobook"
                    stb.Append(CommandPrefix & " Action Play Audiobook - " & " for player - " & PlayerName & ". Audiobook: " & AddInfoSelection)
                Case "Play Podcast"
                    stb.Append(CommandPrefix & " Action Play Podcast - " & " for player - " & PlayerName & ". Podcast: " & AddInfoSelection)
                Case "Play Favorite"
                    stb.Append(CommandPrefix & " Action Play Favorite - " & " for player - " & PlayerName & ". Favorite: " & AddInfoSelection)
                Case "Play AudioInput"
                    stb.Append(CommandPrefix & " Action Play Audio Input - " & " for player - " & PlayerName & " with Audio Input from Player: " & AudioInputPlayerName)
                Case "Play TV"
                    stb.Append(CommandPrefix & " Action Play TV for player - " & PlayerName)
                Case "Pause if Playing"
                    stb.Append(CommandPrefix & " Action Pause if Playing for player - " & PlayerName)
                Case "Resume if Paused"
                    stb.Append(CommandPrefix & " Action Resume if Paused for player - " & PlayerName)
                Case "Stop Playing"
                    stb.Append(CommandPrefix & " Action Stop Playing for player - " & PlayerName)
                Case "Next Track"
                    stb.Append(CommandPrefix & " Action Next Track for player - " & PlayerName)
                Case "Previous Track"
                    stb.Append(CommandPrefix & " Action Previous Track for player - " & PlayerName)
                Case "Next RadioStation"
                    stb.Append(CommandPrefix & " Action Next Radiostation for player - " & PlayerName)
                Case "Previous RadioStation"
                    stb.Append(CommandPrefix & " Action Previous Radiostation for player - " & PlayerName)
                Case "Next Playlist"
                    stb.Append(CommandPrefix & " Action Next Playlist for player - " & PlayerName)
                Case "Previous Playlist"
                    stb.Append(CommandPrefix & " Action Previous Playlist for player - " & PlayerName)
                Case "Link"
                    Dim Players As String()
                    Players = Split(LinkList, ",")
                    Dim PlayerNames As String = ""
                    For Each player In Players
                        If PlayerNames <> "" Then PlayerNames = PlayerNames & ","
                        PlayerNames = PlayerNames & GetZoneByUDN(player)
                    Next
                    stb.Append(CommandPrefix & " Action Link Master player - " & PlayerName & " to - " & PlayerNames)
                Case "AddGroup"
                    Dim Players As String()
                    Players = Split(LinkList, ",")
                    Dim PlayerNames As String = ""
                    For Each player In Players
                        If PlayerNames <> "" Then PlayerNames = PlayerNames & ","
                        PlayerNames = PlayerNames & GetZoneByUDN(player)
                    Next
                    stb.Append(CommandPrefix & " Action Group Master player - " & PlayerName & " to - " & PlayerNames)
                Case "Unlink"
                    stb.Append(CommandPrefix & " Action Unlink for player - " & PlayerName)
                Case "Clear Queue"
                    stb.Append(CommandPrefix & " Action ClearQueue for player - " & PlayerName)
                Case "Save State All Players"
                    stb.Append(CommandPrefix & " Action Save State all Players")
                Case "Restore State All Players"
                    stb.Append(CommandPrefix & " Action Restore State all Players")
                Case "Set Volume"
                    stb.Append(CommandPrefix & " Action Set Volume for player - " & PlayerName & " to - " & InputBox.ToString)
                Case "Set Mute"
                    stb.Append(CommandPrefix & " Action Set Mute for player - " & PlayerName & " - " & MuteIndex.ToString)
                Case "Set Loudness"
                    stb.Append(CommandPrefix & " Action Set Loudness for player - " & PlayerName & " - " & LoudnessIndex.ToString)
                Case "Set Repeat"
                    stb.Append(CommandPrefix & " Action Set Repeat for player - " & PlayerName & " - " & RepeatIndex.ToString)
                Case "Set Shuffle"
                    stb.Append(CommandPrefix & " Action Set Shuffle for player - " & PlayerName & " - " & ShuffleIndex.ToString)
                Case "Set Track Position"
                    stb.Append(CommandPrefix & " Action Set Track Position for player - " & PlayerName & " - " & InputBox.ToString)
                Case "Play URL"
                    stb.Append(CommandPrefix & " Action Play URL for player - " & PlayerName & " - " & InputString.ToString)
            End Select

        Catch ex As Exception
            Log("Error in ActionFormatUI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return stb.ToString
    End Function

    Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection,
                                        ByVal ActInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn _
                                        Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI

        Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI called", LogType.LOG_TYPE_INFO)

        Ret.sResult = ""
        ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        Ret.DataOut = ActInfoIN.DataIn
        Ret.TrigActInfo = ActInfoIN

        If PostData Is Nothing Then Return Ret
        If PostData.Count < 1 Then Return Ret
        Dim Action As New action
        If Not (ActInfoIN.DataIn Is Nothing) Then
            'DeSerializeObject(ActInfoIN.DataIn, Action)
        End If

        Dim parts As Collections.Specialized.NameValueCollection

        Dim sKey As String
        Dim Command As String = ""
        'Ret.DataOut = Nothing

        parts = PostData
        Try
            For Each sKey In parts.Keys
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI found sKey " & sKey.ToString, LogType.LOG_TYPE_INFO)
                If sKey Is Nothing Then Continue For
                If String.IsNullOrEmpty(sKey.Trim) Then Continue For
                Select Case True

                    Case InStr(sKey, "PlayerListAction") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "LinkPlayerAction") > 0
                        ' need to convert player name to UDN before storing
                        Dim PlayerNameString As String() = Split(parts(sKey), ",")
                        Dim PlayerUDNString As String = ""
                        For Each Playername In PlayerNameString
                            ' changed on 2/17/2020 in v.52 we now set a flag to return the value, so we now have the UDN. This was an issue with the international characters set
                            Dim playerParts As String() = Split(Playername, "|")
                            ' the first part is the name and the second part is the UDN
                            If PlayerUDNString <> "" Then PlayerUDNString = PlayerUDNString & ","
                            If UBound(playerParts) > 0 Then
                                ' the UDN is there
                                PlayerUDNString &= playerParts(1)
                            Else
                                ' the old way
                                PlayerUDNString = PlayerUDNString & GetUDNbyZoneName(Playername)
                            End If
                        Next
                        Action.Add(CObj(PlayerUDNString), sKey)
                    Case InStr(sKey, "AudioInputPlayerAction") > 0
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = parts(sKey)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI found Command " & Command.ToString, LogType.LOG_TYPE_INFO)
                        Action.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        Select Case Command
                            Case "Set Volume", "Set Track Position"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "MuteAction") > 0
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ActionProcessPostUI found MuteAction and command =  " & Command.ToString, LogType.LOG_TYPE_INFO)
                        Select Case Command
                            Case "Set Mute"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "ShuffleAction") > 0
                        Select Case Command
                            Case "Set Shuffle"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "RepeatAction") > 0
                        Select Case Command
                            Case "Set Repeat"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "LoudnessAction") > 0
                        Select Case Command
                            Case "Set Loudness"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "SelectionAction") > 0
                        'Action.Add(CObj(parts(sKey)), sKey) ' dcor I keep this out for time being so it never gets stored!
                    Case InStr(sKey, "AddInfoAction") > 0
                        Select Case Command
                            Case "Play Artist", "Play Album", "Play Playlist", "Play Radiostation", "Play Audiobook", "Play Podcast", "Play Favorite"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "ClearQueueAction") > 0
                        Select Case Command
                            Case "Play Track", "Play Artist", "Play Album", "Play Playlist"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "PlayNowAction") > 0
                        Select Case Command
                            Case "Play Track", "Play Artist", "Play Album", "Play Playlist"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case InStr(sKey, "InputEditAction") > 0
                        Select Case Command
                            Case "Play URL"
                                Action.Add(CObj(parts(sKey)), sKey) ' only copy for correct command                                
                        End Select
                    Case Else
                        Action.Add(CObj(parts(sKey)), sKey)
                End Select
            Next
            If Not SerializeObject(Action, Ret.DataOut) Then
                Ret.sResult = sIFACE_NAME & " Error, Serialization failed. Signal Action not added."
                Return Ret
            End If
        Catch ex As Exception
            Ret.sResult = "ERROR, Exception in Action UI of " & sIFACE_NAME & ": " & ex.Message
            Return Ret
        End Try

        ' All OK
        Ret.sResult = ""
        Return Ret



    End Function

    Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
        HandleAction = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleAction called with evRef = " & ActInfo.evRef.ToString & " and SubTANumber = " & ActInfo.SubTANumber.ToString & " and TANumber = " & ActInfo.TANumber.ToString & " and UID = " & ActInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If ActInfo.DataIn Is Nothing Then
            ' no info, can't be good
            Return False
        End If
        Dim action As New action
        Try
            'Select Case ActInfo.TANumber
            'Case 14     ' Save State all Players,
            'SaveAllPlayersState()
            'Return True
            'Case 15     ' Restore state
            'RestoreAllPlayersState()
            'Return True
            'End Select

            DeSerializeObject(ActInfo.DataIn, action)
            Dim sKey As String
            Dim PlayerUDN As String = ""
            Dim LinkList As String = ""
            Dim Command As String = ""
            Dim InputBox As String = ""
            Dim ShuffleIndex As String = ""
            Dim RepeatIndex As String = ""
            Dim LoudnessIndex As String = ""
            Dim MuteIndex As String = ""
            Dim SelectionIndex As String = ""
            Dim ArtistSelection As String = ""
            Dim AlbumSelection As String = ""
            Dim TrackSelection As String = ""
            Dim FavoriteSelection As String = ""
            Dim AddInfoSelection As String = ""
            Dim ClearQueueSelection As Boolean = False
            Dim PlayNowSelection As String = ""
            Dim QueueAction As QueueActions = QueueActions.qaDontPlay
            Dim AudioInputPlayerUDN As String = ""
            Dim InputString As String = ""

            For Each sKey In action.Keys
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleAction found sKey = " & sKey.ToString & " and Value = " & action(sKey), LogType.LOG_TYPE_INFO)
                Select Case True
                    Case InStr(sKey, "PlayerListAction") > 0
                        PlayerUDN = action(sKey)
                    Case InStr(sKey, "LinkPlayerAction") > 0
                        LinkList = action(sKey)
                    Case InStr(sKey, "AudioInputPlayerAction") > 0
                        AudioInputPlayerUDN = action(sKey)
                    Case InStr(sKey, "CommandListAction") > 0
                        Command = action(sKey)
                    Case InStr(sKey, "InputBoxAction") > 0
                        InputBox = action(sKey)
                    Case InStr(sKey, "MuteAction") > 0
                        MuteIndex = action(sKey)
                    Case InStr(sKey, "ShuffleAction") > 0
                        ShuffleIndex = action(sKey)
                    Case InStr(sKey, "RepeatAction") > 0
                        RepeatIndex = action(sKey)
                    Case InStr(sKey, "LoudnessAction") > 0
                        LoudnessIndex = action(sKey)
                    Case InStr(sKey, "SelectionAction") > 0
                        SelectionIndex = action(sKey)
                    Case InStr(sKey, "ArtistAction") > 0
                        ArtistSelection = action(sKey)
                    Case InStr(sKey, "AlbumAction") > 0
                        AlbumSelection = action(sKey)
                    Case InStr(sKey, "TrackAction") > 0
                        TrackSelection = action(sKey)
                    Case InStr(sKey, "FavoriteAction") > 0
                        FavoriteSelection = action(sKey)
                    Case InStr(sKey, "AddInfoAction") > 0
                        AddInfoSelection = action(sKey)
                    Case InStr(sKey, "ClearQueueAction") > 0
                        If action(sKey) = "Yes" Then
                            ClearQueueSelection = True
                        End If
                    Case InStr(sKey, "PlayNowAction") > 0
                        PlayNowSelection = action(sKey)
                        Select Case PlayNowSelection
                            Case "Now"
                                QueueAction = QueueActions.qaPlayNow
                            Case "Last"
                                QueueAction = QueueActions.qaPlayLast
                            Case "Next"
                                QueueAction = QueueActions.qaPlayNext
                            Case "No"
                                QueueAction = QueueActions.qaDontPlay
                        End Select
                    Case InStr(sKey, "InputEditAction") > 0
                        InputString = action(sKey)
                End Select
            Next

            Select Case Command ' added here in version 21 because a group command has no MusicAPI and the retrieval just below for the MusicAPI would result in "nothing"
                Case "Save State All Players"
                    SaveAllPlayersState()
                    Return True
                Case "Restore State All Players"
                    RestoreAllPlayersState()
                    Return True
            End Select

            Dim MusicApi As HSPI = Nothing
            Try
                MusicApi = MyHSPIControllerRef.GetAPIByUDN(PlayerUDN)
            Catch ex As Exception
            End Try
            If MusicApi Is Nothing Then Return False

            Select Case Command
                Case "Play Track"
                    MusicApi.PlayMusic(ArtistSelection, AlbumSelection, "", "", TrackSelection, "", "", "", "", "", ClearQueueSelection, QueueAction)
                    Return True
                Case "Play Artist"
                    MusicApi.PlayMusic(AddInfoSelection, "", "", "", "", "", "", "", "", "", ClearQueueSelection, QueueAction)
                Case "Play Album"
                    MusicApi.PlayMusic("", AddInfoSelection, "", "", "", "", "", "", "", "", ClearQueueSelection, QueueAction)
                    Return True
                Case "Play Playlist"
                    MusicApi.PlayMusic("", "", AddInfoSelection, "", "", "", "", "", "", "", ClearQueueSelection, QueueAction)
                    Return True
                Case "Play Radiostation"
                    MusicApi.PlayMusic("", "", AddInfoSelection, "", "", "", "", "", "", "", False, QueueActions.qaPlayNow)
                    Return True
                Case "Play Audiobook"
                    MusicApi.PlayMusic("", "", "", "", "", "", "", "", AddInfoSelection, "", False, QueueActions.qaPlayNow)
                    Return True
                Case "Play Podcast"
                    MusicApi.PlayMusic("", "", "", "", "", "", "", "", "", AddInfoSelection, False, QueueActions.qaPlayNow)
                    Return True
                Case "Play Favorite"
                    MusicApi.PlayFavorite(AddInfoSelection, True, QueueActions.qaPlayNow)
                    Return True
                Case "Play AudioInput"
                    MusicApi.PlayLineInput(AudioInputPlayerUDN)
                    Return True
                Case "Play TV"
                    MusicApi.PlayTV()
                    Return True
                Case "Pause if Playing"
                    MusicApi.PauseIfPlaying()
                    Return True
                Case "Resume if Paused"
                    MusicApi.PlayIfPaused()
                    Return True
                Case "Stop Playing"
                    MusicApi.StopPlay()
                    Return True
                Case "Next Track"
                    MusicApi.TrackNext()
                    Return True
                Case "Previous Track"
                    MusicApi.TrackPrev()
                    Return True
                Case "Next RadioStation"
                    MusicApi.RadioStationNext()
                    Return True
                Case "Previous RadioStation"
                    MusicApi.RadioStationPrev()
                    Return True
                Case "Next Playlist"
                    MusicApi.PlaylistNext()
                    Return True
                Case "Previous Playlist"
                    MusicApi.PlaylistPrev()
                    Return True
                Case "Link"
                    Dim Players As String()
                    Players = Split(LinkList, ",")
                    Dim PlayerNames As String = ""
                    Dim PlayerAPI As HSPI
                    If Players.Length > 0 Then
                        For Each player In Players
                            PlayerAPI = MyHSPIControllerRef.GetAPIByUDN(player)
                            If Not PlayerAPI Is Nothing Then
                                PlayerAPI.Link(PlayerUDN, False)
                            End If
                        Next
                    End If
                Case "AddGroup"
                    Dim Players As String()
                    Players = Split(LinkList, ",")
                    Dim PlayerNames As String = ""
                    Dim PlayerAPI As HSPI
                    If Players.Length > 0 Then
                        For Each player In Players
                            PlayerAPI = MyHSPIControllerRef.GetAPIByUDN(player)
                            If Not PlayerAPI Is Nothing Then
                                PlayerAPI.Link(PlayerUDN, True)
                            End If
                        Next
                    End If
                Case "Unlink"
                    MusicApi.Unlink()
                    Return True
                Case "Clear Queue"
                    MusicApi.ClearQueue()
                    Return True
                Case "Save State All Players"
                    SaveAllPlayersState()
                    Return True
                Case "Restore State All Players"
                    RestoreAllPlayersState()
                    Return True
                Case "Set Volume"
                    If Len(InputBox.Trim) > 0 Then
                        Dim Vol As Integer
                        Vol = Val(InputBox.Trim)
                        Try
                            MusicApi.PlayerVolume = Vol
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                Case "Set Mute"
                    Select Case MuteIndex
                        Case "Mute On"
                            MusicApi.SonosMute()
                            Return True
                        Case "Mute Off"
                            MusicApi.UnMute()
                            Return True
                    End Select
                    Return False
                Case "Set Loudness"
                    Select Case LoudnessIndex
                        Case "Loudness Off"
                            MusicApi.SetLoudnessState("Master", False)
                            Return True
                        Case "Loudness On"
                            MusicApi.SetLoudnessState("Master", True)
                            Return True
                    End Select
                    Return False
                Case "Set Repeat"
                    Select Case RepeatIndex
                        Case "Repeat Off"
                            MusicApi.SonosRepeat = Repeat_modes.repeat_off
                            Return True
                        Case "Repeat On"
                            MusicApi.SonosRepeat = Repeat_modes.repeat_all
                            Return True
                    End Select
                    Return False
                Case "Set Shuffle"
                    Select Case ShuffleIndex
                        Case "Shuffle On"
                            MusicApi.SonosShuffle = 1
                            Return True
                        Case "Shuffle Off"
                            MusicApi.SonosShuffle = 2
                            Return True
                    End Select
                    Return False
                Case "Set Track Position"
                    If Len(InputBox) > 0 Then
                        Try
                            If InputBox.IndexOf(":") <> -1 Then
                                MusicApi.SeekTime(InputBox) ' this must be in the format hh:mm:ss or Sonos will error. Should already be in the right format
                            Else
                                MusicApi.SeekTime(ConvertSecondsToTimeFormat(InputBox)) ' this must be in the format hh:mm:ss or Sonos will error
                            End If
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                    Return True
                Case "Play URL"
                    MusicApi.StopPlay()
                    MusicApi.PlayURI(InputString, "", False)
                    MusicApi.SonosPlay()
                    Return True
            End Select
        Catch ex As Exception

        End Try
    End Function

#End Region

#Region "Conditions Properties"

    Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
        ' <summary>
        ' Indicates (when True) that the Trigger is in Condition mode - it is for triggers that can also operate as a condition
        '    or for allowing Conditions to appear when a condition is being added to an event.
        ' </summary>
        ' <param name="TrigInfo">The event, group, and trigger info for this particular instance.</param>
        ' <value></value>
        ' <returns>The current state of the Condition flag.</returns>
        ' <remarks></remarks>
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Condition.get called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
            Return ConditionSetFlag ' added in v3.1.25 to prevent that conditions are used as triggers 
        End Get
        Set(ByVal value As Boolean)
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Condition.set called for instance " & instance & " with Value = " & value.ToString & " and evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
            ConditionSetFlag = value ' added in v3.1.25 to prevent that conditions are used as triggers
        End Set
    End Property

    Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HasConditions.get called for instance " & instance & " with TriggerNumber = " & TriggerNumber.ToString, LogType.LOG_TYPE_INFO)
            Select Case TriggerNumber
                Case 1
                    Return False
                Case 2
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

#End Region

#Region "Trigger Interface"


    Public ReadOnly Property HasTriggers() As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HasTriggers called", LogType.LOG_TYPE_INFO)
            Return True
        End Get
    End Property

    Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerCount called for instance " & instance & " " & " and ConditionFlag = " & ConditionSetFlag.ToString, LogType.LOG_TYPE_INFO)
            If Not isRoot Then
                Return 0
            Else
                Return 2
            End If
        End Get
    End Property

    Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
        Get
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerName called for instance " & instance & " with TriggerNumber = " & TriggerNumber & " and ConditionFlag = " & ConditionSetFlag.ToString, LogType.LOG_TYPE_INFO)
            If MainInstance <> "" Then
                Select Case TriggerNumber
                    Case 1
                        Return "Sonos Instance " & MainInstance & " Trigger"
                    Case 2
                        Return "Sonos Instance " & MainInstance & " Condition"
                    Case Else
                        Return ""
                End Select
            Else
                Select Case TriggerNumber
                    Case 1
                        Return "Sonos Trigger"
                    Case 2
                        Return "Sonos Condition"
                    Case Else
                        Return ""
                End Select
            End If
        End Get
    End Property

    Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SubTriggerCount called for instance " & instance & " with TriggerNumber = " & TriggerNumber, LogType.LOG_TYPE_INFO)
            Return 0
        End Get
    End Property

    Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
        Get
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SubTriggerName called for instance " & instance & " with TriggerNumber = " & TriggerNumber, LogType.LOG_TYPE_INFO)
            Return ""
        End Get
    End Property

    Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
        Dim stb As New StringBuilder
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI called for instance " & instance & " with sUnique = " & sUnique.ToString & " and evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim PlayerList As New clsJQuery.jqDropList("PlayerListTrigger" & sUnique, TriggersPageName, True)
        Dim CommandList As New clsJQuery.jqDropList("CommandListTrigger" & sUnique, TriggersPageName, True)
        Dim LinkgroupList As New clsJQuery.jqDropList("LinkgroupListTrigger" & sUnique, TriggersPageName, True)
        Dim trigger As New trigger
        Dim sKey As String

        CommandList.autoPostBack = True
        CommandList.AddItem("--Please Select--", "", False)

        PlayerList.autoPostBack = True
        PlayerList.AddItem("--Please Select--", "", False)

        LinkgroupList.autoPostBack = True
        LinkgroupList.AddItem("--Please Select--", "", False)

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        Else 'new event, so clean out the trigger object
            trigger = New trigger
        End If

        Dim PlayerIndex As String = "" ' this is the selected UDN??
        Dim CommandIndex As String = ""
        Dim LinkgroupIndex As String = ""
        Dim InputIndex As String = ""
        For Each sKey In trigger.Keys
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found skey = " & sKey.ToString & " and PlayerUDN = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerIndex = trigger(sKey)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found PlayerIndex with triggerinfo = " & PlayerIndex.ToString, LogType.LOG_TYPE_INFO)
                Case InStr(sKey, "CommandListTrigger") > 0
                    CommandIndex = trigger(sKey)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found CommandIndex with triggerinfo = " & CommandIndex.ToString, LogType.LOG_TYPE_INFO)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputIndex = trigger(sKey)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found InputIndex with triggerinfo = " & InputIndex.ToString, LogType.LOG_TYPE_INFO)
                Case InStr(sKey, "LinkgroupListTrigger") > 0
                    LinkgroupIndex = trigger(sKey)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerBuildUI found LinkgroupIndex with triggerinfo = " & LinkgroupIndex.ToString, LogType.LOG_TYPE_INFO)
            End Select
        Next
        Dim InputBox As New clsJQuery.jqTextBox("InputBoxTrigger" & sUnique, "text", InputIndex, TriggersPageName, 40, True)
        Select Case TrigInfo.TANumber
            Case 1 ' Sonos Triggers
                CommandList.AddItem("Sonos Track Change", "Sonos Track Change", CommandIndex = "Sonos Track Change")
                CommandList.AddItem("Sonos Player Stop", "Sonos Player Stop", CommandIndex = "Sonos Player Stop")
                CommandList.AddItem("Sonos Player Paused", "Sonos Player Paused", CommandIndex = "Sonos Player Paused")
                CommandList.AddItem("Sonos Player Start Playing", "Sonos Player Start Playing", CommandIndex = "Sonos Player Start Playing")
                CommandList.AddItem("Sonos Volume Up", "Sonos Volume Up", CommandIndex = "Sonos Volume Up")
                CommandList.AddItem("Sonos Volume Down", "Sonos Volume Down", CommandIndex = "Sonos Volume Down")
                CommandList.AddItem("Sonos Player Docked", "Sonos Player Docked", CommandIndex = "Sonos Player Docked")
                CommandList.AddItem("Sonos Player Undocked", "Sonos Player Undocked", CommandIndex = "Sonos Player Undocked")
                CommandList.AddItem("Sonos Player Line-in Connected", "Sonos Player Line-in Connected", CommandIndex = "Sonos Player Line-in Connected")
                CommandList.AddItem("Sonos Player Line-in Disconnected", "Sonos Player Line-in Disconnected", CommandIndex = "Sonos Player Line-in Disconnected")
                CommandList.AddItem("Sonos Player Alarm Start", "Sonos Player Alarm Start", CommandIndex = "Sonos Player Alarm Start")
                CommandList.AddItem("Sonos Player Config Change", "Sonos Player Config Change", CommandIndex = "Sonos Player Config Change")
                CommandList.AddItem("Sonos Player Online", "Sonos Player Online", CommandIndex = "Sonos Player Online")
                CommandList.AddItem("Sonos Player Offline", "Sonos Player Offline", CommandIndex = "Sonos Player Offline")
                CommandList.AddItem("Sonos Next Track Change", "Sonos Next Track Change", CommandIndex = "Sonos Next Track Change")
                CommandList.AddItem("Sonos Announcement Start", "Sonos Announcement Start", CommandIndex = "Sonos Announcement Start")
                CommandList.AddItem("Sonos Announcement Stop", "Sonos Announcement Stop", CommandIndex = "Sonos Announcement Stop")
            Case 2 ' Sonos Conditions
                CommandList.AddItem("IsPlaying", "IsPlaying", CommandIndex = "IsPlaying")
                CommandList.AddItem("IsPaused", "IsPaused", CommandIndex = "IsPaused")
                CommandList.AddItem("IsStopped", "IsStopped", CommandIndex = "IsStopped")
                CommandList.AddItem("IsNotPlaying", "IsNotPlaying", CommandIndex = "IsNotPlaying")
                CommandList.AddItem("IsNotPaused", "IsNotPaused", CommandIndex = "IsNotPaused")
                CommandList.AddItem("IsNotStopped", "IsNotStopped", CommandIndex = "IsNotStopped")
                CommandList.AddItem("hasTrack", "hasTrack", CommandIndex = "hasTrack")
                CommandList.AddItem("hasAlbum", "hasAlbum", CommandIndex = "hasAlbum")
                CommandList.AddItem("hasArtist", "hasArtist", CommandIndex = "hasArtist")
                CommandList.AddItem("IsMutted", "IsMutted", CommandIndex = "IsMutted")
                CommandList.AddItem("IsNotMutted", "IsNotMutted", CommandIndex = "IsNotMutted")
                CommandList.AddItem("isOnline", "isOnline", CommandIndex = "isOnline")
                CommandList.AddItem("isOffline", "isOffline", CommandIndex = "isOffline")
                CommandList.AddItem("isPlayingAnnouncement", "isPlayingAnnouncement", CommandIndex = "isPlayingAnnouncement")
                CommandList.AddItem("isNotPlayingAnnouncement", "isNotPlayingAnnouncement", CommandIndex = "isNotPlayingAnnouncement")  ' added 7/12/2019 in v3.1.0.31
        End Select

        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                ' special case, just find the first player with a reference
                Dim Playername As String = HSDevice.ZonePlayerControllerRef.ZoneName
                Dim PlayerUDN As String = HSDevice.ZonePlayerControllerRef.UDN
                If HSDevice.ZonePlayerControllerRef.ZoneModel.ToUpper <> "SUB" Then
                    PlayerList.AddItem(Playername, PlayerUDN, PlayerIndex = PlayerUDN)
                End If
            End If
        Next

        Dim LinkGroupString As String = GetStringIniFile("LinkgroupNames", "Names", "")
        If LinkGroupString <> "" Then
            Dim Names() As String
            Names = Split(LinkGroupString, "|")
            If Names IsNot Nothing Then
                For Each Name As String In Names
                    LinkgroupList.AddItem(Name, Name, LinkgroupIndex = Name)
                Next
            End If
        End If



        stb.Append("Select Command:")
        stb.Append(CommandList.Build)

        If CommandIndex = "Sonos Announcement Start" Or CommandIndex = "Sonos Announcement Stop" Or CommandIndex = "isPlayingAnnouncement" Or CommandIndex = "isNotPlayingAnnouncement" Then   ' added 7/12/2019 in v3.1.0.31
            stb.Append("Select Linkgroup:")
            stb.Append(LinkgroupList.Build)
        ElseIf CommandIndex <> "" Then
            stb.Append("Select Player:")
            stb.Append(PlayerList.Build)
        End If


        Select Case CommandIndex
            Case "hasTrack"
                stb.Append("Specify Track:")
                stb.Append(InputBox.Build)
            Case "hasAlbum"
                stb.Append("Specify Album:")
                stb.Append(InputBox.Build)
            Case "hasArtist"
                stb.Append("Specify Artist:")
                stb.Append(InputBox.Build)
        End Select


        Return stb.ToString

    End Function

    Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
        Get
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
            Dim Configured As Boolean = False
            Dim sKey As String
            Dim itemsConfigured As Integer = 0
            Dim itemsToConfigure As Integer = 2
            Dim trigger As New trigger
            If Not (TrigInfo.DataIn Is Nothing) Then
                DeSerializeObject(TrigInfo.DataIn, trigger)
                For Each sKey In trigger.Keys
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
                    Select Case True
                        Case InStr(sKey, "PlayerListTrigger") > 0 AndAlso trigger(sKey) <> ""
                            itemsConfigured += 1
                        Case InStr(sKey, "CommandListTrigger") > 0 AndAlso trigger(sKey) <> ""
                            Select Case trigger(sKey)
                                Case "hasTrack", "hasAlbum", "hasArtist"
                                    itemsToConfigure = 3
                            End Select
                            itemsConfigured += 1
                        Case InStr(sKey, "InputBoxTrigger") > 0 AndAlso trigger(sKey) <> ""
                            itemsConfigured += 1
                        Case InStr(sKey, "LinkgroupListTrigger") > 0 AndAlso trigger(sKey) <> ""
                            itemsConfigured += 1
                    End Select
                Next
                If itemsConfigured = itemsToConfigure Then Configured = True
            End If
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerConfigured returns " & Configured.ToString, LogType.LOG_TYPE_INFO)
            Return Configured
        End Get
    End Property

    Public Function TriggerReferencesDevice(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
        '
        ' Triggers in the sample plug-in do not reference devices, but for demonstration purposes we will pretend they do, 
        '   and that ALL triggers reference our sample devices.
        '
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerReferencesDevice called for instance " & instance & " with TrigInfo = " & TrigInfo.ToString & " and dvRef = " & dvRef.ToString, LogType.LOG_TYPE_INFO)
        'If dvRef = -1 Then Return True
        Return False
    End Function

    Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerFormatUI called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        Dim stb As New StringBuilder
        Dim sKey As String
        Dim PlayerUDN As String = ""
        Dim Command As String = ""
        Dim Linkgroup As String = ""
        Dim InputBox As String = ""
        Dim trigger As New trigger

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        Else
            Return "" ' nothing configured
        End If

        For Each sKey In trigger.Keys
            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerFormatUI found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerUDN = trigger(sKey)
                Case InStr(sKey, "CommandListTrigger") > 0
                    Command = trigger(sKey)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputBox = trigger(sKey)
                Case InStr(sKey, "LinkgroupListTrigger") > 0
                    Linkgroup = trigger(sKey)
            End Select
        Next

        Dim PlayerName = GetZoneByUDN(PlayerUDN)
        Dim CommandPrefix As String
        If MainInstance <> "" Then
            CommandPrefix = "Sonos Instance " & MainInstance
        Else
            CommandPrefix = "Sonos"
        End If

        Select Case TrigInfo.TANumber
            Case 1 ' Sonos Trigger
                Select Case Command
                    Case "Sonos Announcement Start", "Sonos Announcement Stop"
                        stb.Append(CommandPrefix & " trigger - " & Command & " for Linkgroup - " & Linkgroup)
                    Case Else
                        stb.Append(CommandPrefix & " trigger - " & Command & " for player - " & PlayerName)
                End Select
            Case 2 ' Sonos Condition
                Select Case Command
                    Case "hasTrack"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case "hasAlbum"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case "hasArtist"
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for player - " & PlayerName)
                    Case "isPlayingAnnouncement", "isNotPlayingAnnouncement" ' added 7/12/2019 in v3.1.0.31
                        stb.Append(CommandPrefix & " Condition - " & Command & " = " & InputBox & " for Linkgroup - " & Linkgroup)
                    Case Else
                        stb.Append(CommandPrefix & " Condition - " & Command & " for player - " & PlayerName)
                End Select

        End Select


        Return stb.ToString

    End Function

    Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection,
                                         ByVal TrigInfoIn As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerProcessPostUI called for instance " & instance & " with evRef = " & TrigInfoIn.evRef.ToString & " and SubTANumber = " & TrigInfoIn.SubTANumber.ToString & " and TANumber = " & TrigInfoIn.TANumber.ToString & " and UID = " & TrigInfoIn.UID.ToString, LogType.LOG_TYPE_INFO)
        ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn With {
            .sResult = "",
            .DataOut = TrigInfoIn.DataIn,
            .TrigActInfo = TrigInfoIn
        }

        If PostData Is Nothing Then Return Ret
        If PostData.Count < 1 Then Return Ret
        Dim trigger As New trigger
        If Not (TrigInfoIn.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfoIn.DataIn, trigger)
        End If

        Dim parts As Collections.Specialized.NameValueCollection

        Dim sKey As String

        parts = PostData
        Try
            For Each sKey In parts.Keys
                If sKey Is Nothing Then Continue For
                If String.IsNullOrEmpty(sKey.Trim) Then Continue For
                Select Case True
                    Case InStr(sKey, "PlayerListTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "CommandListTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "InputBoxTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                    Case InStr(sKey, "LinkgroupListTrigger") > 0
                        trigger.Add(CObj(parts(sKey)), sKey)
                End Select
            Next
            If Not SerializeObject(trigger, Ret.DataOut) Then
                Ret.sResult = sIFACE_NAME & " Error, Serialization failed. Signal Action not added."
                Return Ret
            End If
        Catch ex As Exception
            Ret.sResult = "ERROR, Exception in Action UI of " & sIFACE_NAME & ": " & ex.Message
            Return Ret
        End Try

        ' All OK
        Ret.sResult = ""
        Return Ret

    End Function

    Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.TriggerTrue
        ' 
        ' Since plug-ins tell HomeSeer when a trigger is true via TriggerFire, this procedure is called just to check
        '   conditions.
        '
        TriggerTrue = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerTrue called for instance " & instance & " with evRef = " & TrigInfo.evRef.ToString & " and SubTANumber = " & TrigInfo.SubTANumber.ToString & " and TANumber = " & TrigInfo.TANumber.ToString & " and UID = " & TrigInfo.UID.ToString, LogType.LOG_TYPE_INFO)
        If TrigInfo.TANumber <> 2 Then Return False ' this should not be!
        If TrigInfo.DataIn Is Nothing Then Return False ' we can't work without data
        Dim trigger As New trigger
        DeSerializeObject(TrigInfo.DataIn, trigger)

        Dim sKey As String
        Dim PlayerUDN As String = ""
        Dim Command As String = ""
        Dim InputBox As String = ""
        Dim Linkgroup As String = ""
        For Each sKey In trigger.Keys
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("TriggerTrue found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
            Select Case True
                Case InStr(sKey, "PlayerListTrigger") > 0
                    PlayerUDN = trigger(sKey)
                Case InStr(sKey, "CommandListTrigger") > 0
                    Command = trigger(sKey)
                Case InStr(sKey, "InputBoxTrigger") > 0
                    InputBox = trigger(sKey)
                Case InStr(sKey, "LinkgroupListTrigger") > 0
                    Linkgroup = trigger(sKey)
            End Select

        Next

        If Command = "isPlayingAnnouncement" Then
            If AnnouncementLink IsNot Nothing Then
                If AnnouncementLink.LinkGroupName = Linkgroup Then
                    Return True
                End If
            End If
            Return False
        End If

        If Command = "isNotPlayingAnnouncement" Then    ' added 7/12/2019 in v3.1.0.31
            If AnnouncementLink IsNot Nothing Then
                If AnnouncementLink.LinkGroupName = Linkgroup Then
                    Return False
                End If
            End If
            Return True
        End If

        Dim MusicApi As HSPI = Nothing
        Try
            MusicApi = MyHSPIControllerRef.GetAPIByUDN(PlayerUDN)
        Catch ex As Exception
        End Try
        If MusicApi Is Nothing Then Return False
        Select Case Command
            Case "IsPlaying"
                If MusicApi.PlayerState = Player_state_values.Playing Then Return True Else Return False
            Case "IsPaused"
                If MusicApi.PlayerState = Player_state_values.Paused Then Return True Else Return False
            Case "IsStopped"
                If MusicApi.PlayerState = Player_state_values.Stopped Then Return True Else Return False
            Case "IsMutted"
                If MusicApi.PlayerMute Then Return True Else Return False
            Case "IsNotMutted"
                If MusicApi.PlayerMute Then Return False Else Return True
            Case "IsNotPlaying"
                If MusicApi.PlayerState <> Player_state_values.Playing Then Return True Else Return False
            Case "IsNotPaused"
                If MusicApi.PlayerState <> Player_state_values.Paused Then Return True Else Return False
            Case "IsNotStopped"
                If MusicApi.PlayerState <> Player_state_values.Stopped Then Return True Else Return False
            Case "hasTrack"
                If Trim(MusicApi.Track.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "hasAlbum"
                If Trim(MusicApi.Album.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "hasArtist"
                If Trim(MusicApi.Artist.ToUpper) = Trim(InputBox.ToUpper) Then Return True Else Return False
            Case "isOnline"
                If MusicApi.DeviceStatus.ToUpper = "ONLINE" Then Return True Else Return False
            Case "isOffline"
                If MusicApi.DeviceStatus.ToUpper = "OFFLINE" Then Return True Else Return False
        End Select

    End Function

#End Region



#Region "    Plug-In Procedures    "


    Public Sub UpdateDevices(ByVal parms As Object)
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("UpdateDevices was called", LogType.LOG_TYPE_INFO)
    End Sub

    Private Function CreatePlayerDevice(ByVal HSRef As Integer, ByVal ZoneName As String, ZoneModel As String, NewDevice As Boolean) As Integer
        CreatePlayerDevice = -1
        Dim dv As Scheduler.Classes.DeviceClass
        Dim DevName As String = "Player" ' ZoneName 'ZoneName & " - Player"
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        Try
            If HSRef = -1 Then
                HSRef = hs.NewDeviceRef(DevName)
                Log("CreatePlayerDevice: Created device " & DevName & " with reference " & HSRef.ToString & " and ZoneModel = " & ZoneModel, LogType.LOG_TYPE_INFO)
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
                NewDevice = True
            End If
            dv = hs.GetDeviceByRef(HSRef)

            If NewDevice Then
                dv.Interface(hs) = sIFACE_NAME
                dv.Location2(hs) = tIFACE_NAME
                dv.InterfaceInstance(hs) = MainInstance
                dv.Location(hs) = ZoneName
                dv.Device_Type_String(hs) = SonosHSDevices.Player.ToString
                dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
                dv.Address(hs) = "Player"
                Dim DT As New DeviceTypeInfo With {
                    .Device_API = DeviceTypeInfo.eDeviceAPI.Media,
                    .Device_Type = DeviceTypeInfo.eDeviceType_Media.Root,
                    .Device_SubType_Description = "Sonos Player Master Control"
                }
                dv.DeviceType_Set(hs) = DT
                dv.Status_Support(hs) = True
                'hs.SetDeviceString(HSRef, "_", False)
                ' dv.Image(hs) is set in myDeviceFinderCallback_DeviceFound after the image was downloaded
                ' This device is a child device, the parent being the root device for the entire security system. 
                ' As such, this device needs to be associated with the root (Parent) device.
                dvParent = hs.GetDeviceByRef(MasterHSDeviceRef)
                If dvParent.AssociatedDevices_Count(Nothing) < 1 Then
                    ' There are none added, so it is OK to add this one.
                    'dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    Dim Found As Boolean = False
                    For Each ref As Integer In dvParent.AssociatedDevices(Nothing)
                        If ref = HSRef Then
                            Found = True
                            Exit For
                        End If
                    Next
                    If Not Found Then
                        'dvParent.AssociatedDevice_Add(hs, HSRef)
                    Else
                        ' This is an error condition likely as this device's reference ID should not already be associated.
                    End If
                End If

                ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
                '   the child's associations.
                dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
                'dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
                dv.Relationship(hs) = Enums.eRelationship.Parent_Root
                hs.DeviceVSP_ClearAll(HSRef, True)
                hs.DeviceVGP_ClearAll(HSRef, True)

                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psPlay,
                    .Status = "Play",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psStop,
                    .Status = "Stop",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psPause,
                    .Status = "Pause",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 1
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)

                If ZoneModel.ToUpper = "WD100" Then
                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                        .PairType = VSVGPairType.SingleValue,
                        .Value = psBuildiPodDB,
                        .Status = "BuildDB",
                        .Render = Enums.CAPIControlType.Button
                    }
                    Pair.Render_Location.Row = 1
                    Pair.Render_Location.Column = 4
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                End If

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psPrevious,
                    .Status = "Prev",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psNext,
                    .Status = "Next",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psShuffle,
                    .Status = "Shuffle",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psRepeat,
                    .Status = "Repeat",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 2
                Pair.Render_Location.Column = 4
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psVolDown,
                    .Status = "Vol - Dn",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 3
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.Range,
                    .Value = psVolSlider,
                    .RangeStart = 0,
                    .RangeEnd = 100,
                    .RangeStatusPrefix = "Volume ",
                    .RangeStatusSuffix = "%",
                    .Render = Enums.CAPIControlType.ValuesRangeSlider
                }
                Pair.Render_Location.Row = 3
                Pair.Render_Location.Column = 2
                'hs.DeviceVSP_AddPair(HSRef, Pair)  ' two sliders on same device doesn't work, neither can I set the value

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psVolUp,
                    .Status = "Vol - Up",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 3
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psMute,
                    .Status = "Mute",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 3
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psBalanceLeft,
                    .Status = "Bal - Left",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 4
                Pair.Render_Location.Column = 1
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.Range,
                    .Value = psBalanceSlider,
                    .RangeStart = 200,
                    .RangeEnd = 400,
                    .RangeStatusPrefix = "Balance L <-> R ",
                    .Render = Enums.CAPIControlType.ValuesRangeSlider
                }
                Pair.Render_Location.Row = 4
                Pair.Render_Location.Column = 2
                'hs.DeviceVSP_AddPair(HSRef, Pair) ' two sliders on same device doesn't work, neither can I set the value

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psBalanceRight,
                    .Status = "Bal - Right",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 4
                Pair.Render_Location.Column = 2
                hs.DeviceVSP_AddPair(HSRef, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                    .PairType = VSVGPairType.SingleValue,
                    .Value = psLoudness,
                    .Status = "Loudness",
                    .Render = Enums.CAPIControlType.Button
                }
                Pair.Render_Location.Row = 4
                Pair.Render_Location.Column = 3
                hs.DeviceVSP_AddPair(HSRef, Pair)
            End If

            Return HSRef
        Catch ex As Exception
            Log("Error in CreatePlayerDevice with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function FindZonePlayers() As PlayerRecord()
        FindZonePlayers = Nothing
        Dim Index As Integer
        Dim ZoneInfo() As PlayerRecord
        ReDim ZoneInfo(0)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers: Attempting to locate all connected ZonePlayers. This may take up to 9 seconds.", LogType.LOG_TYPE_INFO)
        Dim discoveryPort As Integer = GetIntegerIniFile("Options", "SSDPListenerPort", 0)  ' added 10/20/2019 in v.39

        Dim MyDevicesLinkedList As MyUPnPDevices = Nothing
        MyDevicesLinkedList = MySSDPDevice.StartSSDPDiscovery("urn:schemas-upnp-org:device:ZonePlayer:1", discoveryPort) ' ' changed 10/20/2019 in v.39

        If MyDevicesLinkedList Is Nothing Then
            Log("No UPnPDevices found. Please ensure the network is functional and that UPnPDevices devices are attached.", LogType.LOG_TYPE_WARNING)
            'Exit Function removed in version 7. If no players are on the network when we start, the eventhandler will never be called and other devices that come in late are never discovered
        Else    ' moved this code here to avoid error due to MyDeviceLinkedList.Count being Nothing
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers - Discovery succeeded: " & MyDevicesLinkedList.Count & " ZonePlayer(s) found.", LogType.LOG_TYPE_INFO)
        End If

        ' If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers - Discovery succeeded: " & MyDevicesLinkedList.Count & " ZonePlayer(s) found.", LogType.LOG_TYPE_INFO) moved to else case on 9/9/2019

        Dim DeviceCount As Integer = 0
        Try
            If (Not MyDevicesLinkedList Is Nothing) AndAlso MyDevicesLinkedList.Count > 0 Then  ' changed 10/21/2019 to prevent exception when no devices were found
                Index = 0
                For Each Device As MyUPnPDevice In MyDevicesLinkedList
                    If Mid(Device.UniqueDeviceName, 1, 12) = "uuid:RINCON_" And Device.ModelNumber <> "ZB100" And Device.ModelNumber <> "BR200" Then
                        'ReDim Preserve ZoneInfo(Index, 8)
                        ReDim Preserve ZoneInfo(Index)
                        ZoneInfo(Index).UDN = Replace(Device.UniqueDeviceName, "uuid:", "")
                        ZoneInfo(Index).ModelNbr = Device.ModelNumber
                        ZoneInfo(Index).pDevice = Device
                        Dim InArg(0) As Object
                        Dim OutArg(3) As Object ' updated 10/2/2020
                        Try
                            Device.Services.Item("urn:upnp-org:serviceId:DeviceProperties").InvokeAction("GetZoneAttributes", InArg, OutArg)
                            ZoneInfo(Index).ZoneName = OutArg(0) ' this is CurrentZoneName
                            ZoneInfo(Index).PlayerIcon = Replace(OutArg(1), "x-rincon-roomicon:", "") ' this is player icon x-rincon-roomicon:office
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers found Zone Name = " & OutArg(0) & " and Stored in ZoneInfo at Index = " & Index.ToString & " with UDN  = " & Device.UniqueDeviceName & " Friendly Name = " & Device.FriendlyName & " and Icon = " & ZoneInfo(Index).PlayerIcon, LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in FindZonePlayers while getting ZoneAttributes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        ZoneInfo(Index).IPAddress = Device.IPAddress
                        Try
                            Dim IconURL As String = Device.IconURL("image/png", 200, 200, 16) 'image/png image/x-png image/tiff image/bmp image/pjpeg image/jpeg
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers for zoneplayer = " & ZoneInfo(Index).ZoneName & " found IconURL = " & IconURL, LogType.LOG_TYPE_INFO)
                            ZoneInfo(Index).PlayerIcon = IconURL
                        Catch ex As Exception
                            Log("Error in FindZonePlayers for zoneplayer = " & ZoneInfo(Index).ZoneName & ". Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Index = Index + 1
                    ElseIf Mid(Device.UniqueDeviceName, 1, 11) <> "uuid:RINCON" Then
                        ' this is the UPNP service of HS itself on an XP machine responding
                        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "FindZonePlayers found non Sonos device with UDN =  " & Device.UniqueDeviceName & " Friendly Name = " & Device.FriendlyName)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers found non Sonos device with UDN =  " & Device.UniqueDeviceName & " Friendly Name = " & Device.FriendlyName, LogType.LOG_TYPE_WARNING)
                    End If
                Next
                ZoneCount = Index
            End If
            If ZoneCount > 0 Then FindZonePlayers = ZoneInfo
            MyDevicesLinkedList = Nothing
            ZoneInfo = Nothing
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("FindZonePlayers - Discovery succeeded: " & ZoneCount.ToString & " ZonePlayer(s) found.", LogType.LOG_TYPE_INFO)
            Try
                If MySSDPDevice IsNot Nothing Then RemoveHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
            Catch ex As Exception
            End Try
            Try
                If MySSDPDevice IsNot Nothing Then RemoveHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
            Catch ex As Exception
            End Try
            Try
                If MySSDPDevice IsNot Nothing Then AddHandler MySSDPDevice.NewDeviceFound, AddressOf NewDeviceFound
            Catch ex As Exception
                Log("ERROR in FindZonePlayers trying to add a NewDeviceFound Handler with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                If MySSDPDevice IsNot Nothing Then AddHandler MySSDPDevice.MCastDiedEvent, AddressOf MultiCastDiedEvent
            Catch ex As Exception
                Log("ERROR in FindZonePlayers trying to add a MulticastDied Event Handler with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("ERROR in FindZonePlayers. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub NewDeviceFound(inUDN As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("NewDeviceFound called with UDN = " & inUDN, LogType.LOG_TYPE_WARNING)
        Try
            SyncLock (MyNewDiscoveredDeviceQueue)
                MyNewDiscoveredDeviceQueue.Enqueue(inUDN)
            End SyncLock
        Catch ex As Exception
            Log("Error in NewDeviceFound called with UDN = " & inUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        'AddDeviceFlag = True ' this will make it Asynchronic
        MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue
    End Sub

    Private Sub AddNewDiscoveredDevice()
        If NewDeviceHandlerReEntryFlag Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice has Re-Entry while processing Notification queue with # elements = " & MyNewDiscoveredDeviceQueue.Count.ToString, LogType.LOG_TYPE_WARNING)
            'MissedNewDeviceNotificationHandlerFlag = True
            If MyNewDiscoveredDeviceQueue.Count > 0 Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost added v16
            Exit Sub
        End If
        NewDeviceHandlerReEntryFlag = True

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice is processing Notification queue with # elements = " & MyNewDiscoveredDeviceQueue.Count.ToString, LogType.LOG_TYPE_INFO)
        Dim NewUDN As String = ""
        Dim NeedsToBeAdded As Boolean = False
        Try
            While MyNewDiscoveredDeviceQueue.Count > 0
                SyncLock (MyNewDiscoveredDeviceQueue)
                    NewUDN = MyNewDiscoveredDeviceQueue.Dequeue
                End SyncLock
                If NewUDN <> "" Then
                    Dim UPnPDeviceInfo As MyUPnpDeviceInfo = Nothing
                    Dim NewUPnPDevice As MyUPnPDevice = MySSDPDevice.Item("uuid:" & NewUDN, True)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice dequeued UDN = " & NewUDN, LogType.LOG_TYPE_INFO)
                    If NewUPnPDevice Is Nothing Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice dequeued UDN = " & NewUDN & " but found no UPNPDevice", LogType.LOG_TYPE_WARNING)
                    Else
                        If Mid(NewUPnPDevice.UniqueDeviceName, 1, 12) = "uuid:RINCON_" Or Mid(NewUPnPDevice.UniqueDeviceName, 1, 16) = "uuid:DOCKRINCON_" Then
                            NeedsToBeAdded = True
                            Exit Try
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddNewDiscoveredDevice while processing Notification queue with # elements = " & MyNewDiscoveredDeviceQueue.Count.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If MyNewDiscoveredDeviceQueue.Count > 0 Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost added v16

        Dim NewZoneName As String = ""
        Dim NewPlayerIcon As String = ""
        Try
            If NewUDN <> "" And NeedsToBeAdded Then
                ' go find it in the array
                Dim SonosPlayerInfo As MyUPnpDeviceInfo
                Dim DeviceRef As Integer = -1
                SonosPlayerInfo = FindUPnPDeviceInfo(NewUDN)
                Dim NewUPnPDevice As MyUPnPDevice = MySSDPDevice.Item("uuid:" & NewUDN, True)
                If NewUPnPDevice Is Nothing Then
                    GC.Collect()
                    'If MissedNewDeviceNotificationHandlerFlag Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost
                    'MissedNewDeviceNotificationHandlerFlag = False
                    NewDeviceHandlerReEntryFlag = False
                    Exit Sub
                End If
                If Not (Mid(NewUPnPDevice.UniqueDeviceName, 1, 12) = "uuid:RINCON_" And NewUPnPDevice.ModelNumber <> "ZB100" And NewUPnPDevice.ModelNumber <> "BR200") Then
                    GC.Collect()
                    'If MissedNewDeviceNotificationHandlerFlag Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost
                    'MissedNewDeviceNotificationHandlerFlag = False
                    NewDeviceHandlerReEntryFlag = False
                    Exit Sub
                End If
                Dim InArg(0) As Object
                Dim OutArg(3) As Object ' updated 10/2/2020
                Dim NewZoneModel As String = NewUPnPDevice.ModelNumber
                Try
                    NewUPnPDevice.Services.Item("urn:upnp-org:serviceId:DeviceProperties").InvokeAction("GetZoneAttributes", InArg, OutArg)
                    NewZoneName = OutArg(0) ' this is CurrentZoneName
                    NewPlayerIcon = Replace(OutArg(1), "x-rincon-roomicon:", "") ' this is player icon x-rincon-roomicon:office
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice found Zone Name = " & NewZoneName & " with UDN  = " & NewUPnPDevice.UniqueDeviceName & " Friendly Name = " & NewUPnPDevice.FriendlyName & " and Icon = " & NewPlayerIcon, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddNewDiscoveredDevice while getting ZoneAttributes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Dim NewIconURL As String = ""
                Try
                    NewIconURL = NewUPnPDevice.IconURL("image/png", 200, 200, 16) 'image/png image/x-png image/tiff image/bmp image/pjpeg image/jpeg
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice for zoneplayer = " & NewZoneName & " found IconURL = " & NewIconURL, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    Log("Error in AddNewDiscoveredDevice for zoneplayer = " & NewZoneName & ". Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                If SonosPlayerInfo IsNot Nothing Then
                    ' this zone is already known by HS but must have been off-line when the PI started
                    ' there is a device code ... use it and put in object
                    Try
                        ' this zone was not on-line before when the plugin started 
                        SonosPlayerInfo.ZoneName = GetStringIniFile(NewUDN, DeviceInfoIndex.diFriendlyName.ToString, "")
                        SonosPlayerInfo.ZoneModel = NewZoneModel
                        SonosPlayerInfo.Device = NewUPnPDevice
                        SonosPlayerInfo.ZoneOnLine = NewUPnPDevice.Alive
                        SonosPlayerInfo.UPnPDeviceIPAddress = NewUPnPDevice.IPAddress
                        SonosPlayerInfo.UPnPDeviceAdminStateActive = True
                        SonosPlayerInfo.ZoneCurrentIcon = NewPlayerIcon
                        SonosPlayerInfo.ZoneDeviceIconURL = NewIconURL
                        If Not SonosPlayerInfo.ZonePlayerControllerRef Is Nothing Then
                            SonosPlayerInfo.ZonePlayerControllerRef.PassZoneName(SonosPlayerInfo.ZoneName)
                            SonosPlayerInfo.ZonePlayerControllerRef.DirectConnect(SonosPlayerInfo.Device, SonosPlayerInfo.ZoneUDN)
                        Else
                            CreateOneSonosController(NewUDN)
                        End If
                    Catch ex As Exception
                        Log("Error in AddNewDiscoveredDevice updating ZoneInfo. UDN = " & NewUDN & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice found Zonename = " & NewZoneName & " on line.", LogType.LOG_TYPE_INFO)
                Else
                    ' this is a new zone, go create it
                    ' create this new device
                    If NewZoneModel.ToUpper <> "SUB" Then
                        DeviceRef = CreatePlayerDevice(-1, NewZoneName, NewZoneModel, True)
                        If DeviceRef = -1 Then
                            If piDebuglevel > DebugLevel.dlOff Then Log("Error in AddNewDiscoveredDevice for Zonename = " & NewZoneName & ". HS returned -1 as a reference!!!", LogType.LOG_TYPE_ERROR)
                            Exit Sub
                        End If
                        ' save it in the ini file
                        WriteIntegerIniFile("UPnP UDN to HSRef", NewUDN, DeviceRef)
                        WriteStringIniFile("UPnP HSRef to UDN", DeviceRef, NewUDN)
                    Else
                        DeviceRef = -1
                    End If
                    Try
                        Dim AnythingExist As String = GetStringIniFile("UPnP Devices UDN to Info", NewUDN, "")
                        If AnythingExist = "" Then
                            WriteStringIniFile("UPnP Devices UDN to Info", NewUDN, NewZoneName)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diFriendlyName.ToString, NewZoneName)
                            WriteBooleanIniFile(NewUDN, DeviceInfoIndex.diAdminState.ToString, True)
                            WriteBooleanIniFile(NewUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, True)
                            WriteIntegerIniFile(NewUDN, DeviceInfoIndex.diPlayerHSRef.ToString, DeviceRef)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diSonosPlayerType.ToString, NewZoneModel.ToUpper)
                            WriteStringIniFile(NewUDN, DeviceInfoIndex.diSonosReplicationInfo.ToString, "")
                        End If
                        ' the rest we can safely overwrite
                        WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPAddress.ToString, NewUPnPDevice.IPAddress)
                        WriteStringIniFile(NewUDN, DeviceInfoIndex.diIPPort.ToString, NewUPnPDevice.IPPort)
                        WriteStringIniFile(NewUDN, DeviceInfoIndex.diRoomIcon.ToString, NewPlayerIcon)
                        WriteStringIniFile(NewUDN, DeviceInfoIndex.diDeviceIConURL.ToString, NewIconURL)
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice has logged new device with ZoneName = " & NewZoneName, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        Log("Error in AddNewDiscoveredDevice writing info to ini file for Zoneplayer = " & NewZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        GC.Collect()
                        'If MissedNewDeviceNotificationHandlerFlag Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost
                        'MissedNewDeviceNotificationHandlerFlag = False
                        NewDeviceHandlerReEntryFlag = False
                        Exit Sub
                    End Try
                    Try
                        Dim _SonosPlayerInfo As MyUPnpDeviceInfo
                        _SonosPlayerInfo = DMAdd()
                        _SonosPlayerInfo.ZoneName = NewZoneName
                        _SonosPlayerInfo.ZoneModel = NewZoneModel
                        _SonosPlayerInfo.ZoneUDN = NewUDN
                        _SonosPlayerInfo.Device = NewUPnPDevice
                        _SonosPlayerInfo.ZonePlayerRef = DeviceRef
                        _SonosPlayerInfo.ZoneOnLine = NewUPnPDevice.Alive
                        _SonosPlayerInfo.UPnPDeviceIPAddress = NewUPnPDevice.IPAddress
                        _SonosPlayerInfo.UPnPDeviceAdminStateActive = True
                        _SonosPlayerInfo.UPnPDeviceIsAddedToHS = True
                    Catch ex As Exception
                        Log("Error in AddNewDiscoveredDevice while adding the UPnPDevices for Zoneplayer = " & NewZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    CreateOneSonosController(NewUDN)
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddNewDiscoveredDevice successfully created a new device for Transport. Zonename = " & NewZoneName, LogType.LOG_TYPE_INFO)
                End If
            Else
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddNewDiscoveredDevice; empty ZoneName/UDN. ZoneName = " & NewZoneName & ". ZoneUDN = " & NewUDN, LogType.LOG_TYPE_ERROR)
            End If
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AddNewDiscoveredDevice adding Player = " & NewZoneName & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        GC.Collect()
        'If MissedNewDeviceNotificationHandlerFlag Then MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue ' rearm the timer to prevent events from getting lost
        'MissedNewDeviceNotificationHandlerFlag = False
        NewDeviceHandlerReEntryFlag = False
    End Sub

    Public Sub MultiCastDiedEvent()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error. MultiCastDiedEvent received. Terminating the PI to try to restart it", LogType.LOG_TYPE_ERROR)
        'ShutdownIO() removed dcor in version .
    End Sub

    Private Sub CreateOneSonosController(NewUDN As String)
        ' only create Controller instances for Transport devices, which don't have a Controller instance yet
        Dim SonosPlayer As HSPI
        Dim SonosPlayerInfo As MyUPnpDeviceInfo
        Dim DeviceRef As Integer = -1
        SonosPlayerInfo = FindUPnPDeviceInfo(NewUDN)
        If SonosPlayerInfo Is Nothing Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CreateOneSonosController while retrieving the SonosPlayerInfo for UDN = uuid:" & NewUDN, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Try
            If MainInstance <> "" Then
                SonosPlayer = AddInstance(MainInstance & "-" & SonosPlayerInfo.ZoneUDN)
            Else
                SonosPlayer = AddInstance(SonosPlayerInfo.ZoneUDN)
            End If
            Thread.Sleep(1000)    ' takes some time for new instance to connect and callbacks for registering pages to complete
            SonosPlayerInfo.ZonePlayerControllerRef = SonosPlayer
        Catch ex As Exception
            Log("Error in CreateOneSonosController. Could not instantiate ZonePlayerController with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        If SonosPlayer Is Nothing Then
            Log("Error in CreateOneSonosController while instantiating the SonosPlayerInfo a nill instance was returned for UDN = uuid:" & NewUDN, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Try
            SonosPlayer.PassZoneName(SonosPlayerInfo.ZoneName)
            SonosPlayer.PassUDN(Replace(SonosPlayerInfo.ZoneUDN, "uuid:", ""))
            SonosPlayer.ZoneModel = SonosPlayerInfo.ZoneModel
            SonosPlayer.SetHSDeviceRefPlayer(SonosPlayerInfo.ZonePlayerRef)
            SonosPlayer.HSPIControllerRef = Me
            Dim APIIndex As Integer = GetIntegerIniFile(Replace(SonosPlayerInfo.ZoneUDN, "uuid:", ""), DeviceInfoIndex.diDeviceAPIIndex.ToString, 0)
            If APIIndex = 0 Then
                SonosPlayer.HSTMusicIndex = GetNextFreeDeviceIndex()
            Else
                SonosPlayer.HSTMusicIndex = APIIndex
            End If
            SonosPlayer.SetAdministrativeState(True)
            If SonosPlayerInfo.ZoneOnLine Then
                If SonosPlayerInfo.Device IsNot Nothing Then
                    SonosPlayer.DirectConnect(SonosPlayerInfo.Device, SonosPlayerInfo.ZoneUDN)
                End If
            End If
            'If SonosPlayerInfo.Device IsNot Nothing Then
            'SonosPlayer.DirectConnect(SonosPlayerInfo.Device, SonosPlayerInfo.ZoneUDN)
            'End If
        Catch ex As Exception
            Log("Error in CreateOneSonosController. Could not create instance of ZonePlayerController for zoneplayer = " & SonosPlayerInfo.ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        SonosPlayerInfo.ZoneDeviceAPIIndex = SonosPlayer.HSTMusicIndex
        WriteIntegerIniFile(Replace(SonosPlayerInfo.ZoneUDN, "uuid:", ""), DeviceInfoIndex.diDeviceAPIIndex.ToString, SonosPlayer.HSTMusicIndex)
        If SonosPlayer.ZoneModel.ToUpper <> "SUB" And SonosPlayer.ZoneModel.ToUpper <> "ZB100" And SonosPlayer.ZoneModel.ToUpper <> "ZB200" And Not SonosPlayerInfo.ZoneWeblinkCreated Then
            SonosPlayer.CreateWebLink(SonosPlayerInfo.ZoneName, SonosPlayerInfo.ZoneUDN)
            SonosPlayerInfo.ZoneWeblinkCreated = True
        End If
        CheckForDuplicateZoneNames()
        Log("CreateOneSonosController: Created instance of ZonePlayerController for Zoneplayer = " & SonosPlayerInfo.ZoneName & " with index " & SonosPlayer.HSTMusicIndex.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub BuildSonosInfoList()
        ' This is called first at init
        ' This procedure gets all the players out of the HS Database, look them up in the .ini file and puts them in the SonosplayInfo array
        ' So this is the first place where the zone name could be hosed up because a user can change it in HS
        Dim en As Scheduler.Classes.clsDeviceEnumeration
        Dim Device As Scheduler.Classes.DeviceClass
        Dim SonosPlayerInfo As MyUPnpDeviceInfo
        Try
            en = hs.GetDeviceEnumerator
            While Not en.Finished
                Device = en.GetNext
                If Device.Interface(Nothing) = sIFACE_NAME And Device.InterfaceInstance(Nothing) = MainInstance Then
                    Dim DT As DeviceTypeInfo = Device.DeviceType_Get(Nothing)
                    If DT.Device_SubType_Description = "Sonos Player Master Control" Then
                        SonosPlayerInfo = DMAdd()
                        Dim PlayerUDN As String = GetStringIniFile("UPnP HSRef to UDN", Device.Ref(Nothing), "")
                        If PlayerUDN <> "" Then
                            SonosPlayerInfo.ZoneName = GetStringIniFile(PlayerUDN, DeviceInfoIndex.diFriendlyName.ToString, "")
                            SonosPlayerInfo.ZoneUDN = PlayerUDN
                            SonosPlayerInfo.ZoneModel = GetStringIniFile(PlayerUDN, DeviceInfoIndex.diSonosPlayerType.ToString, "")
                            SonosPlayerInfo.ZonePlayerRef = Device.Ref(Nothing)
                            'SonosPlayerInfo.ZoneDeviceAPIIndex = GetStringIniFile(PlayerUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, "")
                            SonosPlayerInfo.UPnPDeviceAdminStateActive = GetBooleanIniFile(PlayerUDN, DeviceInfoIndex.diAdminState.ToString, False)
                            SonosPlayerInfo.UPnPDeviceIPAddress = GetStringIniFile(PlayerUDN, DeviceInfoIndex.diIPAddress.ToString, "")
                        Else
                            Log("Error in BuildSonosInfoList. Info not found in .ini file Devtype = " & Device.Device_Type_String(Nothing) & " deviceReference=" & Device.Ref(Nothing).ToString, LogType.LOG_TYPE_ERROR)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            Log("Error in BuildSonosInfoList with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub BuildHSSonosDevices(ByVal Refresh As Boolean)
        ' this is called after BuildSonosInfoList was called and pre CreateSonosControllers
        Dim DeviceRef As Integer = -1
        Dim ZoneName As String
        Dim ZoneUDN As String
        Dim ZoneModel As String
        Dim ZoneIPAddress As String = ""
        Dim SonosPlayerInfo As MyUPnpDeviceInfo
        Dim ZoneInfo As PlayerRecord() = Nothing
        Dim ZoneRoomIcon As String = ""
        Dim ZoneDeviceIconURL = ""

        Try
            ZoneInfo = FindZonePlayers() ' go discover the zone players, than map or create HS devices
        Catch ex As Exception
            Log("Error in BuildHSSonosDevices while finding the zoneplayers : " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            For I = 0 To (ZoneCount - 1)
                ZoneName = ZoneInfo(I).ZoneName
                ZoneModel = ZoneInfo(I).ModelNbr
                ZoneIPAddress = ZoneInfo(I).IPAddress
                ZoneRoomIcon = ZoneInfo(I).PlayerIcon
                ZoneDeviceIconURL = ZoneInfo(I).IconURL
                If ZoneModel.ToUpper = "SUB" Then
                    ZoneName = "SUB"
                End If
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildHSSonosDevices: looking for " & ZoneName & " in ZoneInfo and Refresh = " & Refresh.ToString, LogType.LOG_TYPE_INFO)
                ZoneUDN = ZoneInfo(I).UDN
                If ZoneUDN <> "" Then
                    ' go find it in the array
                    SonosPlayerInfo = FindUPnPDeviceInfo(ZoneUDN)
                    If Not SonosPlayerInfo Is Nothing Then
                        ' this zone is already known by HS
                        ' there is a device code ... use it and put in object
                        CreatePlayerDevice(SonosPlayerInfo.ZonePlayerRef, ZoneName, ZoneModel, Refresh) ' this will update fields in between SW versions
                        If ZoneInfo(I).pDevice IsNot Nothing Then
                            SonosPlayerInfo.ZoneOnLine = ZoneInfo(I).pDevice.Alive
                        Else
                            SonosPlayerInfo.ZoneOnLine = False
                        End If
                        SonosPlayerInfo.ZoneModel = ZoneModel
                        SonosPlayerInfo.Device = ZoneInfo(I).pDevice
                        SonosPlayerInfo.UPnPDeviceIPAddress = ZoneIPAddress
                        SonosPlayerInfo.UPnPDeviceAdminStateActive = True ' dcor needs fixing
                        SonosPlayerInfo.ZoneCurrentIcon = ZoneRoomIcon
                        SonosPlayerInfo.ZoneDeviceIconURL = ZoneDeviceIconURL
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildHSSonosDevices found Zone Name = " & ZoneName, LogType.LOG_TYPE_INFO)
                    Else
                        ' this is a new zone, go create it
                        ' create this new device
                        Dim PlayerUDN As String = Replace(ZoneUDN, "uuid:", "")
                        If ZoneModel.ToUpper <> "SUB" Then
                            DeviceRef = CreatePlayerDevice(-1, ZoneName, ZoneModel, True)
                            If DeviceRef = -1 Then
                                If piDebuglevel > DebugLevel.dlOff Then Log("Error in BuildHSSonosDevices for Zonename = " & ZoneName & ". HS returned -1 as a reference!!!", LogType.LOG_TYPE_ERROR)
                                Exit Sub
                            End If
                            ' save it in the ini file
                            WriteIntegerIniFile("UPnP UDN to HSRef", PlayerUDN, DeviceRef)
                            WriteStringIniFile("UPnP HSRef to UDN", DeviceRef, PlayerUDN)
                        Else
                            DeviceRef = -1
                        End If

                        '
                        Try
                            Dim AnythingExist As String = GetStringIniFile("UPnP Devices UDN to Info", PlayerUDN, "")
                            If AnythingExist = "" Then
                                WriteStringIniFile("UPnP Devices UDN to Info", PlayerUDN, ZoneName)
                                WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diFriendlyName.ToString, ZoneName)
                                WriteBooleanIniFile(PlayerUDN, DeviceInfoIndex.diAdminState.ToString, True)
                                WriteBooleanIniFile(PlayerUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, True)
                                WriteIntegerIniFile(PlayerUDN, DeviceInfoIndex.diPlayerHSRef.ToString, DeviceRef)
                                'WriteIntegerIniFile(PlayerUDN, DeviceInfoIndex.diDeviceAPIIndex.ToString, 0)
                                WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diSonosPlayerType.ToString, ZoneModel.ToUpper)
                                WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diSonosReplicationInfo.ToString, "")
                            End If
                            ' the rest we can safely overwrite
                            WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diIPAddress.ToString, ZoneIPAddress)
                            WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diRoomIcon.ToString, ZoneRoomIcon)
                            WriteStringIniFile(PlayerUDN, DeviceInfoIndex.diDeviceIConURL.ToString, ZoneDeviceIconURL)
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildHSSonosDevices has logged new device with UPnPDeviceName = " & ZoneName, LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            Log("Error in DetectUPnPDevices 1 while adding the UPnPDevices with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Sub
                        End Try
                        Try
                            Dim _SonosPlayerInfo As MyUPnpDeviceInfo
                            _SonosPlayerInfo = DMAdd()
                            _SonosPlayerInfo.ZoneName = ZoneName
                            _SonosPlayerInfo.ZoneModel = ZoneModel
                            _SonosPlayerInfo.ZoneUDN = ZoneUDN
                            _SonosPlayerInfo.Device = ZoneInfo(I).pDevice
                            _SonosPlayerInfo.ZonePlayerRef = DeviceRef
                            If ZoneInfo(I).pDevice IsNot Nothing Then
                                _SonosPlayerInfo.ZoneOnLine = ZoneInfo(I).pDevice.Alive
                            Else
                                _SonosPlayerInfo.ZoneOnLine = False
                            End If
                            _SonosPlayerInfo.UPnPDeviceIPAddress = ZoneIPAddress
                            _SonosPlayerInfo.UPnPDeviceAdminStateActive = True
                            _SonosPlayerInfo.UPnPDeviceIsAddedToHS = True
                        Catch ex As Exception
                            Log("Error in DetectUPnPDevices 2 while adding the UPnPDevices with Index = " & I.ToString & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try

                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildHSSonosDevices: Creating new device for Transport. Zonename = " & ZoneName, LogType.LOG_TYPE_INFO)
                    End If
                Else
                    Log("BuildHSSonosDevices: Error. Looking for empty ZoneName/UDN. ZoneName = " & ZoneName & ". ZoneUDN = " & ZoneUDN, LogType.LOG_TYPE_ERROR)
                End If
                Try ' this was added to v.80 to try to avoid handle leak as reported by ed tenholde
                    ZoneInfo(I).UDN = ""
                    ZoneInfo(I).ZoneName = ""
                    ZoneInfo(I).ModelNbr = ""
                    ZoneInfo(I).pDevice = Nothing
                    ZoneInfo(I).IPAddress = ""
                    ZoneInfo(I).PlayerIcon = ""
                    ZoneInfo(I).IconURL = ""
                Catch ex As Exception
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in BuildHSSonosDeviceswhen clearing out ZoneInfo with error = " & ex.Message, LogType.LOG_TYPE_WARNING)
                End Try
            Next
        Catch ex As Exception
            Log("Error in BuildHSSonosDevices with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ZoneInfo = Nothing
        GC.Collect()
    End Sub

    Public Function CheckPlayerIsPairable(PlayerType As String) As Boolean
        CheckPlayerIsPairable = False
        If PlayerType = "S5" Or PlayerType = "S3" Or PlayerType = "S1" Or PlayerType = "S12" Or PlayerType = "S13" Or PlayerType = "S6" Or PlayerType = "S18" Or PlayerType = "S21" Then
            ' added playertype 18 (S1 type) on 7/12/2019 in v3.1.0.31
            Return True
        End If
    End Function

    'added 7/12/2019 in v3.1.0.31
    Public Function CheckPlayerCanPlayTV(PlayerType As String) As Boolean
        CheckPlayerCanPlayTV = False
        If PlayerType = "S9" Or PlayerType = "S11" Or PlayerType = "S14" Or PlayerType = "S16" Then ' added S16 (amp) on 7/12/2019 in v3.1.0.31
            ' S9 in playbar
            ' S11 is playbase
            ' S14 is playbeam
            ' S16 AMP
            Return True
        End If
    End Function

    'added 7/12/2019 in v3.1.0.31
    Private Function CheckPlayerHasAudioInput(PlayerType As String) As Boolean
        CheckPlayerHasAudioInput = False
        If (PlayerType <> "SUB") And (PlayerType <> "WD100") And (PlayerType <> "S1") And (PlayerType <> "S12") And (PlayerType <> "S13") And (PlayerType <> "S3") Then
            Return True
        End If
    End Function

    Private Function ZoneNamesHaveChanged(ByVal SonosPlayerInfo As MyUPnpDeviceInfo, ByVal inZoneName As String, ByVal ZoneModel As String) As Boolean
        ZoneNamesHaveChanged = False
        'If CheckPlayerIsPairable(SonosPlayerInfo.ZoneModel) Then ' removed in v23 because we can also pair a Connect:Amp to a playbar
        ' either pairing change or zone Name change
        Dim SonosPlayer As HSPI = Nothing
        SonosPlayer = GetAPIByUDN(SonosPlayerInfo.ZoneUDN)
        If Not SonosPlayer Is Nothing Then
            Try
                If SonosPlayer.MyChannelMapSet <> "" Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ZoneNamesHaveChanged was called for " & SonosPlayerInfo.ZoneName & " but no changes were made because the zone is paired", LogType.LOG_TYPE_INFO)
                    Exit Function
                ElseIf SonosPlayer.MyHTSatChanMapSet <> "" Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ZoneNamesHaveChanged was called for " & SonosPlayerInfo.ZoneName & " but no changes were made because the zone is paired to Playbar", LogType.LOG_TYPE_INFO)
                    Exit Function
                Else
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ZoneNamesHaveChanged was called for " & SonosPlayerInfo.ZoneName & " but no ChannelMapSet or HTSatChanMapSet info found", LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
                Log("Error in ZoneNamesHaveChanged for " & SonosPlayerInfo.ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            Log("Error in ZoneNamesHaveChanged for " & SonosPlayerInfo.ZoneName & " no player info found", LogType.LOG_TYPE_ERROR)
        End If
        'End If
        Log("WARNING ZoneNamesHaveChanged called and changed " & SonosPlayerInfo.ZoneName & " to " & inZoneName, LogType.LOG_TYPE_WARNING)
        WriteStringIniFile("UPnP Devices UDN to Info", SonosPlayerInfo.ZoneUDN, inZoneName)
        WriteStringIniFile(SonosPlayerInfo.ZoneUDN, DeviceInfoIndex.diFriendlyName.ToString, inZoneName)
        ZoneNamesHaveChanged = True
    End Function

    Private Sub CheckForDuplicateZoneNames()

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForDuplicateZoneNames called", LogType.LOG_TYPE_INFO)

        If MyHSDeviceLinkedList.Count = 0 Then
            Exit Sub
        End If

        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                Dim SonosPlayer As HSPI = Nothing
                SonosPlayer = GetAPIByUDN(HSDevice.ZoneUDN)
                If SonosPlayer IsNot Nothing Then
                    If SonosPlayer.ZoneIsASlave Then
                        If CheckZoneNameAlreadyExists(SonosPlayer.ZoneName, SonosPlayer.UDN) Then
                            ' we need to rename the zone
                            Dim NewSuffix As String = ""
                            If SonosPlayer.MyZoneIsPairSlave Then
                                If SonosPlayer.MyZonePairLeftFrontUDN = SonosPlayer.UDN Then
                                    NewSuffix = "_LF"
                                Else
                                    NewSuffix = "_RF"
                                End If
                            ElseIf (SonosPlayer.MyZonePlayBarLeftRearUDN = SonosPlayer.UDN) And (SonosPlayer.MyZonePlayBarRightRearUDN = SonosPlayer.UDN) Then
                                NewSuffix = "_LRR"
                            ElseIf SonosPlayer.MyZonePlayBarLeftRearUDN = SonosPlayer.UDN Then
                                NewSuffix = "_LR"
                            ElseIf SonosPlayer.MyZonePlayBarRightRearUDN = SonosPlayer.UDN Then
                                NewSuffix = "_RR"
                            ElseIf SonosPlayer.MyZonePlayBarLeftFrontUDN = SonosPlayer.UDN Then
                                NewSuffix = "_LF"
                            ElseIf SonosPlayer.MyZonePlayBarRightFrontUDN = SonosPlayer.UDN Then
                                NewSuffix = "_RF"
                            End If
                            If NewSuffix <> "" Then ZoneNameChanged(SonosPlayer.ZoneName, SonosPlayer.UDN, SonosPlayer.ZoneName & NewSuffix, True)
                        End If
                    End If
                    'End If
                End If
            Next
        Catch ex As Exception
            Log("Error in CheckForDuplicateZoneNames with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Function CheckZoneNameAlreadyExists(inZoneName As String, inZoneUDN As String) As Boolean
        CheckZoneNameAlreadyExists = False
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckZoneNameAlreadyExists called with inZoneName = " & inZoneName & " and inZoneUDN" & inZoneUDN, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then
            Exit Function
        End If
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                Dim SonosPlayer As HSPI = Nothing
                SonosPlayer = GetAPIByUDN(HSDevice.ZoneUDN)
                If SonosPlayer IsNot Nothing Then
                    If SonosPlayer.ZoneName = inZoneName And SonosPlayer.UDN <> inZoneUDN Then
                        Return True
                    End If
                End If
            Next
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CheckZoneNameAlreadyExists with inZoneName = " & inZoneName & " and inZoneUDN" & inZoneUDN & " and error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub CreateSonosControllers()

        If MyHSDeviceLinkedList.Count = 0 Then
            Log("CreateSonosControllers called but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateSonosControllers: found " & MyHSDeviceLinkedList.Count.ToString & " Device Codes", LogType.LOG_TYPE_INFO)

        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If Not HSDevice.ZoneOnLine Then
                ' this player was in the HS list but not detected, it may be off-line at this point
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateSonosControllers: ZonePlayer in HS DB not found on-line. ZoneHSRef = " & HSDevice.ZonePlayerRef.ToString & ". ZoneUDN = " & HSDevice.ZoneUDN, LogType.LOG_TYPE_INFO)
            Else
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "CreateSonosControllers: ZonePlayer in HS DB found on-line. ZoneName = " & SonosPlayersInfo(Index).ZoneName & ". ZoneHSCode = " & SonosPlayersInfo(Index ).ZoneHSCode & ". Zone Device Type = " & SonosPlayersInfo(Index ).ZoneDeviceType & ". ZoneUDN = " & SonosPlayersInfo(Index ).ZoneUDN)
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "CreateSonosControllers: ZonePlayer in HS DB found on-line. ZoneTDC = " & SonosPlayersInfo(Index).ZoneTDC & ". ZoneCDC = " & SonosPlayersInfo(Index).ZoneCDC & ". ZoneRDC = " & SonosPlayersInfo(Index ).ZoneRDC)
            End If
        Next

        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZonePlayerControllerRef Is Nothing Then
                ' only create Controller instances for Transport devices, which don't have a Controller instance yet
                CreateOneSonosController(HSDevice.ZoneUDN)
            End If
        Next
        WriteSonosNamesToIniFile()
    End Sub

    Private Function GetNextFreeDeviceIndex() As Integer
        GetNextFreeDeviceIndex = 0
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeDeviceIndex called", LogType.LOG_TYPE_INFO)
        Try
            Dim SonosDevices As New System.Collections.Generic.Dictionary(Of String, String)()
            SonosDevices = GetIniSection("UPnP Devices UDN to Info") '  As Dictionary(Of String, String)
            If SonosDevices IsNot Nothing Then
                Dim LowestFreeIndex As Integer = 0
                Dim IndexFound As Boolean = True
                While IndexFound
                    LowestFreeIndex = LowestFreeIndex + 1
                    IndexFound = False
                    For Each SonosDevice In SonosDevices
                        If SonosDevice.Key <> "" Then
                            Try
                                If GetBooleanIniFile(SonosDevice.Key, DeviceInfoIndex.diDeviceIsAdded.ToString, False) = True Then
                                    If GetIntegerIniFile(SonosDevice.Key, DeviceInfoIndex.diDeviceAPIIndex.ToString, 0) = LowestFreeIndex Then
                                        IndexFound = True
                                        Exit For
                                    End If
                                End If
                            Catch ex As Exception
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetNextFreeDeviceIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                        End If
                    Next
                    If LowestFreeIndex > 100 Then
                        IndexFound = False ' force an exit
                        LowestFreeIndex = 0
                    End If
                End While
                GetNextFreeDeviceIndex = LowestFreeIndex
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeDeviceIndex found Index = " & LowestFreeIndex.ToString, LogType.LOG_TYPE_INFO)
            Else
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetNextFreeDeviceIndex found no devices in the .ini file", LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
            Log("Error in GetNextFreeDeviceIndex1 with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function


    Private Sub WriteSonosNamesToIniFile()
        Dim SonosZoneNames As String = ""
        Dim SonosUDNs As String = ""

        If MyHSDeviceLinkedList.Count = 0 Then
            Exit Sub
        End If

        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZoneModel.ToUpper <> "SUB" Then
                    If SonosZoneNames <> "" Then SonosZoneNames = SonosZoneNames & ":|:"
                    If SonosUDNs <> "" Then SonosUDNs = SonosUDNs & ":|:"
                    SonosZoneNames = SonosZoneNames & HSDevice.ZoneName & ";:;" & HSDevice.ZoneModel
                    SonosUDNs = SonosUDNs & HSDevice.ZoneUDN
                End If
            Next
            WriteStringIniFile("Sonos Zonenames", "Names", SonosZoneNames)
            WriteStringIniFile("Sonos Zonenames", "UDNs", SonosUDNs)
            If piDebuglevel > DebugLevel.dlEvents Then Log("WriteSonosNamesToIniFile wrote new string " & SonosZoneNames, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in WriteSonosNamesToIniFile creating SonosZoneNames with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DestroySonosControllers()
        Dim SonosPlayer As HSPI
        If MyHSDeviceLinkedList Is Nothing Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroySonosControllers called for instance " & instance & " but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroySonosControllers called for instance " & instance & " but no devices found", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DestroySonosControllers for instance " & instance & " found " & MyHSDeviceLinkedList.Count.ToString & " Device Codes", LogType.LOG_TYPE_INFO)
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If Not HSDevice.ZonePlayerControllerRef Is Nothing Then
                Try
                    SonosPlayer = HSDevice.ZonePlayerControllerRef
                    SonosPlayer.Disconnect(True)
                    SonosPlayer.DeleteWebLink(SonosPlayer.UDN, SonosPlayer.ZonePlayerName)
                    Dim OldUDN = SonosPlayer.UDN ' changed in v3.1.0.25
                    SonosPlayer.DestroyPlayer(True)
                    HSDevice.ZonePlayerControllerRef = Nothing
                    RemoveInstance(OldUDN) ' changed in v3.1.0.25
                    'SonosPlayer.ShutdownIO()' changed in v3.1.0.25
                Catch ex As Exception
                    Log("DestroySonosControllers for instance " & instance & ". Could not disconnect with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    'Exit For
                End Try
            End If
        Next
    End Sub

    Public Function GetSonosPlayerByHSDeviceRef(ByVal hsRef As Integer) As HSPI
        GetSonosPlayerByHSDeviceRef = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error: ZonePlayer not found in GetSonosPlayerByHSDeviceCode. hsRef = " & hsRef.ToString & ". ZoneCount = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            Try
                If HSDevice.ZonePlayerRef = hsRef And Not HSDevice.ZonePlayerControllerRef Is Nothing Then
                    GetSonosPlayerByHSDeviceRef = HSDevice.ZonePlayerControllerRef
                    Exit Function
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error: ZonePlayer not found in GetSonosPlayerByHSDeviceCode. hsRef = " & hsRef.ToString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error: ZonePlayer not found in GetSonosPlayerByHSDeviceCode. hsRef = " & hsRef.ToString & ". ZoneCount = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetSonosPlayerByAPIIndex(ByVal Index As Integer) As HSPI
        GetSonosPlayerByAPIIndex = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonosPlayerByAPIIndex. ZonePlayer not found. ZoneIndex : " & Index.ToString, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneDeviceAPIIndex = Index Then
                GetSonosPlayerByAPIIndex = HSDevice.ZonePlayerControllerRef
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetSonosPlayer found ZonePlayer : " & Index.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonosPlayer. ZonePlayer not found. ZoneIndex : " & Index.ToString, LogType.LOG_TYPE_ERROR)
    End Function

    Public Function GetSonosPlayerByUDN(ByVal Name As String) As HSPI
        GetSonosPlayerByUDN = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonosPlayer. ZonePlayer not found. ZoneName : " & Name, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        If Mid(Name, 1, 5) = "uuid:" Then
            Mid(Name, 1, 5) = "     "
            Name = Trim(Name)
        End If
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If Name = "" And HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    ' special case, just find the first player with a reference
                    If HSDevice.ZonePlayerControllerRef.DeviceStatus.ToUpper = "ONLINE" Then ' pick the first one alive
                        GetSonosPlayerByUDN = HSDevice.ZonePlayerControllerRef
                        Exit Function
                    End If
                ElseIf (HSDevice.ZoneName = Name) Or (HSDevice.ZoneUDN = Name) Then
                    GetSonosPlayerByUDN = HSDevice.ZonePlayerControllerRef
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetSonosPlayer found ZonePlayer : " & Name)
                    Exit Function
                End If
            Next
        Catch ex As Exception
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSonosPlayer. ZonePlayer not found. ZoneName : " & Name, LogType.LOG_TYPE_ERROR)
    End Function

    Private Function FindUPnPDeviceInfo(ByVal UDN As String) As MyUPnpDeviceInfo
        FindUPnPDeviceInfo = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlEvents Then Log("Warning in FindUPnPDeviceInfo for UDN = " & UDN & ". The array does not exist ", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZoneUDN = UDN Then
                    If piDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDeviceInfo found UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_INFO)
                    FindUPnPDeviceInfo = HSDevice
                    Exit Function
                Else
                    If piDebuglevel > DebugLevel.dlEvents Then Log("FindUPnPDeviceInfo did not find UDN = " & UDN & " but found = " & HSDevice.ZoneUDN, LogType.LOG_TYPE_INFO)
                End If
            Next
        Catch ex As Exception
            Log("Error in FindUPnPDeviceInfo Finding UPnPDevicveInfo. UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlEvents Then Log("Warning in FindUPnPDeviceInfo did not find UDN = " & UDN & ". Array Size = " & MyHSDeviceLinkedList.Count.ToString, LogType.LOG_TYPE_WARNING)
    End Function

    Private Sub BuildButtonStringRef(Ref As Integer)
        Dim NbrOfLinkgroupZones As Integer = 0
        Dim LinkgroupZones() As String
        Dim LinkgroupZone As String
        hs.DeviceVSP_ClearAll(Ref, True)
        Try
            Dim Pair As VSPair

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msDisconnected,
                .Status = "Disconnected"
            }
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msInitializing,
                .Status = "Initializing"
            }
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msConnected,
                .Status = "Connected"
            }
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msBuildingDB,
                .Status = "Building DB"
            }
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msPlayAll,
                .Status = "All Zones Play",
                .Render = Enums.CAPIControlType.Button
            }
            Pair.Render_Location.Row = 1
            Pair.Render_Location.Column = 1
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msPauseAll,
                .Status = "All Zones Pause",
                .Render = Enums.CAPIControlType.Button
            }
            Pair.Render_Location.Row = 1
            Pair.Render_Location.Column = 2
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msMuteAll,
                .Status = "All Zones Mute On",
                .Render = Enums.CAPIControlType.Button
            }
            Pair.Render_Location.Row = 2
            Pair.Render_Location.Column = 1
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msUnmuteAll,
                .Status = "All Zones Mute Off",
                .Render = Enums.CAPIControlType.Button
            }
            Pair.Render_Location.Row = 2
            Pair.Render_Location.Column = 2
            hs.DeviceVSP_AddPair(Ref, Pair)

            Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                .PairType = VSVGPairType.SingleValue,
                .Value = msBuildDB,
                .Status = "BuildDB",
                .Render = Enums.CAPIControlType.Button
            }
            Pair.Render_Location.Row = 3
            Pair.Render_Location.Column = 1
            hs.DeviceVSP_AddPair(Ref, Pair)

            LinkgroupZone = GetStringIniFile("LinkgroupNames", "Names", "")
            If LinkgroupZone <> "" Then
                LinkgroupZones = Split(LinkgroupZone, "|")

                Dim ButtonValue As Integer = LinkGroupButtonOffset
                Dim RowIndex As Integer = 4

                For Each LinkgroupZone In LinkgroupZones
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("BuildButtonStringRef Found LinkgroupZoneSource = " & LinkgroupZone, LogType.LOG_TYPE_INFO)

                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                        .PairType = VSVGPairType.SingleValue,
                        .Value = ButtonValue,
                        .Status = "Link-" & LinkgroupZone,
                        .Render = Enums.CAPIControlType.Button
                    }
                    Pair.Render_Location.Row = RowIndex
                    Pair.Render_Location.Column = 1
                    hs.DeviceVSP_AddPair(Ref, Pair)

                    Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control) With {
                        .PairType = VSVGPairType.SingleValue,
                        .Value = ButtonValue + 1,
                        .Status = "Unlink-" & LinkgroupZone,
                        .Render = Enums.CAPIControlType.Button
                    }
                    Pair.Render_Location.Row = RowIndex
                    Pair.Render_Location.Column = 2
                    hs.DeviceVSP_AddPair(Ref, Pair)
                    ButtonValue += 2
                    RowIndex += 1
                Next
            End If
        Catch ex As Exception
            Log("Error in BuildButtonStringRef with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub ReadIniFile()
        upnpDebuglevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlErrorsOnly)
        piDebuglevel = GetIntegerIniFile("Options", "PIDebugLevel", DebugLevel.dlErrorsOnly)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadIniFile called", LogType.LOG_TYPE_INFO)

        Try
            If GetStringIniFile("Options", "NoRediscovery", "") = "" Then
                WriteBooleanIniFile("Options", "NoRediscovery", True)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile accessing NoRediscovery with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            PreviousVersion = GetIntegerIniFile("Options", "PreviousVersion", 0)
        Catch ex As Exception
            Log("Error in ReadIniFile reading PreviousVersion with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "MaxNbrofUPNPObjects", "") = "" Then
                WriteIntegerIniFile("Options", "MaxNbrofUPNPObjects", cMaxNbrOfUPNP)
            End If
            MaxNbrOfUPNPObjects = GetIntegerIniFile("Options", "MaxNbrofUPNPObjects", -1)
            If MaxNbrOfUPNPObjects = -1 Or MaxNbrOfUPNPObjects > cMaxNbrOfUPNP Then
                MaxNbrOfUPNPObjects = cMaxNbrOfUPNP ' = maximum
                WriteIntegerIniFile("Options", "MaxNbrofUPNPObjects", cMaxNbrOfUPNP)
            End If
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("INIT: MaxNbrOfUPNPObjects set to " & MaxNbrOfUPNPObjects, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in ReadIniFile reading MaxNbrOfUPNPObjects with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "Auto Update", "") = "" Then
                WriteBooleanIniFile("Options", "Auto Update", False)
                AutoUpdate = False
            Else
                AutoUpdate = GetBooleanIniFile("Options", "Auto Update", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Auto Update with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            AutoUpdateTime = GetStringIniFile("Options", "Auto Update Time", "")
            If AutoUpdateTime = "" Then
                WriteStringIniFile("Options", "Auto Update Time", "02:10")
            End If
            'AutoUpdateZoneName = GetStringIniFile("Options", "Auto Update Zone", "")
        Catch ex As Exception
            Log("Error in ReadIniFile reading Auto Update Zone with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            DBZoneName = GetStringIniFile("Options", "DB Zone", "")
            Dim DBItemsString As String = ""
            DBItemsString = Trim(GetStringIniFile("Options", "DBItems", ""))
            Dim DBItems
            Dim DBItem
            If DBItemsString <> "" Then
                ' this is the old way
                DBItems = Split(DBItemsString, "|")
                For Each DBItem In DBItems
                    Select Case DBItem
                        Case "Genres"
                            MyMusicDBItems.Genres = True
                        Case "Tracks"
                            MyMusicDBItems.Tracks = True
                        Case "Artists"
                            MyMusicDBItems.Artists = True
                        Case "Albums"
                            MyMusicDBItems.Albums = True
                        Case "PlayLists"
                            MyMusicDBItems.Playlists = True
                        Case "RadioStations"
                            MyMusicDBItems.Radiostations = True
                        Case "Audiobooks"
                            MyMusicDBItems.Audiobooks = True
                        Case "Podcasts"
                            MyMusicDBItems.Podcasts = True
                    End Select
                Next
                ' now write the new way
                WriteBooleanIniFile("DBItems", "Genres", MyMusicDBItems.Genres)
                WriteBooleanIniFile("DBItems", "Tracks", MyMusicDBItems.Tracks)
                WriteBooleanIniFile("DBItems", "Artists", MyMusicDBItems.Artists)
                WriteBooleanIniFile("DBItems", "Albums", MyMusicDBItems.Albums)
                WriteBooleanIniFile("DBItems", "PlayLists", MyMusicDBItems.Playlists)
                WriteBooleanIniFile("DBItems", "RadioStations", MyMusicDBItems.Radiostations)
                WriteBooleanIniFile("DBItems", "AudioBooks", MyMusicDBItems.Audiobooks)
                WriteBooleanIniFile("DBItems", "Podcasts", MyMusicDBItems.Podcasts)
                ' now delete the old string
                DeleteEntryIniFile("Options", "DBItems")
            ElseIf GetStringIniFile("DBItems", "Genres", "") = "" Then
                ' this should not be, this means we have the old way
                ' now write the new way
                WriteBooleanIniFile("DBItems", "Genres", True)
                WriteBooleanIniFile("DBItems", "Tracks", True)
                WriteBooleanIniFile("DBItems", "Artists", True)
                WriteBooleanIniFile("DBItems", "Albums", True)
                WriteBooleanIniFile("DBItems", "PlayLists", True)
                WriteBooleanIniFile("DBItems", "RadioStations", True)
                WriteBooleanIniFile("DBItems", "AudioBooks", True)
                WriteBooleanIniFile("DBItems", "Podcasts", True)
                ' now remove any remnants
                DeleteEntryIniFile("Options", "DBItems")
            Else
                ' this is the new way
                MyMusicDBItems.Genres = GetBooleanIniFile("DBItems", "Genres", False)
                MyMusicDBItems.Tracks = GetBooleanIniFile("DBItems", "Tracks", False)
                MyMusicDBItems.Artists = GetBooleanIniFile("DBItems", "Artists", False)
                MyMusicDBItems.Albums = GetBooleanIniFile("DBItems", "Albums", False)
                MyMusicDBItems.Playlists = GetBooleanIniFile("DBItems", "PlayLists", False)
                MyMusicDBItems.Radiostations = GetBooleanIniFile("DBItems", "RadioStations", False)
                MyMusicDBItems.Audiobooks = GetBooleanIniFile("DBItems", "AudioBooks", False)
                MyMusicDBItems.Podcasts = GetBooleanIniFile("DBItems", "Podcasts", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading DBItems with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            VolumeStep = GetIntegerIniFile("Options", "VolumeStep", 0)
            If VolumeStep = 0 Then
                VolumeStep = 5
                WriteIntegerIniFile("Options", "VolumeStep", 5)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading VolumeStep with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        SetSpeakerProxy() ' changed in v.19 to set proxy AFTER init complete
        Try
            If GetStringIniFile("Options", "Learn RadioStations", "") = "" Then
                WriteBooleanIniFile("Options", "Learn RadioStations", True)
                LearnRadioStations = True
            Else
                LearnRadioStations = GetBooleanIniFile("Options", "Learn RadioStations", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Learn RadioStations flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AutoBuildDockedDB", "") = "" Then
                WriteBooleanIniFile("Options", "AutoBuildDockedDB", False)
                AutoBuildDockedDB = False
            Else
                AutoBuildDockedDB = GetBooleanIniFile("Options", "AutoBuildDockedDB", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AutoBuildDockedDB flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            NbrOfPingRetries = GetIntegerIniFile("Options", "NbrOfPingRetries", 0)
            If NbrOfPingRetries = 0 Then
                NbrOfPingRetries = 3
                WriteIntegerIniFile("Options", "NbrOfPingRetries", 3)
            End If
            If GetStringIniFile("Options", "ShowFailedPings", "") = "" Then
                WriteBooleanIniFile("Options", "ShowFailedPings", False)
                ShowFailedPings = False
            Else
                ShowFailedPings = GetBooleanIniFile("Options", "ShowFailedPings", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Ping info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "ArtworkHSize", "") = "" Then
                WriteIntegerIniFile("Options", "ArtworkHSize", cArtworkSize)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile writing ArtworkHSize with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "ArtworkVsize", "") = "" Then
                WriteIntegerIniFile("Options", "ArtworkVsize", cArtworkSize)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile writing ArtworkVsize with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            ArtworkHSize = GetIntegerIniFile("Options", "ArtworkHSize", cArtworkSize)
            ArtworkVSize = GetIntegerIniFile("Options", "ArtworkVsize", cArtworkSize)
        Catch ex As Exception
            Log("Error in ReadIniFile reading Artwork Size info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Dim TempIfaceName As String = ""
        Try
            TempIfaceName = GetStringIniFile("Options", "Plugin Name", "")
            If TempIfaceName = "" Then
                WriteStringIniFile("Options", "Plugin Name", sIFACE_NAME)
            Else
                tIFACE_NAME = TempIfaceName
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Plugin Name with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("ListIndexes", "Include Learned Radiostations", "") = "" Then
                WriteBooleanIniFile("ListIndexes", "Include Learned Radiostations", True)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Include Learned Radiostations flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AnnouncementTitle", "") = "" Then
                WriteStringIniFile("Options", "AnnouncementTitle", "HomeSeer Announcement")
            Else
                AnnouncementTitle = GetStringIniFile("Options", "AnnouncementTitle", "")
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AnnouncementTitle info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AnnouncementAuthor", "") = "" Then
                WriteStringIniFile("Options", "AnnouncementAuthor", "Dirk Corsus")
            Else
                AnnouncementAuthor = GetStringIniFile("Options", "AnnouncementAuthor", "")
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AnnouncementAuthor info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AnnouncementAlbum", "") = "" Then
                WriteStringIniFile("Options", "AnnouncementAlbum", "SonosController")
            Else
                AnnouncementAlbum = GetStringIniFile("Options", "AnnouncementAlbum", "")
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AnnouncementAlbum info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "MaxAnnouncementTime", "") = "" Then
                WriteIntegerIniFile("Options", "MaxAnnouncementTime", MyMaxAnnouncementTime)
            Else
                MyMaxAnnouncementTime = GetIntegerIniFile("Options", "MaxAnnouncementTime", 100)
                MyAnnouncementCountdown = MyMaxAnnouncementTime
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading Maximum Announcement Time with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AnnouncementWaitforPlayTime", "") = "" Then
                WriteIntegerIniFile("Options", "AnnouncementWaitforPlayTime", MyAnnouncementWaitToSendPlay)
            Else
                MyAnnouncementWaitToSendPlay = GetIntegerIniFile("Options", "AnnouncementWaitforPlayTime", 2)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AnnouncementWaitforPlayTime with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "AnnouncementWaitBetweenPlayers", "") = "" Then
                WriteIntegerIniFile("Options", "AnnouncementWaitBetweenPlayers", MyAnnouncementWaitBetweenPlayers)
            Else
                MyAnnouncementWaitBetweenPlayers = GetIntegerIniFile("Options", "AnnouncementWaitBetweenPlayers", 0)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading AnnouncementWaitforPlayTime with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "NoPinging", "") <> "" Then
                MyNoPingingFlag = GetBooleanIniFile("Options", "NoPinging", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading NoPinging flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' convert the TTSSpeakDevice to the NewTTSSpeakDevice
        Try
            Dim TTSSpeakDeviceInfo As New System.Collections.Generic.Dictionary(Of String, String)()
            TTSSpeakDeviceInfo = GetIniSection("TTSSpeakDevice")
            If Not TTSSpeakDeviceInfo Is Nothing Then
                DeleteIniSection("TTSSpeakDevice")
                For Each KeyValue In TTSSpeakDeviceInfo
                    WriteStringIniFile("NewTTSSpeakDevice", KeyValue.Value, KeyValue.Key)
                Next
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile converting TTSSpeakDevice info with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "PostAnnouncementAction", "") = "" Then
                WriteIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
            End If
            MyPostAnnouncementAction = GetIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
        Catch ex As Exception
            Log("Error in ReadIniFile reading PostAnnouncementAction with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "HSSTrackLengthSetting", "") = "" Then
                WriteIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSHoursMinutesSeconds)
            End If
            MyHSTrackLengthFormat = GetIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSHoursMinutesSeconds)
        Catch ex As Exception
            Log("Error in ReadIniFile reading HS Device Tracklengh Format Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "HSSTrackPositionSetting", "") = "" Then
                WriteIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSHoursMinutesSeconds)
            End If
            MyHSTrackPositionFormat = GetIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSHoursMinutesSeconds)
        Catch ex As Exception
            Log("Error in ReadIniFile reading HS Device Trackposition Format Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "UPnPSubscribeTimeOut", "") = "" Then
                WriteIntegerIniFile("Options", "UPnPSubscribeTimeOut", UPnPSubscribeTimeOut)
            End If
            UPnPSubscribeTimeOut = GetIntegerIniFile("Options", "UPnPSubscribeTimeOut", UPnPSubscribeTimeOut)
        Catch ex As Exception
            Log("Error in ReadIniFile reading UPnPSubscribeTimeOut Setting with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If GetStringIniFile("Options", "MediaAPIEnabled", "") = "" Then
                WriteBooleanIniFile("Options", "MediaAPIEnabled", False)
                MediaAPIEnabled = False
            Else
                MediaAPIEnabled = GetBooleanIniFile("Options", "MediaAPIEnabled", False)
            End If
        Catch ex As Exception
            Log("Error in ReadIniFile reading MediaAPIEnabled flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub SetSpeakerProxy()
        If Not MyPIisInitialized Then Exit Sub
        Try
            If GetStringIniFile("SpeakerProxy", "Active", "") = "" Then
                WriteBooleanIniFile("SpeakerProxy", "Active", True)
                ActAsSpeakerProxy = True
            Else
                ActAsSpeakerProxy = GetBooleanIniFile("SpeakerProxy", "Active", False)
            End If
        Catch ex As Exception
            Log("Error in SetSpeakerProxy reading Speakerproxy flag with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If ProxySpeakerActive And Not ActAsSpeakerProxy Then
                callback.UnRegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                Log("Deactivated SpeakerProxy", LogType.LOG_TYPE_INFO)
                ProxySpeakerActive = False
            ElseIf Not ProxySpeakerActive And ActAsSpeakerProxy Then
                callback.RegisterProxySpeakPlug(sIFACE_NAME, MainInstance)
                Log("Registered SpeakerProxy", LogType.LOG_TYPE_INFO)
                ProxySpeakerActive = True
            End If
        Catch ex As Exception
            Log("Error in SetSpeakerProxy registering/unregistering Speakerproxy with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function ConvertToZoneName(ByVal Zone As String) As String
        ConvertToZoneName = GetStringIniFile("UPnP Devices UDN to Info", Zone, "")
    End Function

    Public Function ConvertToZoneUDN(ByVal Zone As String) As String
        ConvertToZoneUDN = Zone
        Try
            If Mid(Zone, 1, 5) = "uuid:" Then
                ' this is already in the format of UDN so do nothing
                ConvertToZoneUDN = Zone
            Else
                '  convert
                Dim DeviceUDNS As New System.Collections.Generic.Dictionary(Of String, String)()
                DeviceUDNS = GetIniSection("UPnP Devices UDN to Info")
                If DeviceUDNS Is Nothing Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ConvertToZoneUDN. Could not convert Zone = " & Zone, LogType.LOG_TYPE_WARNING)
                    Exit Function
                End If
                For Each DeviceUDN In DeviceUDNS
                    If DeviceUDN.Value.ToUpper = Zone.ToUpper Then
                        ConvertToZoneUDN = DeviceUDN.Key
                        Exit Function
                    End If
                Next
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in ConvertToZoneUDN. Could not convert Zone = " & Zone, LogType.LOG_TYPE_WARNING)
            End If
        Catch ex As Exception
        End Try
    End Function


    Public Sub ReadLinkgroups()
        Dim DeviceRef As Integer = -1
        DeviceRef = GetIntegerIniFile("Settings", "MasterHSDeviceRef", -1)
        If DeviceRef = -1 Then
            ' shouldn't be
            Exit Sub
        End If
        Try
            BuildButtonStringRef(DeviceRef)
        Catch ex As Exception
            Log("Error in ReadLinkgroups with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub InitializeSonosDevices()

        Dim dv As Scheduler.Classes.DeviceClass
        Dim ZoneInfo As String(,) = Nothing
        Dim SonosPlayer As HSPI
        Dim NewStart As Boolean = GetBooleanIniFile("Options", "RefreshDevices", False)
        TCPListenerPort = GetIntegerIniFile("Options", "TCPListenerPort", 0)

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Initializing Sonos Devices", LogType.LOG_TYPE_INFO)
        '
        ' Let's set the initial HomeSeer options for our plug-in
        '

        Try
            MasterHSDeviceRef = -1

            MasterHSDeviceRef = GetIntegerIniFile("Settings", "MasterHSDeviceRef", -1)
            If MasterHSDeviceRef <> -1 Then ' we already have a a masterHS device
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeSonosDevices found MasterHSDeviceRef in the inifile. MasterHSDeviceRef = " & MasterHSDeviceRef.ToString, LogType.LOG_TYPE_INFO)
            Else
                NewStart = True
                Log("InitializeSonosDevices is deleting all existing HS devices", LogType.LOG_TYPE_WARNING)
                hs.DeleteIODevices(sIFACE_NAME, MainInstance)
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
                MasterHSDeviceRef = hs.NewDeviceRef(SonosHSDevices.Master.ToString)
                Log("InitializeSonosDevices is creating new MasterDevice with Ref = " & MasterHSDeviceRef, LogType.LOG_TYPE_INFO)
                If MasterHSDeviceRef = -1 Then   ' checks if valid ref
                    Log("Error in InitializeSonosDevices. No More House Codes Available", LogType.LOG_TYPE_ERROR)
                    Exit Sub
                Else
                    Try
                        WriteIntegerIniFile("Settings", "MasterHSDeviceRef", MasterHSDeviceRef) ' saves our new base housecode
                    Catch ex As Exception
                        Log("Error in InitializeSonosDevices while writing to the ini file : " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Sub
                    End Try
                End If
                ' Force HomeSeer to save changes to devices and events so we can find our new device
                hs.SaveEventsDevices()
            End If

            Try
                dv = hs.GetDeviceByRef(MasterHSDeviceRef)
                dv.Interface(hs) = sIFACE_NAME
                If NewStart Then
                    dv.Location(hs) = "Master"
                    dv.Location2(hs) = tIFACE_NAME
                    dv.Device_Type_String(hs) = SonosHSDevices.Master.ToString
                    dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
                    dv.Image(hs) = ImagesPath & "Sonos.jpg"
                    dv.ImageLarge(hs) = ImagesPath & "Sonos.jpg"
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeSonosDevices added image  " & ImagesPath & "sonos.jpg", LogType.LOG_TYPE_INFO)
                    Dim DT As New DeviceTypeInfo With {
                        .Device_API = DeviceTypeInfo.eDeviceAPI.Media,
                        .Device_Type = DeviceTypeInfo.eDeviceType_Media.Root,
                        .Device_SubType_Description = "Sonos PlugIn Master Controller"
                    }
                    dv.DeviceType_Set(hs) = DT
                    dv.Status_Support(hs) = True
                    dv.Address(hs) = "Master"
                    dv.Relationship(hs) = Enums.eRelationship.Standalone
                    dv.MISC_Set(hs, Enums.dvMISC.CONTROL_POPUP)
                    dv.InterfaceInstance(hs) = MainInstance
                    hs.SaveEventsDevices()
                End If

                BuildButtonStringRef(MasterHSDeviceRef)
                hs.SetDeviceValueByRef(MasterHSDeviceRef, msInitializing, True)

            Catch ex As Exception
                Log("Error in InitializeSonosDevices creating the UPNP Master with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Try
                BuildSonosInfoList()       'This procedure gets all the players out of the HS Database, look them up in the .ini file and puts them in the SonosplayInfo array
            Catch ex As Exception
                Log("Error in InitializeSonosDevices building the Sonos Info List with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                If MySSDPDevice Is Nothing Then
                    MySSDPDevice = New MySSDP
                    If MySSDPDevice IsNot Nothing Then
                        UPnPViewerPage = New UPnPDebugWindow("SonosViewer")
                        If UPnPViewerPage IsNot Nothing Then
                            UPnPViewerPage.RefToSSDPn = MySSDPDevice
                        End If
                    End If
                End If
            Catch ex As Exception
                Log("Error in InitializeSonosDevices creating the SSDP Device with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            Try
                BuildHSSonosDevices(NewStart) ' parameter = refresh = true for periodic check to see if new devices are discovered
            Catch ex As Exception
                Log("Error in InitializeSonosDevices building the Sonos Devices List with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

        Catch ex As Exception
            Log("Error in InitializeSonosDevices: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        '
        ' Now check which ZonePlayers are not found

        CreateSonosControllers()
        'CheckForDuplicateZoneNames()

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("InitializeSonosDevices: Done Initializing Sonos Devices", LogType.LOG_TYPE_INFO)
        Try
            WriteIntegerIniFile("Options", "PreviousVersion", CurrentVersion)
        Catch ex As Exception
            Log("Error in InitializeSonosDevices writing PreviousVersion with error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        WriteBooleanIniFile("Options", "RefreshDevices", False) ' just in case this flag was set
        SetDeviceStringConnected()
        SonosPlayer = Nothing

    End Sub

    Public Sub DeleteDevice(DeviceUDN As String)
        Log("DeleteDevice called with DeviceUDN = " & DeviceUDN.ToString, LogType.LOG_TYPE_INFO)

        'BackUpIniFile(hs.GetAppPAth & gINIFile)
        If GetBooleanIniFile(DeviceUDN, DeviceInfoIndex.diDeviceIsAdded.ToString, False) Then
            DMRemove(DeviceUDN)
        End If
        ' [UPnP Devices UDN to Info] Key is UDN
        DeleteEntryIniFile("UPnP Devices UDN to Info", DeviceUDN)
        Dim KeyValues As New Dictionary(Of String, String)()
        ' [UPnP HSRef to UDN] scan through the section and remove the entries with value = UDN and call HS to delete the device!
        KeyValues = GetIniSection("UPnP HSRef to UDN")
        For Each Entry In KeyValues
            If Entry.Value = DeviceUDN Then
                If Entry.Key <> -1 Then
                    Log("DeleteDevice deleted HS Device Ref = " & Entry.Key.ToString, LogType.LOG_TYPE_INFO)
                    hs.DeleteDevice(Entry.Key)
                End If
                DeleteEntryIniFile("UPnP HSRef to UDN", Entry.Key)
            End If
        Next

        ' [UPnP UDN to HSRef]
        DeleteEntryIniFile("UPnP UDN to HSRef", DeviceUDN)

        ' Whole section [UDN]
        DeleteIniSection(DeviceUDN)
        ' Store only existing names 
        WriteSonosNamesToIniFile()
        hs.SaveEventsDevices()

    End Sub


    Public Sub DoRediscover()

        If piDebuglevel > DebugLevel.dlEvents Then Log("DoRediscover called", LogType.LOG_TYPE_INFO)
        'MySSDPDevice.SendMSearch()
        Try
            Dim AllDevices As MyUPnPDevices = MySSDPDevice.GetAllDevices()
            If Not AllDevices Is Nothing And AllDevices.Count > 0 Then
                For Each DLNADevice As MyUPnPDevice In AllDevices
                    If piDebuglevel > DebugLevel.dlEvents Then Log("DoRediscover found UDN = " & DLNADevice.UniqueDeviceName & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_INFO) ' moved this here on 11/16/2019
                    If (DLNADevice.UniqueDeviceName <> "") And (DLNADevice.Location <> "") And DLNADevice.Alive Then
                        ' check whether this devices was known to us and on-line
                        ' go find it in the array
                        Dim DLNADeviceInfo As MyUPnpDeviceInfo = Nothing
                        Dim NewUDN As String = Replace(DLNADevice.UniqueDeviceName, "uuid:", "")
                        DLNADeviceInfo = FindUPnPDeviceInfo(NewUDN)
                        If DLNADeviceInfo Is Nothing Then
                            If (DLNADevice.ModelNumber <> "ZB100" And DLNADevice.ModelNumber <> "BR200") Then
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover found New UDN = " & NewUDN & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                NewDeviceFound(NewUDN)
                            End If
                        Else
                            Dim SonosPlayer As HSPI = GetSonosPlayerByUDN(NewUDN)
                            If SonosPlayer Is Nothing Then
                                ' this should really not be
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover shouldn't have found UDN = " & NewUDN & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                NewDeviceFound(NewUDN)
                            Else
                                If SonosPlayer.DeviceStatus.ToUpper <> "ONLINE" Then
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoRediscover found Known UDN = " & NewUDN & ", with location = " & DLNADevice.Location & " and Alive = " & DLNADevice.Alive.ToString, LogType.LOG_TYPE_WARNING)
                                    NewDeviceFound(NewUDN)
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in DoRediscover with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub DoCheckChange()
        SonosSettingsHaveChanged = GetBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", False)
        If Not SonosSettingsHaveChanged Then Exit Sub
        If Not AutoUpdate Then Exit Sub
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckChange Called", LogType.LOG_TYPE_INFO)
        Dim ImmediateUpdate As Boolean = GetBooleanIniFile("Options", "Immediate Auto Update", False)
        If AutoUpdateTime = "" And Not ImmediateUpdate Then Exit Sub
        Dim Times
        Try
            Dim Now As DateTime = DateTime.Now
            Dim NowinMinutes As Integer
            Dim AutoUpdateInMinutes As Integer
            Times = Split(AutoUpdateTime, ":")
            Try
                AutoUpdateInMinutes = Times(0) * 60 + Times(1)
            Catch ex As Exception
                Log("Error in DoCheckChange. Time in Ini file is of wrong format. Auto Update Time = " & AutoUpdateTime, LogType.LOG_TYPE_ERROR)
            End Try
            NowinMinutes = Now.Hour * 60 + Now.Minute
            'Log( Now.ToShortTimeString) ' comes in the form of 7:55 AM
            'Log( Now.AddSeconds(TOCheckChangeValue).ToShortTimeString)
            'Log( Now.Hour)
            'Log( Now.Minute)
            If (AutoUpdateInMinutes >= (24 * 60)) Or (NowinMinutes >= AutoUpdateInMinutes And NowinMinutes <= AutoUpdateInMinutes + Math.Round(TOCheckChangeValue * 10 / 60)) Or ImmediateUpdate Then
                ' if a time is enterred that is bigger then 24h, immediately start creating DB
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Update Time Reached", LogType.LOG_TYPE_INFO)
                'SonosSettingsHaveChanged = False ' Hopefully this will prevent some reentry
                WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", False)
                If MusicDBIsBeingEstablished Then Exit Sub ' we have retry right here
                Dim ZonePlayer As HSPI
                Try
                    'If AutoUpdateZoneName = "" Then
                    If DBZoneName = "" Then
                        ' pick first zone
                        ZonePlayer = GetSonosPlayerByUDN("")
                    Else
                        'ZonePlayer = GetSonosPlayerByUDN(AutoUpdateZoneName)
                        ZonePlayer = GetSonosPlayerByUDN(DBZoneName)
                    End If
                    If Not ZonePlayer Is Nothing Then
                        ZonePlayer.BuildTrackDatabase(CurrentAppPath & MusicDBPath)
                    End If
                Catch ex As Exception
                    Log("Error in DoCheckChange getting ZonePlayer with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                ZonePlayer = Nothing
            End If
        Catch ex As Exception
            Log("Error in DoCheckChange with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub CreateWebLink(ByVal inZoneName As String, ByVal inZoneUDN As String)
        If Mid(inZoneUDN, 1, 5) = "uuid:" Then
            Mid(inZoneUDN, 1, 5) = "     "
            inZoneUDN = Trim(inZoneUDN)
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CreateWebLink called with ZoneUDN = " & inZoneUDN & " and PageName = " & PlayerControlPage & ":" & inZoneUDN, LogType.LOG_TYPE_INFO)
        Try
            ' Private TestPlayerConfigPage As PlayerControl   ' a jquery web page
            If MainInstance <> "" Then
                MyPlayerControlWebPage = New PlayerControl(PlayerControlPage & ":" & MainInstance & "-" & inZoneUDN, "0")
            Else
                MyPlayerControlWebPage = New PlayerControl(PlayerControlPage & ":" & inZoneUDN, "0")
            End If

            MyPlayerControlWebPage.RefToPlugIn = MyHSPIControllerRef 'Me
            MyPlayerControlWebPage.ZoneUDN = inZoneUDN

            ' register the page with the HS web server, HS will post back to the WebPage class
            ' "pluginpage" is the URL to access this page
            ' comment this out if you are going to use the GenPage/PutPage API istead
            If MainInstance <> "" Then
                hs.RegisterPage(PlayerControlPage, sIFACE_NAME, MainInstance & "-" & inZoneUDN)
            Else
                hs.RegisterPage(PlayerControlPage, sIFACE_NAME, inZoneUDN)
            End If


            ' register a configuration link that will appear on the interfaces page
            Dim wpd As New WebPageDesc
            ' register a normal page to appear in the HomeSeer menu
            wpd = New WebPageDesc With {
                .link = PlayerControlPage,
                .page_title = "Sonos " & inZoneName & " Config",
                .plugInName = sIFACE_NAME
            }
            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & inZoneName
                wpd.plugInInstance = MainInstance & "-" & inZoneUDN & "&clientid=0"
            Else
                wpd.linktext = inZoneName
                wpd.plugInInstance = inZoneUDN & "&clientid=0"
            End If
            hs.RegisterLinkEx(wpd)
        Catch ex As Exception
            Log("Error in CreateWebLink , unable to register the link(ex) with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeleteWebLink(inZoneUDN As String, inZoneName As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteWebLink called with ZoneUDN = " & inZoneUDN & ", ZoneName = " & inZoneName, LogType.LOG_TYPE_INFO)
        Try
            If MyPlayerControlWebPage IsNot Nothing Then
                Dim wpd As New WebPageDesc
                ' register a normal page to appear in the HomeSeer menu
                wpd = New WebPageDesc With {
                    .link = PlayerControlPage,
                    .page_title = "Sonos " & inZoneName & " Config",
                    .plugInName = sIFACE_NAME
                }
                If MainInstance <> "" Then
                    wpd.linktext = "Instance " & MainInstance & " " & inZoneName
                    wpd.plugInInstance = MainInstance & "-" & inZoneUDN & "&clientid=0"
                Else
                    wpd.linktext = inZoneName
                    wpd.plugInInstance = inZoneUDN & "&clientid=0"
                End If
                hs.UnRegisterLinkEx(wpd)
                'If MainInstance <> "" Then
                'hs.UnRegisterLinkEx(sIFACE_NAME, MainInstance & "-" & ZoneUDN)
                'Else
                'hs.UnRegisterLinkEx(sIFACE_NAME, ZoneUDN)
                'End If
                MyPlayerControlWebPage.Dispose()
                MyPlayerControlWebPage = Nothing
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteWebLink for ZoneUDN = " & inZoneUDN & ", ZoneName = " & inZoneName & " removed weblink successfully", LogType.LOG_TYPE_INFO) ' added v3.1.025
            End If
        Catch ex As Exception
            Log("Error in DeleteWebLink for ZoneUDN = " & inZoneUDN & ". Unable to UnRegister the link(ex) with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ChangeWebLink(OldZoneName As String, ByVal NewZoneName As String, ByVal ZoneUDN As String)
        If Mid(ZoneUDN, 1, 5) = "uuid:" Then
            Mid(ZoneUDN, 1, 5) = "     "
            ZoneUDN = Trim(ZoneUDN)
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ChangeWebLink called with ZoneUDN = " & ZoneUDN & " and NewZoneName = " & NewZoneName, LogType.LOG_TYPE_INFO)
        Try

            ' register a configuration link that will appear on the interfaces page
            Dim wpd As New WebPageDesc
            ' register a normal page to appear in the HomeSeer menu
            wpd = New WebPageDesc With {
                .link = PlayerControlPage,
                .page_title = "Sonos " & NewZoneName & " Config",
                .plugInName = sIFACE_NAME
            }

            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & OldZoneName
                wpd.plugInInstance = MainInstance & "-" & ZoneUDN & "&clientid=0" 'instance
            Else
                wpd.linktext = NewZoneName
                wpd.plugInInstance = ZoneUDN & "&clientid=0" 'instance
            End If

            hs.UnRegisterLinkEx(wpd)

            If MainInstance <> "" Then
                wpd.linktext = "Instance " & MainInstance & " " & NewZoneName
                wpd.plugInInstance = MainInstance & "-" & ZoneUDN & "&clientid=0" 'instance
            Else
                wpd.linktext = NewZoneName
                wpd.plugInInstance = ZoneUDN & "&clientid=0" 'instance
            End If
            hs.RegisterLinkEx(wpd)
        Catch ex As Exception
            Log("Error in CreateWebLink , unable to register the link(ex) with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function ReplaceSpecialCharacters(ByVal InString) As String
        ReplaceSpecialCharacters = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.length
                If InString(InIndex) = " " Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "!" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "#" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "$" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "&" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "(" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ")" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "*" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "+" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "," Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "-" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "." Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "/" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ":" Then
                    Outstring = Outstring + "_"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
            Loop
        Catch ex As Exception
            Log("Error in ReplaceSpecialCharacters. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ReplaceSpecialCharacters = Outstring
    End Function


    Public Function ZoneNameChanged(ByVal OldZoneName As String, ByVal ZoneUDN As String, ByVal NewZoneName As String, OverRulePairing As Boolean) As Boolean
        ' need complete fixing. All HS devices needs to be checked against this UDN and then name needs to be changed
        ZoneNameChanged = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ZoneNameChanged called with OldZoneName = " & OldZoneName & " and NewZoneName = " & NewZoneName & " and OverrulePairing = " & OverRulePairing.ToString, LogType.LOG_TYPE_INFO)
        ' we need to update the HS device Names
        ' Update the SonosInfo
        ' Update the .ini file
        ' Update the Weblinks
        ' Update the events?
        Dim dv As Scheduler.Classes.DeviceClass
        Dim Ref As Integer

        If MyHSDeviceLinkedList.Count = 0 Then
            Log("Error in ZoneNameChanged called with OldZoneName = " & OldZoneName & " and NewZoneName = " & NewZoneName & ". No Zones were found", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If

        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZoneUDN = ZoneUDN Then
                    If HSDevice.ZoneName <> OldZoneName Then
                        ' this should not be let's print out a warning
                        Log("Warning in ZoneNameChanged where ZoneName in Array didn't match. Found = " & HSDevice.ZoneName & " supposed to = " & OldZoneName, LogType.LOG_TYPE_WARNING)
                    End If
                    If OverRulePairing Then
                        WriteStringIniFile("UPnP Devices UDN to Info", ZoneUDN, NewZoneName)
                        WriteStringIniFile(ZoneUDN, DeviceInfoIndex.diFriendlyName.ToString, NewZoneName)
                    ElseIf Not ZoneNamesHaveChanged(HSDevice, NewZoneName, HSDevice.ZoneModel) Then
                        Exit Function
                    End If
                    'If Not (OverRulePairing Or ZoneNamesHaveChanged(HSDevice, NewZoneName, HSDevice.ZoneModel)) Then Exit Function                    
                    Try
                        Ref = GetIntegerIniFile(ZoneUDN, DeviceInfoIndex.diPlayerHSRef.ToString, -1)
                        If Ref <> -1 Then
                            dv = hs.GetDeviceByRef(Ref)
                            dv.Location(hs) = NewZoneName
                            hs.SaveEventsDevices()
                        End If
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diTrackHSRef, "Track")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diNextTrackHSRef, "Next Track")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diArtistHSRef, "Artist")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diNextArtistHSRef, "Next Artist")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diAlbumHSRef, "Album")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diNextAlbumHSRef, "Next Album")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diArtHSRef, "Art")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diNextArtHSRef, "Next Art")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diPlayStateHSRef, "State")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diVolumeHSRef, "Volume")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diMuteHSRef, "Mute")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diLoudnessHSRef, "Loudness")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diBalanceHSRef, "Balance")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diTrackLengthHSRef, "Track Length")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diTrackPosHSRef, "Track Position")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diRadiostationNameHSRef, "Radiostation Name")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diRepeatHSRef, "Repeat")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diShuffleHSRef, "Shuffle")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diTrackDescrHSRef, "Track Desc")
                        UpdateIniFileNameChange(ZoneUDN, NewZoneName, DeviceInfoIndex.diGenreHSRef, "Genre")
                    Catch ex As Exception
                        Log("Error in ZoneNameChanged where HS Device cannot be found with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Dim SonosPlayer As HSPI = GetAPIByUDN(HSDevice.ZoneUDN)                 ' added 9/26/2019
                    If SonosPlayer IsNot Nothing Then SonosPlayer.ZoneName = NewZoneName    ' added 9/26/2019
                    HSDevice.ZoneName = NewZoneName
                    ZoneNameChanged = True
                    WriteSonosNamesToIniFile()
                Else
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "ZoneNameChanged was looking for " & ZoneUDN & " and found " & SonosPlayersInfo(Index).ZoneUDN)
                End If
            Next
        Catch ex As Exception
            Log("Error in ZoneNameChanged modifying info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ChangeWebLink(OldZoneName, NewZoneName, ZoneUDN)
    End Function

    Private Sub UpdateIniFileNameChange(ZoneUDN As String, NewZoneName As String, infoIndex As DeviceInfoIndex, DevType As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateIniFileNameChange called with ZoneUDN = " & ZoneUDN & " and NewZoneName = " & NewZoneName & " and infoIndex = " & infoIndex.ToString & " and DevType = " & DevType, LogType.LOG_TYPE_INFO)
        Dim dv As Scheduler.Classes.DeviceClass
        Dim Ref As Integer
        Ref = GetIntegerIniFile(ZoneUDN, infoIndex.ToString, -1)
        If Ref <> -1 Then
            Try
                dv = hs.GetDeviceByRef(Ref)
                Dim DevName As String = DevType 'NewZoneName & " - " & DevType
                dv.Name(hs) = DevName
                dv.Location(hs) = NewZoneName
                'dv.Address(hs) = "SPlayer_" & NewZoneName
                hs.SaveEventsDevices()
            Catch ex As Exception
                Log("Error in UpdateIniFileNameChange modifying info with ZoneUDN = " & ZoneUDN & " and NewZoneName = " & NewZoneName & " and infoIndex = " & infoIndex.ToString & " and DevType = " & DevType & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Sub

    Public Sub ZoneRoomIconChanged(ByVal ZoneUDN As String, ByVal NewRoomIcon As String)
        ' need complete fixing. All HS devices needs to be checked against this UDN and then name needs to be changed

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ZoneRoomIconChanged called with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon, LogType.LOG_TYPE_INFO)

        Dim dv As Scheduler.Classes.DeviceClass
        Dim Ref As Integer

        If MyHSDeviceLinkedList.Count = 0 Then
            Log("Error in ZoneRoomIconChanged called with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon & ". No Zones were found", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim RoomIconPath As String = ""

        If NewRoomIcon <> "" Then RoomIconPath = ImagesPath & NewRoomIcon & ".png"

        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZoneUDN = ZoneUDN Then
                    HSDevice.ZoneCurrentIcon = NewRoomIcon
                    WriteStringIniFile(ZoneUDN, DeviceInfoIndex.diRoomIcon.ToString, NewRoomIcon)
                    Try
                        Ref = GetIntegerIniFile(ZoneUDN, DeviceInfoIndex.diPlayerHSRef.ToString, -1)
                        If Ref <> -1 Then
                            dv = hs.GetDeviceByRef(Ref)
                            dv.Image(hs) = RoomIconPath
                            dv.ImageLarge(hs) = RoomIconPath
                            hs.SaveEventsDevices()
                        End If
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diTrackHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diNextTrackHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diArtistHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diNextArtistHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diAlbumHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diNextAlbumHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diArtHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diNextArtHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diPlayStateHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diVolumeHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diMuteHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diLoudnessHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diBalanceHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diTrackLengthHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diTrackPosHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diRadiostationNameHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diRepeatHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diShuffleHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diTrackDescrHSRef)
                        UpdateIniFileRoomIconChange(ZoneUDN, RoomIconPath, DeviceInfoIndex.diGenreHSRef)
                    Catch ex As Exception
                        Log("Error in ZoneRoomIconChanged with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            Next
        Catch ex As Exception
            Log("Error in ZoneRoomIconChanged modifying info with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub UpdateIniFileRoomIconChange(ZoneUDN As String, NewRoomIcon As String, infoIndex As DeviceInfoIndex)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateIniFileRoomIconChange called with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon & " and infoIndex = " & infoIndex.ToString, LogType.LOG_TYPE_INFO)
        Dim dv As Scheduler.Classes.DeviceClass
        Dim Ref As Integer
        Ref = GetIntegerIniFile(ZoneUDN, infoIndex.ToString, -1)
        If Ref <> -1 Then
            Try
                dv = hs.GetDeviceByRef(Ref)
                dv.Image(hs) = NewRoomIcon
                dv.ImageLarge(hs) = NewRoomIcon
                hs.SaveEventsDevices()
            Catch ex As Exception
                Log("Error in UpdateIniFileRoomIconChange modifying info with ZoneUDN = " & ZoneUDN & " and NewRoomIcon = " & NewRoomIcon & " and infoIndex = " & infoIndex.ToString & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Sub



    Public Sub SetDeviceStringConnected()
        hs.SetDeviceValueByRef(MasterHSDeviceRef, msConnected, True)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SetDeviceStringConnected called", LogType.LOG_TYPE_INFO)
    End Sub

#End Region

#Region "    Music Related Procedures    "


    ReadOnly Property LinkgroupState As Boolean
        Get
            LinkgroupState = MyLinkState
        End Get
    End Property

    Public Function GetLinkgroupZoneDestination(ByVal LinkgroupName As String) As String
        GetLinkgroupZoneDestination = GetStringIniFile("LinkgroupZoneDestination", LinkgroupName, "")
    End Function

    Public Sub SetLinkgroupZoneDestination(ByVal LinkgroupName As String, ByVal Value As String)
        WriteStringIniFile("LinkgroupZoneDestination", LinkgroupName, Value)
    End Sub

    Public Function GetLinkgroupZoneSource(ByVal LinkgroupName As String) As String
        GetLinkgroupZoneSource = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
    End Function

    Public Sub SetLinkgroupZoneSource(ByVal LinkgroupName As String, ByVal Value As String)
        WriteStringIniFile("LinkgroupZoneSource", LinkgroupName, Value)
    End Sub



    Private Sub MyControllerTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyControllerTimer.Elapsed
        If InitDeviceFlag Then
            InitDeviceFlag = False
            InitializeSonosDevices()
            MyPIisInitialized = True
            SetSpeakerProxy()
        Else
            Dim Index As Integer
            For Index = 0 To MaxTOActionArray
                If MyTimeoutActionArray(Index) <> 0 Then HandleTimeout(Index)
            Next
            If CapabilitiesCalledFlag Then
                CapabilitiesCalledFlag = False
                SendEventForAllZones()
            End If
        End If
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub MyAnnouncementTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyAnnouncementTimer.Elapsed
        MyAnnouncementTimer.Stop()
        DoCheckAnnouncementQueue()
        e = Nothing
        sender = Nothing
        MyAnnouncementTimer.Start()
    End Sub

    Private Sub MyDBCreationTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyDBCreationTimer.Elapsed
        Try
            Dim SonosPlayer As HSPI = Nothing
            If DBZoneName = "" Then
                ' pick first zone
                SonosPlayer = GetSonosPlayerByUDN("")
                If SonosPlayer.ZoneModel = "WD100" Then
                    ' we can't use this to create databases
                    Log("Error in MyDBCreationTimer_Elapsed. BuildDB only works on none WD100 device. Your first device is a WD100", LogType.LOG_TYPE_ERROR)
                    Log("Error in MyDBCreationTimer_Elapsed. Go to the .ini file and specify which Zoneplayer to use as source to build a database", LogType.LOG_TYPE_ERROR)
                    Log("Error in MyDBCreationTimer_Elapsed. Consult the help file for more info on specifying a source for database creation", LogType.LOG_TYPE_ERROR)
                    Exit Try
                End If
            Else
                SonosPlayer = GetSonosPlayerByUDN(DBZoneName)
            End If
            If Not SonosPlayer Is Nothing Then
                SonosPlayer.BuildTrackDatabase(CurrentAppPath & MusicDBPath)
            End If
        Catch ex As Exception
            Log("Error in MyDBCreationTimer_Elapsed getting ZonePlayer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub HandleTimeout(ByVal TOIndex As Integer)

        Select Case TOIndex
            Case TORediscover
                ' every 5 minutes I want to see if any zones got added
                MyTimeoutActionArray(TORediscover) = MyTimeoutActionArray(TORediscover) - 1
                If MyTimeoutActionArray(TORediscover) <= 0 Then
                    If MyPIisInitialized Then DoRediscover() ' changed in v3.1.0.27 because discovery kicked in while still initializing, Changed again in v.50 on 11/16/2019 because using initdeviceflag is wrong
                    MyTimeoutActionArray(TORediscover) = TORediscoverValue
                End If
            Case TOCheckChange
                MyTimeoutActionArray(TOCheckChange) = MyTimeoutActionArray(TOCheckChange) - 1
                If MyTimeoutActionArray(TOCheckChange) <= 0 Then
                    DoCheckChange()
                    MyTimeoutActionArray(TOCheckChange) = TOCheckChangeValue
                End If
            Case TOAddNewDevice
                If MyPIisInitialized Then
                    MyTimeoutActionArray(TOAddNewDevice) = MyTimeoutActionArray(TOAddNewDevice) - 1
                    If MyTimeoutActionArray(TOAddNewDevice) <= 0 Then
                        Try
                            AddNewDiscoveredDevice()
                        Catch ex As Exception
                            Log("Error in HandleTimeout. Unable to add a new device with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                Else
                    ' reset the timer
                    MyTimeoutActionArray(TOAddNewDevice) = TOAddNewDeviceValue
                End If
        End Select
    End Sub


    Public Sub SaveAllPlayersState() ' This procedure will save the state of all on-line players
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveAllPlayersState called", LogType.LOG_TYPE_INFO)
        ' [Sonos Zonenames]
        'Names=Master Bedroom;:;ZP90:|:Patio;:;ZP120:|:Kitchen;:;ZP120:|:Family Room;:;ZP90:|:Wireless Dock;:;WD100:|:Office;:;S5:|:Office2;:;S5
        'UDNs=uuid:RINCON_000E5825227A01400:|:uuid:RINCON_000E5832D2D401400:|:uuid:RINCON_000E5833F3CC01400:|:uuid:RINCON_000E5824C3B001400:|:uuid:RINCON_000E5860905A01400:|:uuid:RINCON_000E5858C97A01400
        Dim AllZoneNameString As String
        Dim AllZoneNames() As String
        AllZoneNameString = GetStringIniFile("Sonos Zonenames", "UDNs", "")
        Dim SonosPlayerToBeSaved As HSPI
        If AllZoneNameString = "" Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SaveAllPlayersState didn't find the UDN string in the [Sonos Zonenames] section in the .ini file", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        AllZoneNames = Split(AllZoneNameString, ":|:")
        Try
            For Each ZoneToBeSaved In AllZoneNames
                SonosPlayerToBeSaved = GetAPIByUDN(ZoneToBeSaved)
                If SonosPlayerToBeSaved IsNot Nothing Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveAllPlayersState: Start Save Current Track Info for Zone = " & ZoneToBeSaved, LogType.LOG_TYPE_INFO)
                    SonosPlayerToBeSaved.SaveCurrentTrackInfo("SaveAllPlayersInternalGroup", True, True)
                End If
            Next
        Catch ex As Exception
            Log("Error in SaveAllPlayersState with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveAllPlayersState done", LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub RestoreAllPlayersState() ' This procedure will restore all on-line players to their previously stored state
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreAllPlayersState called", LogType.LOG_TYPE_INFO)
        ' [Sonos Zonenames]
        'Names=Master Bedroom;:;ZP90:|:Patio;:;ZP120:|:Kitchen;:;ZP120:|:Family Room;:;ZP90:|:Wireless Dock;:;WD100:|:Office;:;S5:|:Office2;:;S5
        'UDNs=uuid:RINCON_000E5825227A01400:|:uuid:RINCON_000E5832D2D401400:|:uuid:RINCON_000E5833F3CC01400:|:uuid:RINCON_000E5824C3B001400:|:uuid:RINCON_000E5860905A01400:|:uuid:RINCON_000E5858C97A01400
        Dim AllZoneNameString As String
        Dim AllZoneNames() As String
        AllZoneNameString = GetStringIniFile("Sonos Zonenames", "UDNs", "")
        Dim SonosPlayerToBeRestored As HSPI
        If AllZoneNameString = "" Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in RestoreAllPlayersState didn't find the UDN string in the [Sonos Zonenames] section in the .ini file", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Try
            AllZoneNames = Split(AllZoneNameString, ":|:")
            For Each ZoneToBeRestored In AllZoneNames
                SonosPlayerToBeRestored = GetAPIByUDN(ZoneToBeRestored)
                If SonosPlayerToBeRestored IsNot Nothing Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreAllPlayersState: Start Restore Track Info for Zone = " & ZoneToBeRestored, LogType.LOG_TYPE_INFO)
                    SonosPlayerToBeRestored.RestoreCurrentTrackInfo("SaveAllPlayersInternalGroup", False)
                End If
            Next
        Catch ex As Exception
            Log("Error in RestoreAllPlayersState with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreAllPlayersState done", LogType.LOG_TYPE_INFO)
    End Sub


    Public Function NumInstances() As Integer
        ' Returns the number of instances this plug-in supports. Plug-ins that support multiple instances probably support multiple output devices and one music library. 
        'Normally, plug-ins return 1 for this value.

        If Not gIOEnabled Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("NumInstances for instance = " & instance & " was called before end of initialization.", LogType.LOG_TYPE_WARNING)
            ' we're not intialized yet, this is a problem. 
            'wait(10)
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("NumInstances was called for instance = " & instance, LogType.LOG_TYPE_INFO)
        End If
        If instance <> "" Then Return 1 Else Return 0
        'NumInstances = NbrOfSonosPlayers
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("NumInstances called. Instances is " & NumInstances.ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetInstanceName(ByVal Instance As Integer) As String
        ' Returns the name of this instance as set in the plug-in configuration.
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetInstanceName called with value : " & Instance.ToString, LogType.LOG_TYPE_INFO)
        If Instance <> "" Then
            Return ZoneName
        Else
            Return ""
        End If


        GetInstanceName = ""
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetInstanceName called with value : " & Instance.ToString, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetInstanceName. ZoneInstance not found. Instance : " & Instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneDeviceAPIIndex = Instance Then
                ' special case, just find the first player with a reference
                GetInstanceName = HSDevice.ZoneName
                Exit Function
            End If
        Next
    End Function


    Public Function GetInstanceUDN(ByVal Instance As Integer) As String
        ' Returns the name of this instance as set in the plug-in configuration.
        GetInstanceUDN = ""
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetInstanceUDN called with value : " & Instance.ToString)
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetInstanceUDN. ZonePlayer not found. Instance : " & Instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneDeviceAPIIndex = Instance Then
                ' special case, just find the first player with a reference
                GetInstanceUDN = HSDevice.ZoneUDN
                Exit Function
            End If
        Next
    End Function

    Public Function GetInstanceByName(ByVal Instance As String) As Integer
        ' Returns the name of this instance as set in the plug-in configuration.
        ' this is either a zone name or a zone UDN in starting with uuid:RINCON
        GetInstanceByName = 0
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetInstanceByName called with value : " & Instance)
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetInstanceByName. ZonePlayer not found. Instance : " & Instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Instance = Trim(Instance)
        If Instance = "" Then
            Exit Function
        End If
        If Mid(Instance, 1, 5) = "uuid:" Then
            Mid(Instance, 1, 5) = "     "
            Instance = Trim(Instance)
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If (HSDevice.ZoneName = Instance) Or (HSDevice.ZoneUDN = Instance) Then
                GetInstanceByName = HSDevice.ZoneDeviceAPIIndex
                Exit Function
            End If
        Next
    End Function

    Public Function GetAPIByUDN(ByVal inUDN As String) As HSPI
        ' Returns the name of this instance as set in the plug-in configuration.
        ' this is either a zone name or a zone UDN in starting with uuid:RINCON
        GetAPIByUDN = Nothing
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetInstanceByName called with value : " & Instance)
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetAPIByUDN. Device not found. UDN : " & inUDN, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        inUDN = Trim(inUDN)
        If inUDN = "" Then
            Exit Function
        End If
        If Mid(inUDN, 1, 5) = "uuid:" Then
            Mid(inUDN, 1, 5) = "     "
            inUDN = Trim(inUDN)
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneUDN = inUDN Then
                GetAPIByUDN = HSDevice.ZonePlayerControllerRef
                Exit Function
            End If
        Next
    End Function


    Public Function GetZoneNamebyUDN(ByVal inUDN As String) As String
        GetZoneNamebyUDN = ""
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetZoneNamebyUDN. ZonePlayer not found. inUDN : " & inUDN, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        inUDN = Trim(inUDN)
        If inUDN = "" Then
            Exit Function
        End If
        If Mid(inUDN, 1, 5) = "uuid:" Then
            Mid(inUDN, 1, 5) = "     "
            inUDN = Trim(inUDN)
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneUDN = inUDN Then
                GetZoneNamebyUDN = HSDevice.ZoneName
                Exit Function
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning GetZoneNamebyUDN found no ZoneName for UDN = " & inUDN & " and has " & MyHSDeviceLinkedList.Count.ToString & " players in array", LogType.LOG_TYPE_WARNING)
    End Function

    Public Function GetUDNbyZoneName(ByVal inZoneName As String) As String
        GetUDNbyZoneName = ""
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetUDNbyZoneName called for for Zone Name = " & inZoneName, LogType.LOG_TYPE_INFO)

        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetUDNbyZoneName. ZonePlayer not found. inZoneName : " & inZoneName, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        inZoneName = Trim(inZoneName)
        If inZoneName = "" Then
            Exit Function
        End If
        ' Go find the zone where this belongs to
        If Mid(inZoneName, 1, 5) = "uuid:" Then
            ' this is already in the format of UDN so do nothing
            GetUDNbyZoneName = inZoneName
            Exit Function
        End If
        If Mid(inZoneName, 1, 7) = "RINCON_" Then
            ' this is already in the format of UDN so do nothing
            GetUDNbyZoneName = inZoneName
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneName = inZoneName Then
                GetUDNbyZoneName = HSDevice.ZoneUDN
                Exit Function
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning GetUDNbyZoneName found no UDN for Zone Name = " & inZoneName & " and has " & MyHSDeviceLinkedList.Count.ToString & " players in array", LogType.LOG_TYPE_WARNING)
    End Function

    Public Function GetMultiZoneAPI() As Object
        ' returns HSMultiZoneAPI class
        ' Returns a reference to the HSMultiZoneAPI object if supported. If the plug-in does not support this class then this call returns nothing.
        'If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("GetMultiZoneAPI was called for instance = " & instance, LogType.LOG_TYPE_INFO)
        GetMultiZoneAPI = Nothing
    End Function

    Public Function GetMusicAPI(ByVal instance As Integer) As HSPI
        ' return a reference to a music API given an instance number
        ' returns HSMusicAPI Class
        ' Returns a reference to the HSMusicAPI for the given instance number or instance name. This function is overloaded and may accept either the instance number or name. 
        'The returned object is a reference to the HSMusicAPI as defined in the next section.
        GetMusicAPI = Nothing
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetMusicAPI (Integer) was called with Value: " & instance.ToString, LogType.LOG_TYPE_INFO)
        'Log( "GetMusicAPI (Integer) was called with Value: " & instance.ToString)

        If instance = 0 Then Exit Function
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetMusicAPI. ZonePlayer not found. Integer Instance : " & instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneDeviceAPIIndex = instance Then
                ' special case, just find the first player with a reference
                GetMusicAPI = HSDevice.ZonePlayerControllerRef
                Exit Function
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning GetMusicAPI (Integer) was called with Value: " & instance.ToString & " but found no player", LogType.LOG_TYPE_WARNING)
    End Function

    Public Function GetMusicAPI(ByVal instance As String) As HSPI
        ' returns HSMusicAPI Class
        'Input = Integer = instance number or String = instance name representing either zone name or Zone UDN.
        ' If Zone UDN, then the string starts with uuid:RINCON
        ' Returns a reference to the HSMusicAPI for the given instance number or instance name. This function is overloaded and may accept either the instance number or name. 
        'The returned object is a reference to the HSMusicAPI as defined in the next section.
        GetMusicAPI = Nothing
        Dim SearchIndex As Integer = 0
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "GetMusicAPI (string) was called with Value: " & instance.ToString)
        'Log( "GetMusicAPI (string) was called with Value: " & instance.ToString)
        If instance = "" Then Exit Function
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetMusicAPI. ZonePlayer not found. String Instance : " & instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If

        SearchIndex = Val(instance.ToString)

        If SearchIndex = 0 Then
            ' this might be a zonename rather then an integer
            SearchIndex = GetInstanceByName(instance)
        End If
        GetMusicAPI = GetMusicAPI(SearchIndex)
    End Function

    Public Function GetHSDeviceReference(ByVal inUDN As String) As Integer
        GetHSDeviceReference = -1
        Dim SearchIndex As Integer = 0
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetHSDeviceReference was called with inUDN : " & inUDN.ToString, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetHSDeviceReference. ZonePlayer not found. Instance : " & instance, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        If Mid(inUDN, 1, 5) = "uuid:" Then
            Mid(inUDN, 1, 5) = "     "
            inUDN = Trim(inUDN)
        End If
        If inUDN = "" Then Exit Function
        For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If HSDevice.ZoneUDN = inUDN Or HSDevice.ZoneName = inUDN Then
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    GetHSDeviceReference = HSDevice.ZonePlayerControllerRef.GetHSDeviceRefPlayer()
                    If piDebuglevel > DebugLevel.dlEvents Then Log("GetHSDeviceReference called with Name: " & instance.ToString & " and Found DeviceReference = " & GetHSDeviceReference.ToString, LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
            End If
        Next
    End Function

    Public Function LinkAZone(ByVal TargetUDN As String, ByVal SourceUDN As String) As Boolean
        LinkAZone = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LinkAZone called with TargetUDN : " & TargetUDN & " and SourceUDN = " & SourceUDN, LogType.LOG_TYPE_INFO)
        ' UDN is in the format x-rincon:RINCON_000E5824C3B001400 or x-sonos-dock:RINCON_000E5860905A01400 in case of wireless dock
        If InStr(SourceUDN, "-sonos-dock:") <> 0 Then
            Try
                SourceUDN = Replace(SourceUDN, "x-sonos-dock:", "") ' "x-sonos-dock:" Remove the x-sonos-dock:
            Catch ex As Exception
                Log("Issue with UDN in LinkAZone. UDN = " & SourceUDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            Try
                SourceUDN = Replace(SourceUDN, "x-rincon:", "") ' "x-rincon:" Remove the x-rincon
            Catch ex As Exception
                Log("Issue with UDN in LinkAZone. UDN = " & SourceUDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        End If

        ' Go find the zone where this belongs to 
        Dim SonosPlayer As HSPI = Nothing
        SonosPlayer = MyHSPIControllerRef.GetAPIByUDN(SourceUDN)
        If SonosPlayer IsNot Nothing Then
            SonosPlayer.AddTargetLinkZone(TargetUDN)
            SonosPlayer = Nothing
            LinkAZone = True
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LinkAZone called with TargetUDN : " & TargetUDN & " and SourceUDN = " & SourceUDN & " but didn't find SourcePlayer. Request queued", LogType.LOG_TYPE_WARNING)
        End If
    End Function

    Public Sub UnlinkAZone(ByVal TargetUDN As String, ByVal SourceUDN As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UnlinkAZone called with TargetUDN : " & TargetUDN & " and SourceUDN = " & SourceUDN, LogType.LOG_TYPE_INFO)
        ' SourceUDN is in the format x-rincon:RINCON_000E5824C3B001400 or x-sonos-dock:RINCON_000E5860905A01400 in case of wireless dock
        ' TargetUDN is in the format uuid:RINCON_000E5824C3B001400
        If InStr(SourceUDN, "-sonos-dock:") <> 0 Then
            Try
                SourceUDN = Replace(SourceUDN, "x-sonos-dock:", "") ' "x-sonos-dock:" Remove the x-sonos-dock:
            Catch ex As Exception
                Log("Issue with UDN in UnlinkAZone. UDN = " & SourceUDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        Else
            Try
                SourceUDN = Replace(SourceUDN, "x-rincon:", "") ' "x-rincon:" Remove the x-rincon
            Catch ex As Exception
                Log("Issue with UDN in UnlinkAZone. UDN = " & SourceUDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        End If

        ' Go find the zone where this belongs to 

        Dim SonosPlayer As HSPI
        SonosPlayer = MyHSPIControllerRef.GetAPIByUDN(SourceUDN)
        If SonosPlayer IsNot Nothing Then
            SonosPlayer.RemoveTargetLinkZone(TargetUDN)
            SonosPlayer = Nothing
        End If
    End Sub

    Public Function GetAllActiveZones() As String()
        GetAllActiveZones = Nothing
        If MyHSDeviceLinkedList.Count = 0 Then
            Exit Function
        End If

        Try
            Dim ZoneUDNs As String() = Nothing
            Dim Index As Integer = 0
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZoneOnLine Then
                    ReDim Preserve ZoneUDNs(Index)
                    ZoneUDNs(Index) = HSDevice.ZoneUDN
                    Index += 1
                End If
            Next
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetAllActiveZones called and found " & Index.ToString & " Player UDNs", LogType.LOG_TYPE_INFO)
            GetAllActiveZones = ZoneUDNs
        Catch ex As Exception
            Log("Error in GetAllActiveZones with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function



    Private Sub SendEventForAllZones()
        ' 	generate some event from all players to get ipad/iphone clients updated when they come back on-line
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("SendEventForAllZones called", LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then
            Exit Sub
        End If
        Try
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.PlayChangeNotifyCallback(Player_status_change.PlayStatusChanged, HSDevice.ZonePlayerControllerRef.PlayerState, False)
                    HSDevice.ZonePlayerControllerRef.PlayChangeNotifyCallback(Player_status_change.SongChanged, Player_state_values.UpdateHSServerOnly, False)
                End If
            Next
        Catch ex As Exception
            Log("Error in SendEventForAllZones with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


#End Region

#Region "    Speaker Proxy Related Procedures    "

    Public Sub SpeakIn(ByVal device As Integer, ByVal text As String, ByVal wait As Boolean, ByVal host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SpeakIn called for Device = " & device.ToString & ", Text = " & text & ", Wait=" & wait.ToString & ", Host = " & host & ", PIInitialized = " & MyPIisInitialized.ToString, LogType.LOG_TYPE_INFO)
        If Not MyPIisInitialized Then
            If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, host)
            Exit Sub
        End If
        Dim HostDevice As String
        HostDevice = host
        HostDevice = Trim(HostDevice)
        Dim SpeakerClientList As String() = Nothing
        SpeakerClientList = Split(HostDevice, ",")
        If HostDevice <> "" Then
            For Each SpeakerClient As String In SpeakerClientList
                If Mid(SpeakerClient, 1, 7).ToUpper = "$SONOS$" Then
                    ' this is for the plug-in. The next "$" character ends the Linkgroup ZoneName
                    Mid(SpeakerClient, 1, 7) = "       "
                    SpeakerClient = Trim(SpeakerClient)
                    Dim LinkgroupName As String
                    Dim Delimiter As Integer
                    Delimiter = SpeakerClient.IndexOf("$")
                    If Delimiter = 0 Then
                        ' should not be, just pass it along
                        If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, SpeakerClient)
                        'Exit Sub
                    Else
                        LinkgroupName = SpeakerClient.Substring(0, Delimiter)
                        ' remove the LinkgroupName from the Host
                        SpeakerClient = SpeakerClient.Remove(0, Delimiter + 1)
                        SpeakerClient = Trim(SpeakerClient)
                        If SpeakerClient = ":*" Then SpeakerClient = "*:*"
                        'Log( "SpeakerProxy activated with HostName = " & SpeakerClient & " Text = " & text & " and LinkgroupName = " & LinkgroupName)
                        AddAnnouncementToQueue(LinkgroupName, device, text, SpeakerClient)
                        DoCheckAnnouncementQueue()
                        'Exit Sub
                    End If
                ElseIf Mid(SpeakerClient, 1, 4).ToUpper = "$MC$" Then
                    ' don't do anything
                    'Exit Sub
                Else
                    ' these could be Sonos devices. The .ini file should have the LinkgroupName under [TTSSpeakDevice]
                    If MyLinkState Then
                        ' we are currently linked for TTS and the flag to overwrite is set. A user used BLRandom and apparantly speech is called with Wait flag false and as a result
                        ' the TTS event doesn't work properly
                        wait = True
                    End If
                    If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, SpeakerClient)
                End If
            Next
        Else
            ' these could be Sonos devices. The .ini file should have the LinkgroupName under [TTSSpeakDevice]
            If MyLinkState Then
                ' we are currently linked for TTS and the flag to overwrite is set. A user used BLRandom and apparantly speech is called with Wait flag false and as a result
                ' the TTS event doesn't work properly
                wait = True
            End If
            Dim FoundOne As Boolean = False
            Try
                Dim NewTTSSpeakDeviceInfo As New System.Collections.Generic.Dictionary(Of String, String)()
                NewTTSSpeakDeviceInfo = GetIniSection("NewTTSSpeakDevice")
                If Not NewTTSSpeakDeviceInfo Is Nothing Then
                    For Each KeyValue In NewTTSSpeakDeviceInfo
                        If KeyValue.Value <> "" Then
                            ' there is a deviceID here
                            Dim DeviceIDs As String() = Split(KeyValue.Value, ",")
                            For Each DeviceID In DeviceIDs
                                If (DeviceID <> "") And (Val(DeviceID) = device) Then
                                    AddAnnouncementToQueue(KeyValue.Key, device, text, HostDevice)
                                    DoCheckAnnouncementQueue()
                                    FoundOne = True
                                End If
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                Log("Error in Processing the Speaker Device IDs. Most likely a non integer value. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If Not FoundOne Then
                If Not MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysDrop Then hs.SpeakProxy(device, text, wait, host)
                Exit Sub
            Else
                If MyPostAnnouncementAction = PostAnnouncementAction.paaAlwaysForward Then hs.SpeakProxy(device, text, wait, host)
                Exit Sub
            End If
        End If
    End Sub


    Private Sub HandleLinkEvents(ByVal ButtonName As String)
        ' This is how the ini file should be built
        ' [LinkgroupNames]
        ' Names=TTS|Party
        ' [LinkgroupZoneSource]
        ' TTS=Family Room
        ' Party=Master Bedroom
        ' [LinkgroupZoneDestination]
        ' Test=Master Bedroom|Patio

        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkEvents called with ButtonName = " & ButtonName, LogType.LOG_TYPE_INFO)
        Dim LinkgroupZones() As String
        Dim LinkgroupZone As String
        LinkgroupZone = GetStringIniFile("LinkgroupNames", "Names", "")
        If LinkgroupZone = "" Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in HandleLinkEvents didn't find the [LinkgroupNames] in the .ini file", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        LinkgroupZones = Split(LinkgroupZone, "|")
        For Each LinkgroupZone In LinkgroupZones
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkEvents: Found LinkgroupZoneSource = " & LinkgroupZone, LogType.LOG_TYPE_INFO)
            If ("Link-" & LinkgroupZone) = ButtonName Then
                MyLinkState = True
                'HandleLinking(LinkgroupZone, True)
                HandleLinkingOn(LinkgroupZone, True)  ' changed on 2/2/2021 in v3.1.0.56 Without the parameter "true" the event doesn't ungroup the source player when linking
                Exit Sub
            ElseIf ("Unlink-" & LinkgroupZone) = ButtonName Then
                'HandleLinking(LinkgroupZone, False)
                HandleLinkingOff(LinkgroupZone)
                MyLinkState = False
                Exit Sub
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in HandleLinkEvents. Couldn't find " & LinkgroupZone & "  in [LinkgroupNames] in the .ini file", LogType.LOG_TYPE_ERROR)
    End Sub

    Public Function HandleLinkingOn(ByVal LinkgroupName As String, Optional ByVal IsFile As Boolean = False) As Integer ' returns a delay in seconds
        HandleLinkingOn = 0
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim NbrOfLinkgroupZones As Integer = 0
        Dim LinkgroupZoneSource As String
        Dim MaxWaitForVolumeToAdjust As Integer = 0
        Dim MySavedPlayerList As SavedPlayerList
        Dim MyLocalLinkgroupArray As LinkGroupInfo
        MyLocalLinkgroupArray = MyLinkgroupArray.GetLinkGroupInfo(LinkgroupName)
        MySavedPlayerList = MyLocalLinkgroupArray.MySavedPlayerList
        LinkgroupZoneSource = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
        If LinkgroupZoneSource = "" Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in HandleLinkingOn didn't find " & LinkgroupName & " under [LinkgroupZoneSource] in the .ini file", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Dim LinkgroupZoneSourceDetails() As String
        Dim UseLinkGroupAsTTS As Boolean = False ' changed in v3.0.0.20 from true to false
        LinkgroupZoneSourceDetails = Split(LinkgroupZoneSource, ";")
        If UBound(LinkgroupZoneSourceDetails) > 0 Then
            UseLinkGroupAsTTS = CBool(LinkgroupZoneSourceDetails(1))  ' changed in v3.0.0.20 and removed the Not Cbool
            LinkgroupZoneSource = LinkgroupZoneSourceDetails(0)
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn found " & LinkgroupZoneSource & " as source and Audio Input = " & UseLinkGroupAsTTS.ToString & " and Save Source = " & IsFile.ToString, LogType.LOG_TYPE_INFO)
        Dim SourceSonosPlayer As HSPI
        Dim DestinationSonosPlayer As HSPI
        Try
            SourceSonosPlayer = GetAPIByUDN(LinkgroupZoneSource)
            If SourceSonosPlayer Is Nothing Then
                Log("Error in HandleLinkingOn: Controller for LinkgroupZoneSource = " & LinkgroupZoneSource & " not found", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in HandleLinkingOn: Controller for LinkgroupZoneSource = " & LinkgroupZoneSource & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If IsFile Then
            ' we need to save the Source as well
            If SourceSonosPlayer.ZoneIsASlave Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn found source player to be a slave and switched to the master", LogType.LOG_TYPE_INFO) ' added 7/12/2019 in v3.1.0.31
                SourceSonosPlayer = MyHSPIControllerRef.GetAPIByUDN(SourceSonosPlayer.ZoneMasterUDN)
                If SourceSonosPlayer Is Nothing Then    ' added 7/25/2019 in v3.1.0.36
                    Log("Error in HandleLinkingOn: Controller for LinkgroupZoneSource = " & LinkgroupZoneSource & " after switching to master not found", LogType.LOG_TYPE_ERROR)
                    Exit Function
                End If
            End If
            SourceSonosPlayer.SaveCurrentTrackInfo(LinkgroupName, True, False)
            ' check whether the source is part of a linkgroup
            SaveLinkedPlayers(SourceSonosPlayer.GetUDN, SourceSonosPlayer.GetUDN, LinkgroupName)
            'SourceSonosPlayer.StopPlay() ' stop it
            MySavedPlayerList.Add(SourceSonosPlayer.GetUDN)

        End If
        Dim LinkgroupZoneDestination As String
        Dim LinkgroupZoneDestinations() As String
        Dim LinkgroupZoneDetail() As String
        LinkgroupZoneDestination = GetStringIniFile("LinkgroupZoneDestination", LinkgroupName, "")
        If LinkgroupZoneDestination = "" Then
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in HandleLinkingOn didn't find " & LinkgroupName & " under [LinkgroupZoneDestination] in the .ini file", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        LinkgroupZoneDestinations = Split(LinkgroupZoneDestination, "|")

        ' save all zones before changing anything to their state
        For Each LinkgroupZone In LinkgroupZoneDestinations
            LinkgroupZoneDetail = Split(LinkgroupZone, ";")
            If UBound(LinkgroupZoneDetail) > 2 Then
                If LinkgroupZoneDetail(3) = "1" Then
                    ' this is an active zone
                    DestinationSonosPlayer = GetAPIByUDN(LinkgroupZoneDetail(0))
                    If DestinationSonosPlayer Is Nothing Then
                        Log("Error in HandleLinkingOn for LinkgroupZoneDestination = " & LinkgroupZone & " not found while saving zones", LogType.LOG_TYPE_ERROR)
                    Else
                        If DestinationSonosPlayer.ZoneIsASlave Then
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn detected PairedSlave = " & DestinationSonosPlayer.GetZoneName & " and UDN = " & DestinationSonosPlayer.GetUDN & " and will save the Master Only", LogType.LOG_TYPE_INFO)
                            DestinationSonosPlayer = MyHSPIControllerRef.GetAPIByUDN(DestinationSonosPlayer.ZoneMasterUDN)
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn detected PairedSlave and switched to Master = " & DestinationSonosPlayer.GetZoneName & " and UDN = " & DestinationSonosPlayer.GetUDN, LogType.LOG_TYPE_INFO)
                        End If
                        Try
                            SaveLinkedPlayers(DestinationSonosPlayer.GetUDN, SourceSonosPlayer.GetUDN, LinkgroupName)
                            If Not (IsFile And (SourceSonosPlayer.GetUDN = DestinationSonosPlayer.GetUDN)) Then
                                'If Not (IsFile And (LinkgroupZoneSource = LinkgroupZoneDetail(0))) Then
                                ' for playing announcement files, the source zone info was already stored
                                ' this means the zone was not found so save it
                                Try
                                    MyLocalLinkgroupArray = MyLinkgroupArray.GetLinkGroupInfo(LinkgroupName)
                                    MySavedPlayerList = MyLocalLinkgroupArray.MySavedPlayerList
                                    If Not MySavedPlayerList.IsAlreadyStored(DestinationSonosPlayer.GetUDN) Then
                                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn: Start Save Current Track Info for player = " & DestinationSonosPlayer.GetZoneName, LogType.LOG_TYPE_INFO)
                                        DestinationSonosPlayer.SaveCurrentTrackInfo(LinkgroupName, False, False)
                                        MySavedPlayerList.Add(DestinationSonosPlayer.GetUDN)
                                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn: End Save Current Track Info for player = " & DestinationSonosPlayer.GetZoneName, LogType.LOG_TYPE_INFO)
                                    End If
                                Catch ex As Exception
                                    Log("Error in SaveLinkedPlayers getting the SavedPlayerList with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try

                            End If
                        Catch ex As Exception
                            Log("Error in HandleLinkingOn getting Controller for LinkgroupZoneDestination = " & LinkgroupZone & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Log("DestinationZoneIndex = " & DestinationSonosPlayer.GetUDN, LogType.LOG_TYPE_ERROR)
                        End Try
                        ' 4/15/2020 v.53 moved here else we cause exception which is not caught and caused doCheckAnnouncementqueue to hang
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn: Done Saving LinkgroupZoneDestination = " & DestinationSonosPlayer.GetZoneName, LogType.LOG_TYPE_INFO)
                    End If
                End If
            End If
        Next

        'Try
        'If SourceSonosPlayer.ZoneIsPairMaster Then
        ' separate pair
        'SourceSonosPlayer.SeparateStereoPair()
        'wait(0.5) ' introduced in v.85 because a paired player, which is part of source and dest, can get unpaired twice because the change event doesn't get through before before we check later on at the dest player
        'ElseIf SourceSonosPlayer.ZoneIsPairSlave Then
        'Dim MasterPairPlayer As HSPI
        'MasterPairPlayer = GetSonosPlayerByUDN(SourceSonosPlayer.GetZonePairMasterUDN)
        'MasterPairPlayer.SeparateStereoPair()
        'wait(0.5)
        'End If
        'Catch ex As Exception
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in HandleLinkingOn while separating a Stereo Pair with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        'End Try
        ' moved this post seperation in v.85 because when the slave zone in a paired is source for an announcement, this stop command will be forwarded to the master and
        ' after the announcement is over, the master will never get a play nor was the master saved
        If IsFile Then
            If SourceSonosPlayer.ZoneIsLinked Then  ' add here code for unlinking in case it is linked!  v .92 !!
                SourceSonosPlayer.PlayURI("x-rincon-queue:" & SourceSonosPlayer.GetUDN & "#0", "")
                wait(0.5)
            Else
                SourceSonosPlayer.SetTransportState("Stop", True) ' stop it
            End If
        End If

        For Each LinkgroupZone In LinkgroupZoneDestinations
            LinkgroupZoneDetail = Split(LinkgroupZone, ";")
            If UBound(LinkgroupZoneDetail) > 2 Then
                If LinkgroupZoneDetail(3) = "1" Then
                    Dim NewZoneVolume As String = ""
                    Dim NewMuteOverride As Boolean = False
                    If UBound(LinkgroupZoneDetail) > 1 Then
                        ' we have a new structure with Mute flag
                        If LinkgroupZoneDetail(2) = "1" Then NewMuteOverride = True
                        NewZoneVolume = Trim(LinkgroupZoneDetail(1))
                    ElseIf UBound(LinkgroupZoneDetail) > 0 Then
                        NewZoneVolume = Trim(LinkgroupZoneDetail(1))
                    End If
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn: Start LinkgroupZoneDestination = " & LinkgroupZoneDetail(0), LogType.LOG_TYPE_INFO)
                    DestinationSonosPlayer = GetAPIByUDN(LinkgroupZoneDetail(0))
                    If DestinationSonosPlayer Is Nothing Then
                        Log("Error in HandleLinkingOn for LinkgroupZoneDestination = " & LinkgroupZoneDetail(0) & " not found while linking", LogType.LOG_TYPE_ERROR)
                    Else
                        If DestinationSonosPlayer.ZoneIsASlave Then
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn detected PairedSlave = " & DestinationSonosPlayer.GetZoneName & " and will switch to the Master", LogType.LOG_TYPE_INFO)
                            DestinationSonosPlayer = MyHSPIControllerRef.GetAPIByUDN(DestinationSonosPlayer.ZoneMasterUDN)
                        End If
                        Try
                            Dim HSIpAddress As String
                            HSIpAddress = hs.GetIPAddress()
                            Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                            ' HTTPPort = hs.GetINISetting("Settings", "gWebSvrPort", "") 
                            If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                            Dim MetaData As String = ""
                            MetaData = "<DIDL-Lite xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:upnp=""urn:schemas-upnp-org:metadata-1-0/upnp/"" xmlns:r=""urn:schemas-rinconnetworks-com:metadata-1-0/"" xmlns=""urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/""><item id=""-1"" parentID=""-1"" restricted=""true"">"
                            MetaData = MetaData & "<upnp:albumArtURI>http://"
                            MetaData = MetaData & HSIpAddress & HTTPPort & URLImagesPath & "Announcement.jpg</upnp:albumArtURI><dc:title>"
                            MetaData = MetaData & AnnouncementTitle & "</dc:title><upnp:class>object.item.audioItem.musicTrack</upnp:class><dc:creator>"
                            MetaData = MetaData & AnnouncementAuthor & "</dc:creator><upnp:album>"
                            MetaData = MetaData & AnnouncementAlbum & "</upnp:album><r:albumArtist>"
                            MetaData = MetaData & AnnouncementAuthor & "</r:albumArtist></item></DIDL-Lite>"

                            If SourceSonosPlayer.GetUDN <> DestinationSonosPlayer.GetUDN Then ' this part was rewritten for V3.0.0.20
                                ' only when TTS mode should we have the source zone participate as destination. In all other cases this will cause a loop
                                If SourceSonosPlayer.ZoneModel <> "WD100" Then
                                    DestinationSonosPlayer.PlayURI("x-rincon:" & SourceSonosPlayer.GetUDN, MetaData) ' group
                                    If MyAnnouncementWaitBetweenPlayers > 0 Then
                                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn waiting for " & MyAnnouncementWaitBetweenPlayers & " 10th of a seconds between issues play commands to other players", LogType.LOG_TYPE_INFO)
                                        wait(MyAnnouncementWaitBetweenPlayers / 10)
                                    End If
                                Else
                                    DestinationSonosPlayer.PlayURI("x-sonos-dock:" & SourceSonosPlayer.GetUDN, MetaData) ' group
                                End If
                            Else
                                If UseLinkGroupAsTTS And Not IsFile Then
                                    If SourceSonosPlayer.ZoneModel <> "WD100" Then
                                        DestinationSonosPlayer.PlayURI("x-rincon-stream:" & SourceSonosPlayer.GetUDN, MetaData) ' set to line input
                                        DestinationSonosPlayer.SetTransportState("Play", True)
                                    Else
                                        DestinationSonosPlayer.PlayURI("x-sonos-dock:" & SourceSonosPlayer.GetUDN, MetaData) ' group
                                    End If
                                End If
                            End If
                            Dim Difference As Integer = 0 ' was 1, let's see what happens
                            If NewMuteOverride And DestinationSonosPlayer.PlayerMute Then
                                DestinationSonosPlayer.UnMute()
                                If NewZoneVolume = "" Then
                                    ' this is an attempt to add an artificial delay based on current volume and requested volume
                                    If MaxWaitForVolumeToAdjust < 2 Then MaxWaitForVolumeToAdjust = 2
                                End If
                            End If
                            If NewZoneVolume <> "" Then
                                ' this is an attempt to add an artificial delay based on current volume and requested volume
                                Dim OldZoneVolume As Integer = DestinationSonosPlayer.PlayerVolume
                                Difference = Difference + System.Math.Abs(OldZoneVolume - NewZoneVolume) / 50
                                DestinationSonosPlayer.SetVolumeLevel("Master", NewZoneVolume)
                                If Difference <> 0 Then
                                    If MaxWaitForVolumeToAdjust < Difference Then MaxWaitForVolumeToAdjust = Difference
                                End If
                            End If
                        Catch ex As Exception
                            Log("Error in HandleLinkingOn getting Controller for LinkgroupZoneDestination = " & DestinationSonosPlayer.GetZoneName & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        ' 4/15/2020 v.53 moved here else we cause exception which is not caught and caused doCheckAnnouncementqueue to hang                   
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn: Done LinkgroupZoneDestination = " & DestinationSonosPlayer.GetZoneName, LogType.LOG_TYPE_INFO)
                    End If
                End If
            End If
        Next
        'If MaxWaitForVolumeToAdjust > 0 Then ' removed in v101
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "HandleLinkingOn waiting for " & MaxWaitForVolumeToAdjust & " seconds to adjust the volume")
        'Wait(MaxWaitForVolumeToAdjust)
        'End If
        HandleLinkingOn = MaxWaitForVolumeToAdjust ' changed in v101
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOn done with LinkgroupName = " & LinkgroupName & " and Delay = " & MaxWaitForVolumeToAdjust.ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Sub HandleLinkingOff(ByVal LinkgroupName As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOff called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim MySavedPlayerList As SavedPlayerList
        Dim MyLocalLinkgroupArray As LinkGroupInfo
        Try
            MyLocalLinkgroupArray = MyLinkgroupArray.GetLinkGroupInfo(LinkgroupName)
            MySavedPlayerList = MyLocalLinkgroupArray.MySavedPlayerList
        Catch ex As Exception
            Log("Error 1 in HandleLinkingOff with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        Dim TempPlayer As HSPI
        ' ungroup all players first. Error in v.94
        MySavedPlayerList.ResetIndex()
        Dim TempUDN As String = ""
        TempUDN = MySavedPlayerList.GetNext

        While TempUDN <> ""
            Try
                TempPlayer = GetAPIByUDN(TempUDN)
            Catch ex As Exception
                Log("Error in HandleLinkingOff: Couldn't find zoneplayer for UDN = " & TempUDN & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If TempPlayer IsNot Nothing Then
                Try
                    TempPlayer.PlayURI("x-rincon-queue:" & TempPlayer.GetUDN & "#0", "")
                    wait(0.25)
                Catch ex As Exception
                    Log("Error 4 in HandleLinkingOff with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            TempUDN = MySavedPlayerList.GetNext
        End While
        'wait(0.5)

        Dim NbrOfPlayers As Integer = 0
        Dim LastUDN As String = ""
        Try
            LastUDN = MySavedPlayerList.GetLastUDN
        Catch ex As Exception
            Log("Error 2 in HandleLinkingOff with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        While LastUDN <> ""
            Try
                TempPlayer = GetAPIByUDN(LastUDN)
            Catch ex As Exception
                Log("Error in HandleLinkingOff: Couldn't find zoneplayer for ZoneUDN = " & LastUDN & " with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
            If TempPlayer IsNot Nothing Then
                Try
                    TempPlayer.RestoreCurrentTrackInfo(LinkgroupName, False)
                    NbrOfPlayers += 1
                Catch ex As Exception
                    Log("Error 3 in HandleLinkingOff with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            LastUDN = MySavedPlayerList.GetLastUDN
        End While
        If NbrOfPlayers <> 0 Then
            wait(NbrOfPlayers)
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("HandleLinkingOff done with LinkgroupName = " & LinkgroupName & " and # players = " & NbrOfPlayers.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Public Function CheckPlayerInLinkgroup(ByVal SearchZoneUDN As String, ByVal SourceZoneUDN As String, ByVal DestinationZoneString As String) As Boolean
        CheckPlayerInLinkgroup = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckPlayerInLinkgroup called with SearchZoneUDN = " & SearchZoneUDN & " and SourceZoneUDN = " & SourceZoneUDN & " and DestinationZoneString = " & DestinationZoneString.ToString, LogType.LOG_TYPE_INFO)
        SearchZoneUDN = GetUDNbyZoneName(SearchZoneUDN) ' just in case this procedure was called with a zone name (old implementation)
        SourceZoneUDN = GetUDNbyZoneName(SourceZoneUDN)
        If SearchZoneUDN = SourceZoneUDN Then
            CheckPlayerInLinkgroup = True
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckPlayerInLinkgroup found match for SearchZoneUDN = " & SearchZoneUDN & " and SourceZoneUDN", LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        Dim DestinationZoneStrings As String()
        DestinationZoneStrings = Split(DestinationZoneString, "|")
        Dim TempLinkGroupZoneDetails As String()
        For Each TempLinkGroupZone In DestinationZoneStrings
            If TempLinkGroupZone <> "" Then
                TempLinkGroupZoneDetails = Split(TempLinkGroupZone, ";")
                If GetUDNbyZoneName(TempLinkGroupZoneDetails(0)) = SearchZoneUDN Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckPlayerInLinkgroup found match for SearchZoneUDN = " & SearchZoneUDN & " in DestinationZoneString = " & DestinationZoneString.ToString, LogType.LOG_TYPE_INFO)
                    CheckPlayerInLinkgroup = True
                    Exit Function
                End If
            End If
        Next
    End Function

    Public Sub SaveLinkedPlayers(ByVal SearchZoneUDN As String, ByVal SourceZoneUDN As String, ByVal LinkGroupName As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveLinkedPlayers called with SearchZoneUDN = " & SearchZoneUDN & " and SourceZoneUDN = " & SourceZoneUDN & " and LinkGroupName = " & LinkGroupName, LogType.LOG_TYPE_INFO)
        'SearchZoneUDN = GetUDNbyZoneName(SearchZoneUDN) ' just in case this procedure was called with a zone name (old implementation)
        'SourceZoneUDN = GetUDNbyZoneName(SourceZoneUDN)
        Dim LinkgroupZoneDestination As String
        LinkgroupZoneDestination = GetStringIniFile("LinkgroupZoneDestination", LinkGroupName, "")
        Dim SearchSonosPlayer As HSPI
        Try
            SearchSonosPlayer = GetAPIByUDN(SearchZoneUDN)
            Dim TargetZones As String()
            Dim TempLinkGroupZoneDetails As String()
            TargetZones = SearchSonosPlayer.GetZoneDestination
            If TargetZones.Count = 0 Then Exit Sub
            Dim WasFound As Boolean = False
            For Each TargetZone In TargetZones
                WasFound = False
                Dim TargetPlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(TargetZone)
                If TargetPlayer IsNot Nothing Then
                    If TargetPlayer.ZoneIsASlave Then
                        TargetPlayer = MyHSPIControllerRef.GetAPIByUDN(TargetPlayer.ZoneMasterUDN)
                    End If
                    If TargetPlayer.GetUDN <> SourceZoneUDN Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveLinkedPlayers is comparing TargetZone = " & TargetPlayer.GetZoneName & " in String = " & LinkgroupZoneDestination & " and SourceZoneUDN = " & SourceZoneUDN, LogType.LOG_TYPE_INFO)
                        Dim DestinationZoneStrings As String()
                        DestinationZoneStrings = Split(LinkgroupZoneDestination, "|")
                        For Each TempLinkGroupZone In DestinationZoneStrings
                            If TempLinkGroupZone <> "" Then
                                TempLinkGroupZoneDetails = Split(TempLinkGroupZone, ";")
                                If TempLinkGroupZoneDetails(3) = "1" Then ' is this zone selected?
                                    If Mid(TempLinkGroupZoneDetails(0), 1, 5) = "uuid:" Then
                                        Mid(TempLinkGroupZoneDetails(0), 1, 5) = "     "
                                        TempLinkGroupZoneDetails(0) = Trim(TempLinkGroupZoneDetails(0))
                                    End If
                                    Dim TempPlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(TempLinkGroupZoneDetails(0))
                                    If TempPlayer IsNot Nothing Then
                                        If TempPlayer.ZoneIsASlave Then
                                            TempPlayer = MyHSPIControllerRef.GetAPIByUDN(TempPlayer.ZoneMasterUDN)
                                        End If
                                        ' 4/15/2020 moved code here to avoid error when tempPlayer is not found v.53
                                        If TempPlayer.GetUDN = TargetPlayer.GetUDN Then
                                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveLinkedPlayers found a match between announcement link group and Linkgroup with Linked Zone = " & TargetPlayer.GetZoneName & " and LinkgroupZone = " & TempPlayer.GetZoneName, LogType.LOG_TYPE_INFO)
                                            ' found 
                                            WasFound = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If Not WasFound Then
                            ' this means the zone was not found so save it
                            Dim MySavedPlayerList As SavedPlayerList
                            Dim MyLocalLinkgroupArray As LinkGroupInfo
                            Try
                                MyLocalLinkgroupArray = MyLinkgroupArray.GetLinkGroupInfo(LinkGroupName)
                                MySavedPlayerList = MyLocalLinkgroupArray.MySavedPlayerList
                                If Not MySavedPlayerList.IsAlreadyStored(TargetPlayer.GetUDN) Then
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveLinkedPlayers did not find the linked zone in the announcement Linkgroup, most likely a paired/linked player. Linked ZoneUDN = " & TargetZone, LogType.LOG_TYPE_INFO)
                                    TargetPlayer.SaveCurrentTrackInfo(LinkGroupName, False, False)
                                    MySavedPlayerList.Add(TargetPlayer.GetUDN)
                                End If
                            Catch ex As Exception
                                Log("Error in SaveLinkedPlayers getting the SavedPlayerList with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in SaveLinkedPlayers with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub RestoreLinkedPlayers(ByVal SearchZoneUDN As String, ByVal SourceZoneUDN As String, ByVal DestinationZoneString As String, ByVal LinkGroupName As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreLinkedPlayers called with SearchZoneUDN = " & SearchZoneUDN & " and SourceZoneUDN = " & SourceZoneUDN & " and DestinationZoneString = " & DestinationZoneString.ToString & " and LinkGroupName = " & LinkGroupName, LogType.LOG_TYPE_INFO)
        SearchZoneUDN = GetUDNbyZoneName(SearchZoneUDN) ' just in case this procedure was called with a zone name (old implementation)
        SourceZoneUDN = GetUDNbyZoneName(SourceZoneUDN)
        Dim LinkgroupZoneDestination As String
        LinkgroupZoneDestination = GetStringIniFile("LinkgroupZoneDestination", LinkGroupName, "")
        Dim SearchSonosPlayer As HSPI
        Try
            SearchSonosPlayer = GetAPIByUDN(SearchZoneUDN)
            Dim TargetZones As String()
            Dim TempLinkGroupZoneDetails As String()
            TargetZones = SearchSonosPlayer.GetZoneDestination
            Dim WasFound As Boolean = False
            For Each TargetZone In TargetZones
                WasFound = False
                If TargetZone <> "" And TargetZone <> SourceZoneUDN Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreLinkedPlayers is comparing TargetZone = " & TargetZone & " in String = " & LinkgroupZoneDestination, LogType.LOG_TYPE_INFO)
                    Dim DestinationZoneStrings As String()
                    DestinationZoneStrings = Split(LinkgroupZoneDestination, "|")
                    For Each TempLinkGroupZone In DestinationZoneStrings
                        If TempLinkGroupZone <> "" Then
                            TempLinkGroupZoneDetails = Split(TempLinkGroupZone, ";")
                            If Mid(TempLinkGroupZoneDetails(0), 1, 5) = "uuid:" Then
                                Mid(TempLinkGroupZoneDetails(0), 1, 5) = "     "
                                TempLinkGroupZoneDetails(0) = Trim(TempLinkGroupZoneDetails(0))
                            End If
                            If GetUDNbyZoneName(TempLinkGroupZoneDetails(0)) = TargetZone Then
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreLinkedPlayers found match between announcement link group and Linkgroup with Linked Zone = " & TargetZone & " and LinkgroupZone = " & GetUDNbyZoneName(TempLinkGroupZoneDetails(0)), LogType.LOG_TYPE_INFO)
                                ' found 
                                WasFound = True
                                Exit For
                            End If
                        End If
                    Next
                    If Not WasFound Then
                        ' this means the zone was not found so save it
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("RestoreLinkedPlayers did not find the linked zone in the announcement Linkgroup. Linked Zone = " & TargetZone, LogType.LOG_TYPE_WARNING)
                        Dim SaveSonosPlayer As HSPI
                        SaveSonosPlayer = GetAPIByUDN(TargetZone)
                        If SaveSonosPlayer IsNot Nothing Then
                            SaveSonosPlayer.RestoreCurrentTrackInfo(LinkGroupName, False)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in RestoreLinkedPlayers with error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub AddAnnouncementToQueue(ByVal LinkGroupName As String, ByVal device As Short, ByVal text As String, ByVal host As String)
        'AnnouncementsInQueue = True  ' moved in v83 to the end of this procedure
        'MyTimeoutActionArray(TOCheckAnnouncement) = 1
        'MyAnnouncementCountdown = 100
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddAnnouncementToQueue called for LinkGroupName = " & LinkGroupName & " and Text = " & text, LogType.LOG_TYPE_INFO)
        Dim AnnouncementItem As New AnnouncementItems With {
            .LinkGroupName = LinkGroupName,
            .device = device,
            .text = text,
            .host = host,
            .IsFile = Not GetLinkgroupSourceZoneAudioInputFlag(LinkGroupName),
            .SourceZoneMusicAPI = GetLinkgroupSourceZone(LinkGroupName)
        }
        If AnnouncementItem.SourceZoneMusicAPI Is Nothing Then
            Log("Error in AddAnnouncementToQueue for LinkGroupName = " & LinkGroupName & ". SourceZone was not found", LogType.LOG_TYPE_ERROR)
            AnnouncementItem = Nothing
            Exit Sub
        End If
        Dim LastAnnouncementInQueue As AnnouncementItems
        Try
            LastAnnouncementInQueue = GetTailOfAnnouncementQueue()
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue getting tail with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            AnnouncementItem.Previous_ = LastAnnouncementInQueue
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue setting previous with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            If LastAnnouncementInQueue Is Nothing Then
                ' this is the very first
                AnnouncementLink = AnnouncementItem
            Else
                LastAnnouncementInQueue.Next_ = AnnouncementItem
            End If
        Catch ex As Exception
            Log("Error in AddAnnouncementToQueue setting Next with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        ' moved this to the end of procedure to avoid race condition where TO services calls the DoCheckAnnouncementqueue before it was set up.
        ' changed in v.83
        AnnouncementsInQueue = True
        MyTimeoutActionArray(TOCheckAnnouncement) = 1
        MyAnnouncementCountdown = MyMaxAnnouncementTime
    End Sub

    Private Function GetTailOfAnnouncementQueue() As AnnouncementItems
        GetTailOfAnnouncementQueue = AnnouncementLink
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called", LogType.LOG_TYPE_INFO)
        If AnnouncementLink Is Nothing Then
            Exit Function
        End If
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink
        Dim LoopIndex As Integer = 0
        Try
            Do
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called and found Linkgroup = " & AnnouncementItem.LinkGroupName & " and text = " & AnnouncementItem.text, LogType.LOG_TYPE_INFO)
                If AnnouncementItem.Next_ Is Nothing Then
                    GetTailOfAnnouncementQueue = AnnouncementItem
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetTailOfAnnouncementQueue called and tail found", LogType.LOG_TYPE_INFO)
                    Exit Function
                End If
                AnnouncementItem = AnnouncementItem.Next_
                LoopIndex = LoopIndex + 1
                If LoopIndex > 100 Then
                    ' we have a loop, force clean it
                    Log("Error in GetTailOfAnnouncementQueue, loop found, clearing all Announcement info", LogType.LOG_TYPE_ERROR)
                    AnnouncementLink = Nothing
                    Exit Function
                End If
            Loop
        Catch ex As Exception
            Log("Error in GetTailOfAnnouncementQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub DeleteHeadOfAnnouncementQueue()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DeleteHeadOfAnnouncementQueue called", LogType.LOG_TYPE_INFO)
        If AnnouncementLink Is Nothing Then
            ' this should not be!
            Exit Sub
        End If
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink.Next_
        Try
            AnnouncementLink.Next_ = Nothing ' make sure there are no references left
            AnnouncementLink.SourceZoneMusicAPI = Nothing
            AnnouncementLink = Nothing ' return this memory
            AnnouncementLink = AnnouncementItem
            If Not AnnouncementLink Is Nothing Then
                AnnouncementLink.Previous_ = Nothing
            End If
        Catch ex As Exception
            Log("Error in DeleteHeadOfAnnouncementQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DoCheckAnnouncementQueue()

        If Not AnnouncementsInQueue Then Exit Sub
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called AnnouncementinQueue = " & AnnouncementsInQueue.ToString & " and AnnouncementInProgress = " & AnnouncementInProgress.ToString & " and AnnouncementCountdown = " & MyAnnouncementCountdown.ToString & " announcementReEntryCounter = " & announcementReEntryCounter.ToString, LogType.LOG_TYPE_INFO)

        If announcementReEntryCounter > 0 Then
            If piDebuglevel > DebugLevel.dlEvents Then
                If AnnouncementLink IsNot Nothing Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called and cause re-entry. Announcement state = " & AnnouncementLink.State_.ToString, LogType.LOG_TYPE_WARNING)
                Else
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called and cause re-entry. No Announcementlink ?", LogType.LOG_TYPE_WARNING)
                End If
            End If
            announcementReEntryCounter -= 1
            Exit Sub ' re-entry
        End If
        announcementReEntryCounter = 100
        If AnnouncementInProgress Then
            If Not AnnouncementLink Is Nothing Then
                If AnnouncementLink.State_ = AnnouncementState.asLinking Then
                    ' this is re-entrance
                    announcementReEntryCounter = 0
                    Exit Sub
                End If
                If AnnouncementLink.IsFile Then
                    MyAnnouncementCountdown = MyAnnouncementCountdown - 1
                    If MyAnnouncementCountdown < 0 Then
                        ' OK this is really not good. Many seconds have gone by since the last announcement was added and we are still linked
                        AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Log("Error in DoCheckAnnouncementQueue. " & MyMaxAnnouncementTime.ToString & " seconds expired since the announcement started and no end was received.", LogType.LOG_TYPE_ERROR)
                    ElseIf Not AnnouncementLink.SourceZoneMusicAPI Is Nothing Then
                        If AnnouncementLink.SourceZoneMusicAPI.CurrentPlayerState = Player_state_values.Stopped And AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted Then
                            ' OK the announcement is over
                            AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Else
                            ' go check the PositionInfo, if <> 0 then the announcement has started, let's also check the playerstate!
                            AnnouncementLink.SourceZoneMusicAPI.CheckAnnouncementHasStarted() ' added in v109 
                            If Not AnnouncementLink.SourceZoneMusicAPI.HasAnnouncementStarted Then
                                If MyAnnouncementCountdown <= MyMaxAnnouncementTime - 5 And ResendPlay = 0 Then
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DoCheckAnnouncementQueue. Announcement hasn't started after " & (MyMaxAnnouncementTime - MyAnnouncementCountdown + 1).ToString & " seconds", LogType.LOG_TYPE_WARNING)
                                    'AnnouncementLink.SourceZoneMusicAPI.SetTransportState("Play", True) ' was removed/added in v0.91 removed in v.93
                                    'AnnouncementLink.SourceZoneMusicAPI.SubmitDiagnostics() ' was removed/added for test purposes in v0.91
                                    ResendPlay = 1
                                ElseIf MyAnnouncementCountdown <= MyMaxAnnouncementTime - 30 And ResendPlay = 1 Then
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DoCheckAnnouncementQueue. Announcement hasn't started after " & (MyMaxAnnouncementTime - MyAnnouncementCountdown + 1).ToString & " seconds", LogType.LOG_TYPE_WARNING)
                                    'AnnouncementLink.SourceZoneMusicAPI.SetTransportState("Play", True) ' was removed/added in v0.91 removed in v.93
                                    'AnnouncementLink.SourceZoneMusicAPI.SubmitDiagnostics() ' was removed/added for test purposes in v0.91
                                    ResendPlay = 2
                                    ' the announcement has started and is running.
                                    ' Having trouble with repeat_all showing up, in v.18 I'll add a check and force no-repeat in case it is wrong
                                    If AnnouncementLink.SourceZoneMusicAPI.SonosRepeat <> Repeat_modes.repeat_off Then
                                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in DoCheckAnnouncementQueue. Sourceplayer mysteriously jumped to repeat state", LogType.LOG_TYPE_WARNING)
                                        AnnouncementLink.SourceZoneMusicAPI.PlayModeNormal()
                                    End If
                                End If
                            End If
                            announcementReEntryCounter = 0
                            Exit Sub
                        End If
                    Else
                        AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                        Log("Error in DoCheckAnnouncementQueue. Timer is running but there is no instance of a Sonos Player Object.", LogType.LOG_TYPE_ERROR)
                    End If
                Else
                    AnnouncementLink.State_ = AnnouncementState.asFilePlayed
                    Log("Error in DoCheckAnnouncementQueue. Timer is running but Announcement is not marked as Speak-To-File", LogType.LOG_TYPE_ERROR)
                End If
            Else
                ' this should not be
                AnnouncementsInQueue = False
                AnnouncementInProgress = False
                MyAnnouncementIndex = 0
                announcementReEntryCounter = 0
                Exit Sub
            End If
        End If
        AnnouncementInProgress = True

        If AnnouncementLink Is Nothing Then
            ' this should not be
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue. No AnnouncementLink", LogType.LOG_TYPE_ERROR)
            AnnouncementsInQueue = False
            AnnouncementInProgress = False
            MyAnnouncementIndex = 0
            announcementReEntryCounter = 0
            Exit Sub
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue called for linkgroup " & AnnouncementLink.LinkGroupName & " and State = " & AnnouncementLink.State_.ToString & " and isFile = " & AnnouncementLink.IsFile.ToString, LogType.LOG_TYPE_INFO)
        ' Look at first announcement in Queue
        Dim AnnouncementItem As AnnouncementItems
        AnnouncementItem = AnnouncementLink
        If AnnouncementItem.State_ = AnnouncementState.asIdle Then
            ' Announcement event: add new event here
            PlayChangeNotifyCallback(Player_status_change.AnnouncementChange, Player_state_values.AnnouncementStart, True)
            ResendPlay = 0
            AnnouncementItem.State_ = AnnouncementState.asLinking
            If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue called HandleLinking", LogType.LOG_TYPE_INFO)
            AnnouncementItem.DelayinSec = HandleLinkingOn(AnnouncementItem.LinkGroupName, AnnouncementLink.IsFile)
            AnnouncementItem.AbsoluteTime = Now
            If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue done calling HandleLinking", LogType.LOG_TYPE_INFO)
            If AnnouncementItem.IsFile Then
                Try ' added on 7/12/2019 in v3.1.0.31 If source player is not the master, the monitoring and other actions won't work
                    If AnnouncementItem.SourceZoneMusicAPI.ZoneIsASlave Then
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue found source player to be a slave and switched to the master", LogType.LOG_TYPE_INFO)
                        AnnouncementItem.SourceZoneMusicAPI = MyHSPIControllerRef.GetAPIByUDN(AnnouncementItem.SourceZoneMusicAPI.ZoneMasterUDN)
                    End If
                Catch ex As Exception
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue switching to the master with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Try
                    AnnouncementItem.SourceZoneMusicAPI.ClearQueue()
                    ' Also reset shuffle and repeat to avoid reordering and endless repeats
                    AnnouncementItem.SourceZoneMusicAPI.PlayURI("x-rincon-queue:" & AnnouncementItem.SourceZoneMusicAPI.GetUDN & "#0", "") ' added v111
                    AnnouncementItem.SourceZoneMusicAPI.PlayModeNormal()
                Catch ex As Exception
                End Try
            End If
            AnnouncementItem.State_ = AnnouncementState.asLinked
        ElseIf AnnouncementItem.State_ = AnnouncementState.asLinking Then
            ' this is re-entrance
            If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue is waiting in Linking state and existing procedure", LogType.LOG_TYPE_INFO)
            announcementReEntryCounter = 0
            Exit Sub
        End If
        Dim StartQueueIndex As Integer = MyAnnouncementIndex + 1
        If AnnouncementItem.State_ = AnnouncementState.asLinked Then
            Dim TextStrings
            TextStrings = Split(AnnouncementItem.text, "|")
            Dim Index As Integer = 0
            For Each TextString In TextStrings
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is Linked and activated with HostName = " & AnnouncementItem.host & " Text = " & TextString & " and LinkgroupName = " & AnnouncementItem.LinkGroupName, LogType.LOG_TYPE_INFO)
                If AnnouncementItem.IsFile Then
                    Dim FileName As String
                    Dim ExtensionIndex As Integer = 0
                    Dim Extensiontype As String = ""
                    Dim Path As String = ""
                    If HSisRunningOnLinux Then ' this will always be on the HS machine
                        Path = CurrentAppPath & "/html" & AnnouncementPath
                    Else
                        Path = CurrentAppPath & "\html" & AnnouncementPath
                    End If
                    FileName = "Ann_" & RemoveBlanks(AnnouncementItem.LinkGroupName) & "_" & MyAnnouncementIndex.ToString

                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue adds file = " & Path & FileName & " to Queue", LogType.LOG_TYPE_INFO)
                    'hs.SpeakToFile(Text, "VoiceCB", FileName, "")
                    If File.Exists(TextString) Then
                        Try
                            ' get the extension file type
                            ExtensionIndex = TextString.lastindexof(".")
                            If ExtensionIndex <> -1 Then
                                Extensiontype = TextString.Substring(ExtensionIndex, TextString.length - ExtensionIndex)
                            End If
                            FileName = FileName + Extensiontype
                        Catch ex As Exception
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when searching for the file type with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            System.IO.File.Delete(Path & FileName)
                        Catch ex As Exception
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "Error in DoCheckAnnouncementQueue when deleting file " & Path & FileName & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            System.IO.File.Copy(TextString, Path & FileName, True)
                            'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log( "DoCheckAnnouncementQueue copying file " & Path & FileName)
                            AnnouncementItem.State_ = AnnouncementState.asSpeaking
                        Catch ex As Exception
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue in SpeakToFile copying file = " & TextString & " to " & Path & FileName & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                            AnnouncementInProgress = False
                            AnnouncementItem.State_ = AnnouncementState.asFilePlayed
                            announcementReEntryCounter = 0
                            Exit Sub
                        End Try
                    Else
                        FileName = FileName & ".wav"
                        Try
                            System.IO.File.Delete(Path & FileName)
                            'Log( "DoCheckAnnouncementQueue deleted " & Path & FileName & " successfully")
                        Catch ex As Exception
                            'Log( "DoCheckAnnouncementQueue deleted " & Path & FileName & " un-successfully with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        AnnouncementItem.State_ = AnnouncementState.asSpeaking
                        Try
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue calling SpeakToFile with Text " & TextString & " and File " & Path & FileName, LogType.LOG_TYPE_INFO)
                            Dim Voice As String = CheckForVoiceTag(TextString)
                            hs.SpeakToFile(TextString, Voice, Path & FileName)
                            If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue finished SpeakToFile", LogType.LOG_TYPE_INFO)
                        Catch ex As Exception
                            Log("Error in DoCheckAnnouncementQueue called SpeakToFile unsuccessfully with Text " & TextString & " and File " & Path & FileName & " and error " & ex.Message, LogType.LOG_TYPE_ERROR)
                            ' added 9/4/2019 because when failing, it will take forever to timeout
                            AnnouncementItem.State_ = AnnouncementState.asFilePlayed
                            announcementReEntryCounter = 0
                            Exit Sub
                        End Try
                    End If

                    Dim HSIpAddress As String
                    HSIpAddress = hs.GetIPAddress()
                    Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                    ' HTTPPort = hs.GetINISetting("Settings", "gWebSvrPort", "") 
                    If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                    Dim MetaData As String = ""
                    MetaData = "<DIDL-Lite xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:upnp=""urn:schemas-upnp-org:metadata-1-0/upnp/"" xmlns:r=""urn:schemas-rinconnetworks-com:metadata-1-0/"" xmlns=""urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/""><item id=""-1"" parentID=""-1"" restricted=""true""><res protocolInfo=""http-get:*:audio/wav:*"">http://"
                    MetaData = MetaData & HSIpAddress & HTTPPort & AnnouncementURL & FileName & "</res><r:streamContent></r:streamContent><upnp:albumArtURI>http://"
                    MetaData = MetaData & HSIpAddress & HTTPPort & URLImagesPath & "Announcement.jpg</upnp:albumArtURI><dc:title>"
                    MetaData = MetaData & AnnouncementTitle & "</dc:title><upnp:class>object.item.audioItem.musicTrack</upnp:class><dc:creator>"
                    MetaData = MetaData & AnnouncementAuthor & "</dc:creator><upnp:album>"
                    MetaData = MetaData & AnnouncementAlbum & "</upnp:album><r:albumArtist>"
                    MetaData = MetaData & AnnouncementAuthor & "</r:albumArtist></item></DIDL-Lite>"
                    ' adding this in front of the albumarturi might actually have it display on the controller itself  /getaa?m=1&amp;u=
                    If UBound(TextStrings) > 0 Then
                        ' Multiple Announcements, queue them up

                        Try
                            If UCase(Extensiontype) = ".MP3" Then
                                ' fixed in v.99 AnnouncementItem.SourceZoneMusicAPI.AddTrackToQueue("x-rincon-mp3radio://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, MetaData, MyAnnouncementIndex + 1, False)
                                AnnouncementItem.SourceZoneMusicAPI.AddTrackToQueue("http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, MetaData, MyAnnouncementIndex + 1, False)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is adding a track to the Queue = http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, LogType.LOG_TYPE_INFO)
                            Else
                                AnnouncementItem.SourceZoneMusicAPI.AddTrackToQueue("http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, MetaData, MyAnnouncementIndex + 1, False)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is adding a track to the Queue = http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, LogType.LOG_TYPE_INFO)
                            End If
                        Catch ex As Exception
                            Log("Error in DoCheckAnnouncementQueue when adding Announcement to Sonos Queue with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        If Index >= UBound(TextStrings) Then
                            ' Last announcement, start playing
                            AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = False
                            Try
                                'If AnnouncementItem.SourceZoneMusicAPI.PlayerState <> player_state_values.Playing Then // removed in v14
                                'If StartQueueIndex = 1 Then
                                ' we need to make sure the player is "forced" to play its queue. Seen errors here. The right syntax is x-rincon-queue:RINCON_000E5824C3B001400#0
                                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is playing URI = " & "x-rincon-queue:" & AnnouncementItem.SourceZoneMusicAPI.GetUDN & "#0", LogType.LOG_TYPE_INFO)
                                'AnnouncementItem.SourceZoneMusicAPI.PlayURI("x-rincon-queue:" & AnnouncementItem.SourceZoneMusicAPI.GetUDN & "#0", "")
                                Dim WaitTime As Integer = MyAnnouncementWaitToSendPlay
                                If AnnouncementItem.DelayinSec > MyAnnouncementWaitToSendPlay Then
                                    ' let compare how much time has already passed
                                    Dim diff As System.TimeSpan
                                    diff = Now.Subtract(AnnouncementItem.AbsoluteTime)
                                    If diff.Seconds < AnnouncementItem.DelayinSec Then
                                        AnnouncementItem.DelayinSec = AnnouncementItem.DelayinSec - diff.Seconds
                                        If AnnouncementItem.DelayinSec > MyAnnouncementWaitToSendPlay Then
                                            WaitTime = AnnouncementItem.DelayinSec
                                        End If
                                    End If
                                End If
                                'If AnnouncementItem.SourceZoneMusicAPI.ZoneIsLinked Then Wait(0.5)
                                If WaitTime <> 0 Then
                                    If piDebuglevel > DebugLevel.dlEvents Then Log("Waiting in DoCheckAnnouncementQueue before issues play for = " & WaitTime.ToString & " seconds", LogType.LOG_TYPE_INFO)
                                    wait(WaitTime)
                                    If piDebuglevel > DebugLevel.dlEvents Then Log("End Waiting in DoCheckAnnouncementQueue", LogType.LOG_TYPE_INFO)
                                End If
                                AnnouncementItem.SourceZoneMusicAPI.SetTransportState("Play", True)
                                '      Else
                                ' it either never started or already stopped
                                ' AnnouncementItem.SourceZoneMusicAPI.SeekTrack(StartQueueIndex)
                                'AnnouncementItem.SourceZoneMusicAPI.SetTransportState("Play", True) ' changed this from playifpaused 
                                'End If

                                ' End If
                            Catch ex As Exception
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when calling PlayURI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                ' added 9/4/2019 because when failing, it will take forever to timeout
                                AnnouncementItem.State_ = AnnouncementState.asFilePlayed
                                announcementReEntryCounter = 0
                                Exit Sub
                            End Try
                            MyAnnouncementIndex = MyAnnouncementIndex + 1
                            'wait(1) ' this is to make sure the playerstate has moved to playing before the timeout procedure begins checking for the "end of file" which is player stopped
                            If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue is waiting in for player to start playing w/ multiple announcements and existing procedure", LogType.LOG_TYPE_INFO)
                            announcementReEntryCounter = 0
                            Exit Sub
                        End If
                    Else
                        ' Single announcement
                        AnnouncementItem.SourceZoneMusicAPI.HasAnnouncementStarted = False
                        Try
                            If UCase(Extensiontype) = ".MP3" Then
                                ' fixed in v.99 AnnouncementItem.SourceZoneMusicAPI.PlayUri("x-rincon-mp3radio://" & HSIpAddress & HTTPPort & AnnouncementPath & FileName, MetaData)
                                AnnouncementItem.SourceZoneMusicAPI.PlayURI("http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, MetaData)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is calling PlayURI with http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, LogType.LOG_TYPE_INFO)
                            Else
                                AnnouncementItem.SourceZoneMusicAPI.PlayURI("http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, MetaData)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is calling PlayURI with http://" & HSIpAddress & HTTPPort & AnnouncementURL & FileName, LogType.LOG_TYPE_INFO)
                            End If
                        Catch ex As Exception
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when adding Announcement to Sonos Queue with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Dim WaitTime As Integer = MyAnnouncementWaitToSendPlay
                        If AnnouncementItem.DelayinSec > MyAnnouncementWaitToSendPlay Then
                            ' let compare how much time has already passed
                            Dim diff As System.TimeSpan
                            diff = Now.Subtract(AnnouncementItem.AbsoluteTime)
                            If diff.Seconds < AnnouncementItem.DelayinSec Then
                                AnnouncementItem.DelayinSec = AnnouncementItem.DelayinSec - diff.Seconds
                                If AnnouncementItem.DelayinSec > MyAnnouncementWaitToSendPlay Then
                                    WaitTime = AnnouncementItem.DelayinSec
                                End If
                            End If
                        End If
                        'If AnnouncementItem.SourceZoneMusicAPI.ZoneIsLinked Then Wait(0.5)
                        If WaitTime <> 0 Then
                            If piDebuglevel > DebugLevel.dlEvents Then Log("Waiting in DoCheckAnnouncementQueue before issues play for = " & WaitTime.ToString & " seconds", LogType.LOG_TYPE_INFO)
                            wait(WaitTime)
                            If piDebuglevel > DebugLevel.dlEvents Then Log("End Waiting in DoCheckAnnouncementQueue", LogType.LOG_TYPE_INFO)
                        End If
                        Try
                            AnnouncementItem.SourceZoneMusicAPI.SetTransportState("Play", True)
                        Catch ex As Exception
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue when calling PlayURI with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            ' added 9/4/2019 because when failing, it will take forever to timeout
                            AnnouncementItem.State_ = AnnouncementState.asFilePlayed
                            announcementReEntryCounter = 0
                            Exit Sub
                        End Try
                        'Wait(1) ' this is to make sure the playerstate has moved to playing before the timeout procedure begins checking for the "end of file" which is player stopped
                        MyAnnouncementIndex = MyAnnouncementIndex + 1
                        If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue is waiting in for player to start playing w/ single announcement and existing procedure", LogType.LOG_TYPE_INFO)
                        announcementReEntryCounter = 0
                        Exit Sub
                    End If

                Else
                    ' this is the old TTS way
                    AnnouncementItem.State_ = AnnouncementState.asSpeaking
                    If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue starts old TTS way speech", LogType.LOG_TYPE_INFO)
                    hs.SpeakProxy(0, TextString, True, AnnouncementItem.host)
                    If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue ends old TTS way speech", LogType.LOG_TYPE_INFO)
                    AnnouncementItem.State_ = AnnouncementState.asLinked
                End If
                Index = Index + 1
                MyAnnouncementIndex = MyAnnouncementIndex + 1
            Next
        End If
        If Not AnnouncementItem.Next_ Is Nothing Then
            ' there is more
            Try
                Dim NextAnnouncementItem As AnnouncementItems
                NextAnnouncementItem = AnnouncementItem.Next_
                If NextAnnouncementItem.LinkGroupName = AnnouncementItem.LinkGroupName Then
                    ' this is the same, so do not unlink
                    NextAnnouncementItem.State_ = AnnouncementState.asLinked ' indicate we are already linked
                    If AnnouncementItem.SourceZoneMusicAPI IsNot Nothing Then AnnouncementItem.SourceZoneMusicAPI.ClearQueue()         ' clear the queue to avoid repeat , added in v023
                    ' copy the sourceAPI to next announcement entry because there is no linking done, so this info will otherwise be lost - added 7/16/2019 in v3.1.0.35
                    NextAnnouncementItem.SourceZoneMusicAPI = AnnouncementItem.SourceZoneMusicAPI
                    DeleteHeadOfAnnouncementQueue()
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue was done but found new announcement item", LogType.LOG_TYPE_INFO)
                    AnnouncementInProgress = False
                    MyAnnouncementCountdown = MyMaxAnnouncementTime ' reset the clock
                    announcementReEntryCounter = 0
                    Exit Sub
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DoCheckAnnouncementQueue looking at next announcement in queue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        AnnouncementItem.State_ = AnnouncementState.asUnlinking
        If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue is starting to unlink", LogType.LOG_TYPE_INFO)
        HandleLinkingOff(AnnouncementItem.LinkGroupName)
        If piDebuglevel > DebugLevel.dlEvents Then Log("DoCheckAnnouncementQueue is done unlinking", LogType.LOG_TYPE_INFO)
        ' annoucementevent: off
        PlayChangeNotifyCallback(Player_status_change.AnnouncementChange, Player_state_values.AnnouncementStop, True)
        AnnouncementItem.State_ = AnnouncementState.asIdle
        DeleteHeadOfAnnouncementQueue()
        If AnnouncementLink Is Nothing Then
            AnnouncementsInQueue = False
            MyAnnouncementIndex = 0
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is ending and all announcements were processed", LogType.LOG_TYPE_INFO)
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DoCheckAnnouncementQueue is ending but more announcements are queued", LogType.LOG_TYPE_INFO)
        End If
        AnnouncementInProgress = False
        announcementReEntryCounter = 0
    End Sub

    Public Function GetLinkgroupSourceZone(ByVal LinkgroupName As String) As Object
        GetLinkgroupSourceZone = Nothing
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLinkgroupSourceZone called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim LinkgroupZoneSource As String
        LinkgroupZoneSource = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
        If LinkgroupZoneSource = "" Then
            ' This could be a zone Name
            Try
                GetLinkgroupSourceZone = MyHSPIControllerRef.GetAPIByUDN(LinkgroupName)
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLinkgroupSourceZone didn't find " & LinkgroupName & " under [LinkgroupZoneSource] in the .ini file", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            Dim LinkgroupZoneSourceDetails() As String
            LinkgroupZoneSourceDetails = Split(LinkgroupZoneSource, ";")
            LinkgroupZoneSource = LinkgroupZoneSourceDetails(0)
        End If
        Try
            GetLinkgroupSourceZone = MyHSPIControllerRef.GetAPIByUDN(LinkgroupZoneSource)
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetLinkgroupSourceZone didn't find MusicAPI for " & LinkgroupZoneSource, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetLinkgroupSourceZoneAudioInputFlag(ByVal LinkgroupName As String) As Boolean
        GetLinkgroupSourceZoneAudioInputFlag = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetLinkgroupSourceZoneAudioInputFlag called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        Dim LinkgroupZoneSource As String
        LinkgroupZoneSource = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
        If LinkgroupZoneSource = "" Then
            Exit Function
        End If
        Dim LinkgroupZoneSourceDetails() As String
        LinkgroupZoneSourceDetails = Split(LinkgroupZoneSource, ";")
        If UBound(LinkgroupZoneSourceDetails) > 0 Then
            If LinkgroupZoneSourceDetails(1) = 0 Then
                GetLinkgroupSourceZoneAudioInputFlag = False
            Else
                GetLinkgroupSourceZoneAudioInputFlag = True
            End If
        Else
            GetLinkgroupSourceZoneAudioInputFlag = False
        End If
    End Function

    Public Function RemoveBlanks(ByVal InString) As String
        RemoveBlanks = InString
        Dim InIndex As Integer = 0
        Dim Outstring As String = ""
        InString = Trim(InString)
        If InString = "" Then Exit Function
        Try
            Do While InIndex < InString.length
                If InString(InIndex) = " " Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "!" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = """ Then" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "#" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "$" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "%" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "&" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "'" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "(" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = ")" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "*" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "+" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "," Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "-" Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "." Then
                    Outstring = Outstring + "_"
                ElseIf InString(InIndex) = "/" Then
                    Outstring = Outstring + "_"
                Else
                    Outstring = Outstring & InString(InIndex)
                End If
                InIndex = InIndex + 1
            Loop
        Catch ex As Exception
            Log("Error in RemoveBlanks. URI = " & InString & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        RemoveBlanks = Outstring
    End Function

    Private Function CheckForVoiceTag(ByRef inText As String) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForVoiceTag called with inText = " & inText, LogType.LOG_TYPE_INFO)
        CheckForVoiceTag = ""
        ' Structure = <voice required='Name=Microsoft Anna'>Hello World how are things around here said Anna</voice>
        If inText.IndexOf("<voice ") = -1 Then
            Exit Function
        End If
        Try
            ' OK there appear to be something that looks like a tag
            Dim StartIndexVoiceTag As Integer = inText.IndexOf("<voice ")
            Dim EndIndexVoiceTag As Integer = inText.IndexOf(">", StartIndexVoiceTag)
            If EndIndexVoiceTag = -1 Then Exit Function ' shouldn't be!
            Dim StartIndexCloseVoiceTag As Integer = inText.IndexOf("</voice>", EndIndexVoiceTag)
            If StartIndexCloseVoiceTag = -1 Then Exit Function ' shouldn't be!
            Dim VoiceTagInfo As String = Trim(inText.Substring(StartIndexVoiceTag + 7, EndIndexVoiceTag - StartIndexVoiceTag - 7))
            inText = inText.Remove(StartIndexCloseVoiceTag, 8) ' remove the </voice> tag first
            inText = inText.Remove(StartIndexVoiceTag, EndIndexVoiceTag - StartIndexVoiceTag + 1)
            ' now the VoiceTagInfo should look something like this required='Name=Microsoft Anna'
            If VoiceTagInfo.IndexOf("optional") = 0 Then
                ' not sure what to do with this, won't make any difference, but could decide to simply ignore
                'Exit Function '??
                VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 8))
            ElseIf VoiceTagInfo.IndexOf("required") = 0 Then
                VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 8))
            End If
            ' now the VoiceTagInfo should look something like this ='Name=Microsoft Anna'
            If VoiceTagInfo.IndexOf("=") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))
            If VoiceTagInfo.IndexOf("'") <> 0 Then Exit Function ' should not be
            If VoiceTagInfo.LastIndexOf("'") <> VoiceTagInfo.Length - 1 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(VoiceTagInfo.Length - 1, 1)) ' remove ending ' char
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))  ' remove starting ' char
            ' now the VoiceTagInfo should look something like this Name=Microsoft Anna
            If VoiceTagInfo.IndexOf("Name") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 4))  ' remove starting name char
            If VoiceTagInfo.IndexOf("=") <> 0 Then Exit Function ' should not be
            VoiceTagInfo = Trim(VoiceTagInfo.Remove(0, 1))  ' remove the = char, now what is left is the voice required
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("CheckForVoiceTag returns with inText = " & inText & " and Voice = " & VoiceTagInfo, LogType.LOG_TYPE_INFO)
            Return VoiceTagInfo
        Catch ex As Exception
            Log("Error in CheckForVoiceTag called with inText = " & inText & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

#End Region

#Region "Plugin MultiZone functions"
    Public Function AllZonesOn() As Boolean
        ' 	Boolean 	Turns on all zones in the system. Returns TRUE to indicate success, FALSE indicates failure.
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("AllZonesOn called", LogType.LOG_TYPE_INFO)
        AllZonesOn = False
        Try
            If MyHSDeviceLinkedList.Count = 0 Then
                Exit Function
            End If
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.SonosPlay()
                End If
            Next
            AllZonesOn = True
        Catch ex As Exception
            Log("Error in AllZonesOn for MultizoneAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AllZonesOff() As Boolean
        '  	Boolean 	Turns off all zones in the system. Returns TRUE to indicate success, FALSE indicates failure.
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("AllZonesOff called", LogType.LOG_TYPE_INFO)
        AllZonesOff = False
        Try
            If MyHSDeviceLinkedList.Count = 0 Then
                Exit Function
            End If
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.StopPlay()
                End If
            Next
            AllZonesOff = True
        Catch ex As Exception
            Log("Error in AllZonesOff for MultizoneAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AllZonesPause() As Boolean
        '  	Boolean 	Pauses all zones in the system. Returns TRUE to indicate success, FALSE indicates failure.
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("AllZonesPause called", LogType.LOG_TYPE_INFO)
        AllZonesPause = False
        Try
            If MyHSDeviceLinkedList.Count = 0 Then
                Exit Function
            End If
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.SonosPause()
                End If
            Next
            AllZonesPause = True
        Catch ex As Exception
            Log("Error in AllZonesPause for MultizoneAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AllZonesMuteOn() As Boolean
        '  	Boolean 	Turns off all zones in the system. Returns TRUE to indicate success, FALSE indicates failure.
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("AllZonesMuteOn called", LogType.LOG_TYPE_INFO)
        AllZonesMuteOn = False
        Try
            If MyHSDeviceLinkedList.Count = 0 Then
                Exit Function
            End If
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.SonosMute()
                End If
            Next
            AllZonesMuteOn = True
        Catch ex As Exception
            Log("Error in AllZonesMuteOn for MultizoneAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function AllZonesMuteOff() As Boolean
        '  	Boolean 	Turns off all zones in the system. Returns TRUE to indicate success, FALSE indicates failure.
        If piDebuglevel > DebugLevel.dlErrorsOnly And gIOEnabled Then Log("AllZonesMuteOff called", LogType.LOG_TYPE_INFO)
        AllZonesMuteOff = False
        Try
            If MyHSDeviceLinkedList.Count = 0 Then
                Exit Function
            End If
            For Each HSDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
                If HSDevice.ZonePlayerControllerRef IsNot Nothing Then
                    HSDevice.ZonePlayerControllerRef.UnMute()
                End If
            Next
            AllZonesMuteOff = True
        Catch ex As Exception
            Log("Error in AllZonesMuteOff for MultizoneAPI with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


#End Region



#Region "Device Linked List functions"

    Private Function DMAdd() As MyUPnpDeviceInfo
        '   MyPingAddressLinkedList.AddLast(NewArrayElement)
        Dim NewArrayElement As New MyUPnpDeviceInfo
        MyHSDeviceLinkedList.AddLast(NewArrayElement)
        DMAdd = NewArrayElement
        NbrOfSonosPlayers = MyHSDeviceLinkedList.Count
    End Function

    Public Sub DMRemove(UDN As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("DMRemove called with UDN = " & UDN, LogType.LOG_TYPE_INFO)
        If MyHSDeviceLinkedList.Count = 0 Then Exit Sub
        For Each UPnPDevice As MyUPnpDeviceInfo In MyHSDeviceLinkedList
            If UPnPDevice.ZoneUDN = UDN Then
                If Not UPnPDevice.ZonePlayerControllerRef Is Nothing Then
                    'UPnPDevice.ZonePlayerControllerRef.DestroyPlayer(True)   removed in v3.1.0.25 on 9/17/2018
                    UPnPDevice.ZonePlayerControllerRef = Nothing
                End If
                UPnPDevice.Close()
                MyHSDeviceLinkedList.Remove(UPnPDevice)
                NbrOfSonosPlayers = MyHSDeviceLinkedList.Count
                Exit Sub
            End If
        Next
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DMRemove. Could not find UDN = " & UDN, LogType.LOG_TYPE_ERROR)
    End Sub
#End Region

End Class

