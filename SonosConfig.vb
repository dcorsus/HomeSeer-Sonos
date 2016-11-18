Imports Scheduler
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Text
Imports System.Web

Class SonosConfig
    Inherits clsPageBuilder

    Private PIReference As HSPI = Nothing
    Private strState As String = ""
    Private MyPageName As String = ""
    Private stb As New StringBuilder

    Private DebugChkBox As clsJQuery.jqCheckBox
    Private SuperDebugChkBox As clsJQuery.jqCheckBox
    Private UPnPDebugLevelDropList As clsJQuery.jqDropList
    Private LogErrorOnlyChkBox As clsJQuery.jqCheckBox
    Private LogToDiskChkBox As clsJQuery.jqCheckBox
    Private AutoUpdChkBox As clsJQuery.jqCheckBox
    Private ImmediateUpdChkBox As clsJQuery.jqCheckBox
    Private BuildiPodDBWHenDocked As clsJQuery.jqCheckBox
    Private AutoUpdateTimeBox As clsJQuery.jqTextBox
    Private VolumeStepBox As clsJQuery.jqTextBox
    Private SpeakerProxyCheckBox As clsJQuery.jqCheckBox
    Private PostAnnouncementActionBox As clsJQuery.jqDropList
    Private LearnRadioStationCheckBox As clsJQuery.jqCheckBox
    Private UPNPNbrOfElementsBox As clsJQuery.jqTextBox
    Private HSizeBox As clsJQuery.jqTextBox
    Private VSizeBox As clsJQuery.jqTextBox
    Private MaxAnnTimeBox As clsJQuery.jqTextBox
    Private DBZoneNameBox As clsJQuery.jqDropList
    Private AddEntryLinkTableBtn As clsJQuery.jqButton
    Private ResetPingCountersBtn As clsJQuery.jqButton
    Private HSSTrackLengthBox As clsJQuery.jqDropList
    Private HSSTrackPositionBox As clsJQuery.jqDropList
    Private MediaApiChkBox As clsJQuery.jqCheckBox


    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
        MyPageName = pagename
        DebugChkBox = New clsJQuery.jqCheckBox("DebugChkBox", " Debug Flag", MyPageName, True, False)
        DebugChkBox.toolTip = "Turn normal debug logging on or off"
        SuperDebugChkBox = New clsJQuery.jqCheckBox("SuperDebugChkBox", " Super Debug Flag", MyPageName, True, False)
        SuperDebugChkBox.toolTip = "Turn very detailed debug logging on or off. Set this on top of the normal debug log"
        UPnPDebugLevelDropList = New clsJQuery.jqDropList("UPnPDebugLvlBox", MyPageName, False)
        UPnPDebugLevelDropList.toolTip = "Set the level of debug logging for the UPnP functions"
        LogErrorOnlyChkBox = New clsJQuery.jqCheckBox("LogErrorOnlyChkBox", " Log Error Only Flag", MyPageName, True, False)
        LogErrorOnlyChkBox.toolTip = "Not implemented"
        LogToDiskChkBox = New clsJQuery.jqCheckBox("LogToDiskChkBox", " Log to Disk Flag", MyPageName, True, False)
        LogToDiskChkBox.toolTip = "Log the plug-in errors to a standard txt file. Note this slows down performance substantial. Suggested use for remote ran PIs or capture issues when terminating the PI"
        AutoUpdChkBox = New clsJQuery.jqCheckBox("AutoUpdChkBox", " Auto Update Flag", MyPageName, True, False)
        AutoUpdChkBox.toolTip = "Set this flag if you want the MusicDB to be created automatically on a daily basis. You can set the daily time in the Auto Update Time box"
        ImmediateUpdChkBox = New clsJQuery.jqCheckBox("ImmediateUpdChkBox", " Immediate Update Flag", MyPageName, True, False)
        ImmediateUpdChkBox.toolTip = "MusicDB will be re-created each time and immediately after a Sonos player DB has changed, option not recommended unless your Sonos DB hardly ever change"
        BuildiPodDBWHenDocked = New clsJQuery.jqCheckBox("BuildiPodDBWHenDocked", " Build iPod Music DB Immediately when docked", MyPageName, True, False)
        BuildiPodDBWHenDocked.toolTip = "MusicDB will be re-created each time and immediately when you plug in your iPhone/iPod, option not recommended unless your iDevice has few tracks"
        AutoUpdateTimeBox = New clsJQuery.jqTextBox("AutoUpdateTimeBox", "text", "", MyPageName, 5, False)
        AutoUpdateTimeBox.toolTip = "Specify the time when you want daily MusicDB refreshes to occur. Only applicable when Immediate Update is not set AND the Auto Update flag is set"
        VolumeStepBox = New clsJQuery.jqTextBox("VolumeStepBox", "text", "", MyPageName, 5, False)
        VolumeStepBox.toolTip = "Specify the amount of volume you want the player to increase/decrease when you click on the Volume up/down buttons on the HS devices management webpage"
        SpeakerProxyCheckBox = New clsJQuery.jqCheckBox("SpeakerProxyCheckBox", " Proxy Flag", MyPageName, True, False)
        SpeakerProxyCheckBox.toolTip = "Set this if you want the plug-in to participate in announcements"
        PostAnnouncementActionBox = New clsJQuery.jqDropList("PostAnnouncementActionBox", MyPageName, False)
        PostAnnouncementActionBox.toolTip = "Specify what to do with the announcement that was intercepted by the plugin's proxy client"
        LearnRadioStationCheckBox = New clsJQuery.jqCheckBox("LearnRadioStationCheckBox", " Learn Radiostations", MyPageName, True, False)
        LearnRadioStationCheckBox.toolTip = "Set if you want the plugin to learn/store non premium radio channel information. The learning happens when you play them"
        UPNPNbrOfElementsBox = New clsJQuery.jqTextBox("UPNPNbrOfElementsBox", "text", "", MyPageName, 5, False)
        UPNPNbrOfElementsBox.toolTip = "Set the max nbr of UPnP objects that can be retrieved in one read. For XP set to 50 for other versions of Windows less then 999"
        HSizeBox = New clsJQuery.jqTextBox("HSizeBox", "text", "", MyPageName, 5, False)
        HSizeBox.toolTip = "Set the size of the album artwork (in pixels) on the HS Device Management page"
        VSizeBox = New clsJQuery.jqTextBox("VSizeBox", "text", "", MyPageName, 5, False)
        VSizeBox.toolTip = "Set the size of the album artwork (in pixels) on the HS Device Management page"
        MaxAnnTimeBox = New clsJQuery.jqTextBox("MaxAnnTimeBox", "text", "", MyPageName, 5, False)
        MaxAnnTimeBox.toolTip = "Failsafe mechanism set in seconds. If an announcement is not completed within set time, the plugin will unlink all players and return players to original state"
        DBZoneNameBox = New clsJQuery.jqDropList("DBZoneNameBox", MyPageName, False)
        DBZoneNameBox.toolTip = "Pick the player you want to communicate with to retrieve its musicDB and store it in the plug-in musicDB. Pick one with best network connectivity!"
        AddEntryLinkTableBtn = New clsJQuery.jqButton("AddEntryLinkTableBtn", "Add Entry", MyPageName, False)
        AddEntryLinkTableBtn.toolTip = "Add a new entry to the Linkgroup table"
        ResetPingCountersBtn = New clsJQuery.jqButton("ResetPingCountersBtn", "Reset Ping Counters", MyPageName, False)
        ResetPingCountersBtn.toolTip = "Reset ALL the failed Ping counters to zero"
        HSSTrackLengthBox = New clsJQuery.jqDropList("HSSTrackLengthBox", MyPageName, False)
        HSSTrackLengthBox.toolTip = "Set the format on how you want the info to be displayed as part of the HS Device"
        HSSTrackPositionBox = New clsJQuery.jqDropList("HSSTrackPositionBox", MyPageName, False)
        HSSTrackPositionBox.toolTip = "Set the format on how you want the info to be displayed as part of the HS Device"
        MediaApiChkBox = New clsJQuery.jqCheckBox("MediaAPIChkBox", " Enable MediaAPI", MyPageName, True, False)
        MediaApiChkBox.toolTip = "Turn use of the MediaAPI in HSTouch on or off"

    End Sub

    Public WriteOnly Property RefToPlugIn
        Set(value As Object)
            PIReference = value
        End Set
    End Property


    ' build and return the actual page
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String
        If g_bDebug Then Log("GetPagePlugin for SonosControl called with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)

        Dim stb As New StringBuilder
        Dim stbLinkTable As New StringBuilder
        Dim stbPlayerTable As New StringBuilder

        Try
            Me.reset()
            DBZoneNameBox.ClearItems()

            ' handle any queries like mode=something
            Dim parts As Collections.Specialized.NameValueCollection = Nothing
            If (queryString <> "") Then
                parts = HttpUtility.ParseQueryString(queryString)
            End If

            Me.AddHeader(hs.GetPageHeader(pageName, "Sonos Configuration", "", "", False, True))
            stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd)

            ' specific page starts here

            DebugChkBox.checked = GetBooleanIniFile("Options", "Debug", False)
            SuperDebugChkBox.checked = GetBooleanIniFile("Options", "SuperDebug", False)
            Dim UPNPdeblevel As DebugLevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlOff)
            UPnPDebugLevelDropList.ClearItems()
            UPnPDebugLevelDropList.AddItem("Off", DebugLevel.dlOff, UPNPdeblevel = DebugLevel.dlOff)
            UPnPDebugLevelDropList.AddItem("Errors Only", DebugLevel.dlErrorsOnly, UPNPdeblevel = DebugLevel.dlErrorsOnly)
            UPnPDebugLevelDropList.AddItem("Events and Errors", DebugLevel.dlEvents, UPNPdeblevel = DebugLevel.dlEvents)
            UPnPDebugLevelDropList.AddItem("Verbose", DebugLevel.dlVerbose, UPNPdeblevel = DebugLevel.dlVerbose)
            LogErrorOnlyChkBox.checked = GetBooleanIniFile("Options", "LogErrorOnly", False)
            LogToDiskChkBox.checked = GetBooleanIniFile("Options", "LogToDisk", False)
            AutoUpdChkBox.checked = GetBooleanIniFile("Options", "Auto Update", False)
            ImmediateUpdChkBox.checked = GetBooleanIniFile("Options", "Immediate Auto Update", False)
            BuildiPodDBWHenDocked.checked = GetBooleanIniFile("Options", "AutoBuildDockedDB", False)
            AutoUpdateTimeBox.defaultText = GetStringIniFile("Options", "Auto Update Time", "")
            VolumeStepBox.defaultText = GetStringIniFile("Options", "VolumeStep", "")
            SpeakerProxyCheckBox.checked = GetBooleanIniFile("SpeakerProxy", "Active", False)
            LearnRadioStationCheckBox.checked = GetBooleanIniFile("Options", "Learn RadioStations", False)
            UPNPNbrOfElementsBox.defaultText = GetStringIniFile("Options", "MaxNbrofUPNPObjects", "")
            HSizeBox.defaultText = GetStringIniFile("Options", "ArtworkHSize", "")
            VSizeBox.defaultText = GetStringIniFile("Options", "ArtworkVsize", "")
            MaxAnnTimeBox.defaultText = GetStringIniFile("Options", "MaxAnnouncementTime", "")
            MediaApiChkBox.checked = GetBooleanIniFile("Options", "MediaAPIEnabled", False)

            Dim LinkGroupString As String = GetStringIniFile("LinkgroupNames", "Names", "")
            Dim ZoneNameString As String = GetStringIniFile("Sonos Zonenames", "Names", "")
            Dim ZoneUDNString As String = GetStringIniFile("Sonos Zonenames", "UDNs", "")
            Dim ZoneNames() As String
            Dim ZoneUDNs() As String
            Dim ZoneInfos As String
            Dim ZoneUDNMusicDB As String = GetStringIniFile("Options", "DB Zone", "")

            Dim ZoneNameDBIndex As Integer = -1
            ZoneNames = Split(ZoneNameString, ":|:")
            ZoneUDNs = Split(ZoneUDNString, ":|:")
            Try
                Dim ItemIndex As Integer = 0
                For Each ZoneInfos In ZoneNames
                    Dim ZoneNameInfos As String()
                    ZoneNameInfos = Split(ZoneInfos, ";:;")
                    DBZoneNameBox.AddItem(ZoneNameInfos(0), ItemIndex, True)
                    If ZoneUDNs(ItemIndex) = ZoneUDNMusicDB Then
                        ZoneNameDBIndex = ItemIndex
                    End If
                    ItemIndex += 1
                Next
                If ZoneNameDBIndex = -1 Then
                    ZoneNameDBIndex = 0
                    If ZoneUDNs IsNot Nothing Then
                        WriteStringIniFile("Options", "DB Zone", ZoneUDNs(0)) ' write first player UDN
                    End If
                End If
                DBZoneNameBox.selectedItemIndex = ZoneNameDBIndex
            Catch ex As Exception
            End Try

            Dim Names() As String
            Dim Name As String
            Dim Column As Integer = 0
            If LinkGroupString <> "" Then
                Names = Split(LinkGroupString, "|")
            Else
                Names = Nothing
            End If

            stb.Append(clsPageBuilder.DivStart("", "id='result'"))
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(clsPageBuilder.FormEnd) ' PlayerConfigForm

            stb.Append(clsPageBuilder.DivStart("HeaderPanel", "style=""color:#0000FF"" "))
            stb.Append("<h1>Sonos Plugin Configuration</h1>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)

            stb.Append("<table style='width: 80%' >")

            stb.Append("<tr><td colspan='2'>")
            stb.Append("<hr /> ")
            stb.Append("</tr></td><tr><td>")
            stb.Append(clsPageBuilder.DivStart("DebugPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Debug Information</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)


            stb.Append(DebugChkBox.Build)
            stb.Append("</br>")
            stb.Append(SuperDebugChkBox.Build)
            stb.Append("</br>")
            stb.Append(LogToDiskChkBox.Build)
            stb.Append("</br>")
            stb.Append(LogErrorOnlyChkBox.Build) ' 
            stb.Append("</br>")
            stb.Append(UPnPDebugLevelDropList.Build & " UPnP Functions Debug Level")
            stb.Append("<hr /> ")
            stb.Append(clsPageBuilder.DivStart("MusicDataBasePanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Music Database Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)

            stb.Append(AutoUpdChkBox.Build)
            stb.Append("</br>")
            stb.Append(ImmediateUpdChkBox.Build)
            stb.Append("</br>")
            stb.Append(BuildiPodDBWHenDocked.Build)
            stb.Append("</br>")
            stb.Append(AutoUpdateTimeBox.Build)
            stb.Append("Auto Update Time</br>")
            stb.Append(DBZoneNameBox.Build)
            stb.Append("DB Zonename</br>")

            stb.Append(LearnRadioStationCheckBox.Build)
            stb.Append("</br><hr />")
            stb.Append(clsPageBuilder.DivStart("UPnPSettingPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>UPnP Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(UPNPNbrOfElementsBox.Build)
            stb.Append("# of UPnP Elements<hr />")

            stb.Append(clsPageBuilder.DivStart("VolumeSettingPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Volume Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(VolumeStepBox.Build)
            stb.Append("Volume Step")
            stb.Append("</td><td>")

            stb.Append(clsPageBuilder.DivStart("SpeakerProxyPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Speaker Proxy Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(SpeakerProxyCheckBox.Build)
            stb.Append("</br>")
            Dim paaSetting = GetIntegerIniFile("Options", "PostAnnouncementAction", PostAnnouncementAction.paaForwardNoMatch)
            PostAnnouncementActionBox.ClearItems()
            PostAnnouncementActionBox.AddItem("Always Forward", PostAnnouncementAction.paaAlwaysForward, paaSetting = PostAnnouncementAction.paaAlwaysForward)
            PostAnnouncementActionBox.AddItem("Forward When No Match", PostAnnouncementAction.paaForwardNoMatch, paaSetting = PostAnnouncementAction.paaForwardNoMatch)
            PostAnnouncementActionBox.AddItem("Never Forward", PostAnnouncementAction.paaAlwaysDrop, paaSetting = PostAnnouncementAction.paaAlwaysDrop)
            stb.Append(PostAnnouncementActionBox.Build)
            stb.Append(" Post Announcement Action")
            stb.Append("</br>")

            stb.Append(MaxAnnTimeBox.Build)
            stb.Append("Maximum Announcement Time</br><hr />")

            'stb.Append(clsPageBuilder.DivStart("PingSettingsPanel", "style=""color:#0000FF"" "))
            'stb.Append("<h3>Ping Settings</h3>" & vbCrLf)
            'stb.Append(clsPageBuilder.DivEnd)
            'stb.Append(FailPingChkBox.Build)
            'stb.Append("</br>")
            'stb.Append(FailPingCountBox.Build)
            'stb.Append("# of Failing Pings<hr />")

            stb.Append(clsPageBuilder.DivStart("ArtWorkPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>Artwork Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(HSizeBox.Build)
            stb.Append("Artwort Width</br>")
            stb.Append(VSizeBox.Build)
            stb.Append("Artwork Height<hr />")

            'stb.Append(clsPageBuilder.DivStart("PeriodicFunctionsPanel", "style=""color:#0000FF"" "))
            'stb.Append("<h3>Periodic Functions</h3>" & vbCrLf)
            'stb.Append(clsPageBuilder.DivEnd)
            'stb.Append(DoRediscoveryChkBox.Build)
            'stb.Append("Don't do 5 minute rediscoveries<hr />")

            stb.Append(clsPageBuilder.DivStart("HSDeviceSettingsPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>HS3 Devices Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            Dim HSSTrackLengthBoxSetting = GetIntegerIniFile("Options", "HSSTrackLengthSetting", HSSTrackLengthSettings.TLSSeconds)
            HSSTrackLengthBox.ClearItems()
            HSSTrackLengthBox.AddItem("Seconds", HSSTrackLengthSettings.TLSSeconds, HSSTrackLengthBoxSetting = HSSTrackLengthSettings.TLSSeconds)
            HSSTrackLengthBox.AddItem("HH:MM:SS", HSSTrackLengthSettings.TLSHoursMinutesSeconds, HSSTrackLengthBoxSetting = HSSTrackLengthSettings.TLSHoursMinutesSeconds)
            stb.Append(HSSTrackLengthBox.Build & " Tracklength Format")
            stb.Append("</br>")
            Dim HSSTrackPositionBoxSetting = GetIntegerIniFile("Options", "HSSTrackPositionSetting", HSSTrackPositionSettings.TPSSeconds)
            HSSTrackPositionBox.ClearItems()
            HSSTrackPositionBox.AddItem("Seconds", HSSTrackPositionSettings.TPSSeconds, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSSeconds)
            HSSTrackPositionBox.AddItem("HH:MM:SS", HSSTrackPositionSettings.TPSHoursMinutesSeconds, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSHoursMinutesSeconds)
            HSSTrackPositionBox.AddItem("Percentage", HSSTrackPositionSettings.TPSPercentage, HSSTrackPositionBoxSetting = HSSTrackPositionSettings.TPSPercentage)
            stb.Append(HSSTrackPositionBox.Build & " Track Position Format")

            stb.Append(clsPageBuilder.DivStart("MediaAPISettingsPanel", "style=""color:#0000FF"" "))
            stb.Append("<h3>MediaAPI Settings</h3>" & vbCrLf)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append(MediaApiChkBox.Build)

            'stb.Append("</br>") ' this is to line them up on the page
            stb.Append("</td></tr></table>")

            stb.Append("<table><tr><td>")
            stb.Append("<hr /> ")

            stbLinkTable.Append("<table ID='LinkgroupTable' border='1'  style='background-color:DarkGray;color:black'>")
            stbLinkTable.Append("<tr ID='HeaderRow'  style='background-color:DarkGray;color:black'>")
            stbLinkTable.Append("<td><h3> Linkgroup Name </h3></td><td><h3> Source Player Name </h3></td><td><h3>Use Audio Input</h3></td><td><h3> Destination Player + Volume + Mute Override </h3></td><td><h3> Intercept Speaker Devices </h3></td><td><h3> Delete </h3></td></tr>")

            Dim DestGroup() As String
            Dim DestgroupString As String = ""
            Dim SourceName As String = ""
            Dim AudioInput As Boolean = False
            Dim TableRow As Integer = 0
            'Log("nbr of entries = " & UBound(Names), LogType.LOG_TYPE_INFO)

            If Not Names Is Nothing Then
                For Each Name In Names
                    'LinkGroupTable
                    stbLinkTable.Append("<tr ID='EntryRow'  style='background-color:LightGray;color:black'>")

                    Dim NewLinkGroupName As New clsJQuery.jqTextBox("LinkGroupName" & "_" & TableRow.ToString, "text", "", MyPageName, 15, False)
                    NewLinkGroupName.defaultText = Name ' PIReference.ConvertToZoneName(Name)
                    NewLinkGroupName.toolTip = "Enter your own name here"
                    stbLinkTable.Append("<td>" & NewLinkGroupName.Build & "</td>")

                    SourceName = GetStringIniFile("LinkgroupZoneSource", Name, "")
                    If SourceName <> "" Then
                        Try
                            Dim SourceInfo() As String
                            SourceInfo = Split(SourceName, ";")
                            If UBound(SourceInfo) > 0 Then
                                AudioInput = CBool(SourceInfo(1))
                            Else
                                AudioInput = False
                            End If
                            SourceName = SourceInfo(0)
                            SourceName = PIReference.ConvertToZoneName(SourceName)
                        Catch ex As Exception
                            AudioInput = False
                        End Try
                    End If

                    Dim DropList As New clsJQuery.jqDropList("SourceNameDropList" & "_" & TableRow.ToString, MyPageName, False)
                    Dim ListIndex As Integer = 0
                    Dim SourceZoneIndex As Integer = 0
                    DropList.toolTip = "Select which of your players will be used to play the announcement FROM (transmit/source) towards all destination (receiving) players"
                    Try
                        For Each ZoneInfos In ZoneNames
                            Dim ZoneNameInfos As Object
                            ZoneNameInfos = Split(ZoneInfos, ";:;")
                            DropList.AddItem(ZoneNameInfos(0), ListIndex, True)
                            If ZoneNameInfos(0) = SourceName Then
                                SourceZoneIndex = ListIndex
                            End If
                            ListIndex += 1
                        Next
                        DropList.selectedItemIndex = SourceZoneIndex
                    Catch ex As Exception
                    End Try
                    stbLinkTable.Append("<td>" & DropList.Build & "</td>")

                    Dim TTSCheckBox As New clsJQuery.jqCheckBox("TTSCheckBox" & "_" & TableRow.ToString, "", MyPageName, True, False)
                    TTSCheckBox.checked = AudioInput
                    TTSCheckBox.toolTip = "Set this if you want to use the audio input connector from the specified source zoneplayer"
                    stbLinkTable.Append("<td Align='middle'>" & TTSCheckBox.Build & "</td>")

                    stbLinkTable.Append("<td  align='left' style='white-space:nowrap;'>")
                    DestgroupString = GetStringIniFile("LinkgroupZoneDestination", Name, "")
                    If DestgroupString <> "" Then
                        DestGroup = Split(DestgroupString, "|")
                    Else
                        DestGroup = Nothing
                    End If
                    Dim DestZone As String
                    Dim CellRow As Integer = 0
                    Dim NameIndex As Integer = 0
                    For Each ZoneInfos In ZoneNames
                        Dim ZoneNameInfos As String()
                        ZoneNameInfos = Split(ZoneInfos, ";:;")
                        If ZoneNameInfos IsNot Nothing Then
                            If UBound(ZoneNameInfos, 1) > 0 Then
                                If ZoneNameInfos(1) <> "WD100" Then
                                    Dim LinkBox As New clsJQuery.jqCheckBox("LinkBox" & "_" & TableRow.ToString & "_" & CellRow.ToString, ZoneNameInfos(0), MyPageName, True, False)
                                    LinkBox.className = ""
                                    LinkBox.labelStyle = "display:inline-block;width:200px;"
                                    LinkBox.toolTip = "Set flag if you want this player to OUTPUT (play) the announcement"
                                    Dim VolumeBox As New clsJQuery.jqTextBox("VolumeBox" & "_" & TableRow.ToString & "_" & CellRow.ToString, "text", "", MyPageName, 3, False)
                                    VolumeBox.dialogCaption = "Set Volume"
                                    VolumeBox.toolTip = "Set here the volume for the announcement on this player"
                                    Dim MuteChkBox As New clsJQuery.jqCheckBox("MuteBox" & "_" & TableRow.ToString & "_" & CellRow.ToString, "Mute Override", MyPageName, True, False)
                                    MuteChkBox.toolTip = "If a player was in muted state before the announcement, specify here if you want the player to participate and override the mute state"
                                    If Not DestGroup Is Nothing Then
                                        For Each DestZone In DestGroup
                                            Dim DestInfo() As String = {""}
                                            Try
                                                DestInfo = Split(DestZone, ";")
                                            Catch ex As Exception
                                            End Try
                                            Try
                                                If DestInfo(0) = ZoneUDNs(NameIndex) Then
                                                    If UBound(DestInfo) > 2 Then
                                                        LinkBox.checked = CBool(DestInfo(3).ToString)
                                                    End If
                                                    If UBound(DestInfo) > 1 Then
                                                        ' we have mute info
                                                        MuteChkBox.checked = CBool(DestInfo(2).ToString)
                                                        ' we have a volume setting
                                                        VolumeBox.defaultText = DestInfo(1).ToString
                                                    ElseIf UBound(DestInfo) > 0 Then
                                                        ' we have a volume setting
                                                        VolumeBox.defaultText = DestInfo(1).ToString
                                                    End If
                                                    Exit For
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        Next
                                    End If
                                    stbLinkTable.Append(LinkBox.Build & VolumeBox.Build & MuteChkBox.Build & "</br>")
                                End If
                            End If                            
                        End If
                        CellRow += 1
                        NameIndex += 1
                    Next
                    stbLinkTable.Append("</td>")

                    Dim SpeakerDeviceBox As New clsJQuery.jqTextBox("SpeakerDeviceBox" & "_" & TableRow.ToString, "text", "Enter the Speaker device IDs seperated by a comma", MyPageName, 15, False)
                    SpeakerDeviceBox.toolTip = "Enter the Speaker device IDs you want intercepted, seperated by a comma"

                    Try
                        SpeakerDeviceBox.defaultText = GetStringIniFile("NewTTSSpeakDevice", Name, "")
                    Catch ex As Exception
                        Log("Error in Script looking for TTSSpeakerDevices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try

                    stbLinkTable.Append("<td Align='middle'>" & SpeakerDeviceBox.Build & "</td>")
                    stbLinkTable.Append("<td Align='middle'>" & BuildLinkgroupEntryDeleteOverlay("Are you sure?", TableRow.ToString) & "</td>")

                    stbLinkTable.Append("</tr>")
                    TableRow += 1
                Next


            End If
            stbLinkTable.Append("</table>")
            stbLinkTable.Append("</br>" & AddEntryLinkTableBtn.Build)


            stb.Append(clsPageBuilder.FormStart("LinkgroupSlideform", MyPageName, "post"))
            Dim st As New clsJQuery.jqSlidingTab("myLinkTableSlide", MyPageName, False)
            st.initiallyOpen = GetBooleanIniFile("Options", "LinkGroupSliderOpen", False)
            st.toolTip = "Open to add/change/delete Announcement Linkgroups"
            st.tab.name = "myLinkgroupSlide_name"
            st.tab.tabName.Unselected = clsPageBuilder.DivStart("LinkGrpPanel", "style=""color:#0000FF"" ") & "<h3>Linkgroup Configuration Table</h3>" & clsPageBuilder.DivEnd
            st.tab.tabName.Selected = clsPageBuilder.DivStart("LinkGrpPanel", "style=""color:#0000FF"" ") & "<h3>Linkgroup Configuration Table</h3>" & clsPageBuilder.DivEnd & "</br>" & stbLinkTable.ToString
            stb.Append(st.Build)
            stb.Append(clsPageBuilder.FormEnd)

            ' create the player table

            stbPlayerTable.Append("<table ID='PlayerListTable' border='1'  style='background-color:DarkGray;color:black'>")
            stbPlayerTable.Append("<tr ID='HeaderRow'  style='background-color:DarkGray;color:black'>")
            stbPlayerTable.Append("<td><h3> Player Name </h3></td><td><h3> Player UDN </h3></td><td><h3> Player OnLine </h3></td><td><h3> Player Model </h3></td><td><h3> Player IP Address </h3></td><td><h3> Delete </h3></td></tr>")

            Dim PlayerListTableRow As Integer = 0
            Try
                If PIReference.MyHSDeviceLinkedList.Count > 0 Then
                    For Each HSDevice As MyUPnpDeviceInfo In PIReference.MyHSDeviceLinkedList
                        If Not HSDevice Is Nothing Then
                            Dim Player As HSPI = HSDevice.ZonePlayerControllerRef
                            stbPlayerTable.Append("<tr ID='EntryRow'  style='background-color:LightGray;color:black'>")
                            stbPlayerTable.Append("<td>" & HSDevice.ZoneName & "</td>")
                            stbPlayerTable.Append("<td>" & HSDevice.ZoneUDN & "</td>")
                            Dim Alive As String = "?"
                            If MySSDPDevice IsNot Nothing Then
                                Dim DevicesList As MyUPnPDevices = MySSDPDevice.GetAllDevices()
                                If DevicesList IsNot Nothing Then
                                    Dim Device As MyUPnPDevice = DevicesList.Item("uuid:" & HSDevice.ZoneUDN, True)
                                    If Device IsNot Nothing Then
                                        If Device.Alive Then
                                            Alive = "True"
                                        Else
                                            Alive = "False"
                                        End If
                                    End If
                                End If
                            End If
                            stbPlayerTable.Append("<td>" & Alive & "</td>")
                            stbPlayerTable.Append("<td>" & HSDevice.ZoneModel & "</td>")
                            stbPlayerTable.Append("<td>" & HSDevice.UPnPDeviceIPAddress & "</td>")
                            stbPlayerTable.Append("<td Align='middle'>" & BuildPlayerDeleteOverlay("Are you sure?", PlayerListTableRow.ToString) & "</td>")
                            stbPlayerTable.Append("</tr>")
                        End If
                        PlayerListTableRow += 1
                    Next
                End If
            Catch ex As Exception
                Log("Error in Page load building the player list with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try

            stbPlayerTable.Append("</table>")
            BuildAllPlayersDeleteOverlay("Are you sure?")
            'stbPlayerTable.Append("</br>" & ResetPingCountersBtn.Build & BuildAllPlayersDeleteOverlay("Are you sure?") & RediscoverPlayersBtn.Build)


            Dim UPnPViewerBtn As New clsJQuery.jqButton("UPnPViewerBtn", "View Sonos Devices", MyPageName, False)
            UPnPViewerBtn.toolTip = "Open up the configuration page for this device"
            UPnPViewerBtn.urlNewWindow = True
            Dim HTTPPort_ As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
            If HTTPPort_ <> "" Then HTTPPort_ = ":" & HTTPPort_
            If MainInstance <> "" Then
                'UPnPViewerBtn.url = "http://" & hs.GetIPAddress & HTTPPort_ & "/" & UPnPViewPage & ":" & MainInstance
                UPnPViewerBtn.url = "/" & UPnPViewPage & ":" & MainInstance
            Else
                ' UPnPViewerBtn.url = "http://" & hs.GetIPAddress & HTTPPort_ & "/" & UPnPViewPage
                UPnPViewerBtn.url = "/" & UPnPViewPage
            End If



            stbPlayerTable.Append("</br>" & BuildAllPlayersDeleteOverlay("Are you sure?") & UPnPViewerBtn.Build)
            stb.Append("<hr /> ")
            stb.Append(clsPageBuilder.FormStart("PlayerListSlideform", MyPageName, "post"))
            Dim stpl As New clsJQuery.jqSlidingTab("myPlayerListSlide", MyPageName, False)
            stpl.initiallyOpen = GetBooleanIniFile("Options", "PlayerListSliderOpen", False)
            stpl.toolTip = "Player List"
            stpl.tab.name = "myPlayerListSlide_name"
            stpl.tab.tabName.Unselected = clsPageBuilder.DivStart("PlayerListPanel", "style=""color:#0000FF"" ") & "<h3>Player Table</h3>" & clsPageBuilder.DivEnd
            stpl.tab.tabName.Selected = clsPageBuilder.DivStart("PlayerListPanel", "style=""color:#0000FF"" ") & "<h3>Player Table</h3>" & clsPageBuilder.DivEnd & "</br>" & stbPlayerTable.ToString
            stb.Append(stpl.Build)
            stb.Append("<hr /> ")
            stb.Append(clsPageBuilder.FormEnd)

        Catch ex As Exception
            Log("Error in Page load with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' add the body html to the page
        Me.AddBody(stb.ToString)

        Me.AddFooter(hs.GetPageFooter)
        Me.suppressDefaultFooter = True

        ' return the full page
        Return Me.BuildPage()

    End Function

    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        If g_bDebug Then Log("PostBackProc for SonosControl called with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)

        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(data)

        If parts IsNot Nothing Then
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(Part, "_")
                        If g_bDebug Then Log("postBackProc for SonosControl found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                        If g_bDebug Then Log("postBackProc for SonosControl found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                        Dim ObjectValue As String = parts(Part)
                        Select Case ObjectNameParts(0).ToString
                            Case "DebugChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "Debug", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Debug flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "SuperDebugChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "SuperDebug", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving SuperDebug flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "UPnPDebugLvlBox"
                                Try
                                    WriteIntegerIniFile("Options", "UPnPDebugLevel", Val(ObjectValue))
                                Catch ex As Exception
                                    Log("Error in postBackProc for PluginControl saving UPnPDebugLevel with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                UPnPDebuglevel = GetIntegerIniFile("Options", "UPnPDebugLevel", DebugLevel.dlOff)
                            Case "LogToDiskChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "LogToDisk", ObjectValue.ToUpper = "CHECKED")
                                    gLogToDisk = GetBooleanIniFile("Options", "LogToDisk", False)
                                    If gLogToDisk Then OpenLogFile(DebugLogFileName) Else CloseLogFile()
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving LogToDisk flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LogErrorOnlyChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "LogErrorOnly", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving LogErrorOnly flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "AutoUpdChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "Auto Update", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Auto Update flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "ImmediateUpdChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "Immediate Auto Update", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Immediate Auto Update flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "BuildiPodDBWHenDocked"
                                Try
                                    WriteBooleanIniFile("Options", "AutoBuildDockedDB", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Immediate Auto Update flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "AutoUpdateTimeBox"
                                Try
                                    WriteStringIniFile("Options", "Auto Update Time", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Auto Update Time. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "GenreChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "Genres", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "TrackChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "Tracks", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "ArtistChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "Artists", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "AlbumChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "Albums", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "PlayListChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "PlayLists", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "RadioStationChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "RadioStations", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "AudioBookChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "AudioBooks", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "PodcastChkBox"
                                Try
                                    WriteBooleanIniFile("DBItems", "Podcasts", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DBItems. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "VolumeStepBox"
                                Try
                                    WriteStringIniFile("Options", "VolumeStep", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving VolumeStep. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "SpeakerProxyCheckBox"
                                Try
                                    WriteBooleanIniFile("SpeakerProxy", "Active", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Speaker Proxy flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "PostAnnouncementActionBox"
                                Try
                                    WriteIntegerIniFile("Options", "PostAnnouncementAction", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving PostAnnouncementActions. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "LearnRadioStationCheckBox"
                                Try
                                    WriteBooleanIniFile("Options", "Learn RadioStations", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Learn RadioStation flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "UPNPNbrOfElementsBox"
                                Try
                                    WriteStringIniFile("Options", "MaxNbrofUPNPObjects", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Max Nbr of UPNP Objects. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "HSizeBox"
                                Try
                                    WriteStringIniFile("Options", "ArtworkHSize", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Artwork Size Info. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "VSizeBox"
                                Try
                                    WriteStringIniFile("Options", "ArtworkVsize", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Artwork Size Info. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "MaxAnnTimeBox"
                                Try
                                    WriteStringIniFile("Options", "MaxAnnouncementTime", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving Maximum Announcement. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                                'Case "AutoUpdateZoneNameBox"
                                'Try
                                'WriteStringIniFile("Options", "Auto Update Zone", PIReference.ConvertToZoneUDN(GetZoneNameByIndex(ObjectValue)))
                                'Catch ex As Exception
                                'Log("Error in postBackProc for SonosControl saving Auto Update Zone. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                'End Try
                            Case "DBZoneNameBox"
                                Try
                                    WriteStringIniFile("Options", "DB Zone", PIReference.ConvertToZoneUDN(GetZoneNameByIndex(ObjectValue)))
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving DB Zone. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "myLinkTableSlide"
                                If ObjectValue = "myLinkgroupSlide_name_open" Then
                                    'Log("postBackProc has open slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "LinkGroupSliderOpen", True)
                                ElseIf ObjectValue = "myLinkgroupSlide_name_close" Then
                                    'Log("postBackProc has closed slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "LinkGroupSliderOpen", False)
                                End If
                            Case "myPlayerListSlide"
                                If ObjectValue = "myPlayerListSlide_name_open" Then
                                    Log("postBackProc has open slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "PlayerListSliderOpen", True)
                                ElseIf ObjectValue = "myPlayerListSlide_name_close" Then
                                    Log("postBackProc has closed slider", LogType.LOG_TYPE_INFO)
                                    WriteBooleanIniFile("Options", "PlayerListSliderOpen", False)
                                End If
                            Case "AddEntryLinkTableBtn"
                                ItemChange(LinkTableItems.ltiAddBtn, ObjectValue, 0)
                                Me.pageCommands.Add("refresh", "true")
                            Case "LinkGroupName"
                                ItemChange(LinkTableItems.liLinkgroupName, ObjectValue, (ObjectNameParts(1)))
                            Case "SourceNameDropList"
                                ' ObjectNameParts(1).ToString holds the row starting with 0
                                ItemChange(LinkTableItems.ltiSourceName, ObjectValue, (ObjectNameParts(1)))
                            Case "LinkBox"
                                ItemChange(LinkTableItems.ltiLinkBox, ObjectValue, (ObjectNameParts(1)), (ObjectNameParts(2)))
                            Case "VolumeBox"
                                ItemChange(LinkTableItems.ltiVolumeBox, ObjectValue, (ObjectNameParts(1)), (ObjectNameParts(2)))
                            Case "MuteBox"
                                ItemChange(LinkTableItems.ltiMuteBox, ObjectValue, (ObjectNameParts(1)), (ObjectNameParts(2)))
                            Case "TTSCheckBox"
                                ItemChange(LinkTableItems.ltiTTS, ObjectValue, (ObjectNameParts(1)))
                            Case "SpeakerDeviceBox"
                                ItemChange(LinkTableItems.ltiSpeakerDevice, ObjectValue, (ObjectNameParts(1)))
                            Case "DeleteBtn"
                                ItemChange(LinkTableItems.ltiDeleteBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovLnkTblSubmit"
                                ItemChange(LinkTableItems.ltiDeleteBtn, ObjectValue, (ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovPlayerDelSubmit"
                                DeletePlayerClick((ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovDelAllSubmit"
                                DeleteAllPlayersClick()
                                Me.pageCommands.Add("refresh", "true")
                            Case "ovConfirmCancel", "ovLnkTblCancel", "ovPlayerDelCancel", "ovDelAllCancel"
                                Me.pageCommands.Add("refresh", "true")
                            Case "DeletePlayerBtn"
                                DeletePlayerClick((ObjectNameParts(1)))
                                Me.pageCommands.Add("refresh", "true")
                                'Case "DeleteAllPlayersBtn"
                                'DeleteAllPlayersClick()
                                'Me.pageCommands.Add("refresh", "true")
                            Case "ResetPingCountersBtn"
                                ResetPingCountersClick()
                                Me.pageCommands.Add("refresh", "true")
                            Case "HSSTrackLengthBox"
                                Try
                                    WriteIntegerIniFile("Options", "HSSTrackLengthSetting", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving HS TrackLength Settings. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "HSSTrackPositionBox"
                                Try
                                    WriteIntegerIniFile("Options", "HSSTrackPositionSetting", ObjectValue)
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving HS TrackPosition Settings. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                            Case "MediaAPIChkBox"
                                Try
                                    WriteBooleanIniFile("Options", "MediaAPIEnabled", ObjectValue.ToUpper = "CHECKED")
                                Catch ex As Exception
                                    Log("Error in postBackProc for SonosControl saving the MediaAPI flag. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                End Try
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in postBackProc for SonosConfig processing with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            ' call the plug-in to have it read and apply the new settings
            Try
                PIReference.ReadIniFile()
                PIReference.ReadLinkgroups()
            Catch ex As Exception
                Log("Error in postBackProc for SonosControl calling Plugin to update values with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If g_bDebug Then Log("postBackProc for SonosControl found parts to be empty", LogType.LOG_TYPE_INFO)
        End If

        Return MyBase.postBackProc(page, data, user, userRights)
    End Function

    Public Enum LinkTableItems
        ltiSourceName = 0
        ltiLinkBox = 1
        ltiVolumeBox = 2
        ltiMuteBox = 3
        ltiTTS = 4
        ltiSpeakerDevice = 5
        ltiDeleteBtn = 6
        ltiAddBtn = 7
        liLinkgroupName = 8
    End Enum

    Public Sub ItemChange(LinkTableItem As LinkTableItems, Value As String, RowIndex As Integer, Optional CellIndex As Integer = 0)
        If g_bDebug Then Log("ItemChange called with LinbktableItem = " & LinkTableItem.ToString & " and Value = " & Value.ToString & " and RowIndex = " & RowIndex.ToString & " and CellIndex = " & CellIndex.ToString, LogType.LOG_TYPE_INFO)
        Value = Trim(Value)
        Dim KeyValue As New System.Collections.Generic.KeyValuePair(Of String, String)
        Dim LinkgroupName As String = ""
        Try
            Dim LinkGroupString = GetStringIniFile("LinkgroupNames", "Names", "")
            Dim Names As String() = Split(LinkGroupString, "|")
            LinkgroupName = Names(RowIndex)
        Catch ex As Exception
            Log("Error in ItemChange with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Try
            Select Case LinkTableItem
                Case LinkTableItems.liLinkgroupName
                    ' need to update sections [LinkgroupNames] [LinkgroupZoneSource] [LinkgroupZoneDestination]
                    ' we should force a reset of the page after this field is updated
                    If Value = "" Then
                        Log("Warning in ItemChange for LinkgroupName is empty, no change will be made", LogType.LOG_TYPE_WARNING)
                        Exit Sub
                    End If
                    Dim LinkGroupStr As String = GetStringIniFile("LinkgroupNames", "Names", "")
                    Dim LinkGroupNames As String() = Split(LinkGroupStr, "|")
                    Dim NewLinkGroupStr As String = ""
                    For Each LinkgrpName In LinkGroupNames
                        If LinkgrpName = Value Then
                            Log("Warning in ItemChange for LinkgroupName. The new linkgroup name already exists, no change will be made", LogType.LOG_TYPE_WARNING)
                            Exit Sub
                        End If
                    Next
                    For Each LinkgrpName As String In LinkGroupNames
                        If NewLinkGroupStr <> "" Then NewLinkGroupStr = NewLinkGroupStr & "|"
                        If LinkgrpName = LinkgroupName Then
                            NewLinkGroupStr = NewLinkGroupStr & Value
                        Else
                            NewLinkGroupStr = NewLinkGroupStr & LinkgrpName
                        End If
                    Next
                    WriteStringIniFile("LinkgroupNames", "Names", NewLinkGroupStr)

                    ' now update the [LinkgroupZoneSource] section
                    Dim NewSourceInfo As New System.Collections.Generic.Dictionary(Of String, String)()
                    NewSourceInfo = GetIniSection("LinkgroupZoneSource")
                    If Not NewSourceInfo Is Nothing Then
                        DeleteIniSection("LinkgroupZoneSource")
                        For Each KeyValue In NewSourceInfo
                            If KeyValue.Key = LinkgroupName Then
                                WriteStringIniFile("LinkgroupZoneSource", Value, KeyValue.Value)
                            Else
                                WriteStringIniFile("LinkgroupZoneSource", KeyValue.Key, KeyValue.Value)
                            End If
                        Next
                    End If


                    ' now update the [LinkgroupZoneDestination] section
                    Dim NewDestinationInfo As New System.Collections.Generic.Dictionary(Of String, String)()
                    NewDestinationInfo = GetIniSection("LinkgroupZoneDestination")
                    If Not NewDestinationInfo Is Nothing Then
                        DeleteIniSection("LinkgroupZoneDestination")
                        For Each KeyValue In NewDestinationInfo
                            If KeyValue.Key = LinkgroupName Then
                                WriteStringIniFile("LinkgroupZoneDestination", Value, KeyValue.Value)
                            Else
                                WriteStringIniFile("LinkgroupZoneDestination", KeyValue.Key, KeyValue.Value)
                            End If
                        Next
                    End If


                    ' need to update section [TTSSpeakDevice]
                    Dim NewTTSSpeakDeviceInfo As New System.Collections.Generic.Dictionary(Of String, String)()
                    NewTTSSpeakDeviceInfo = GetIniSection("NewTTSSpeakDevice")
                    If Not NewTTSSpeakDeviceInfo Is Nothing Then
                        DeleteIniSection("NewTTSSpeakDevice")
                        For Each KeyValue In NewTTSSpeakDeviceInfo
                            If KeyValue.Key = LinkgroupName Then
                                WriteStringIniFile("NewTTSSpeakDevice", Value, KeyValue.Value)
                            Else
                                WriteStringIniFile("NewTTSSpeakDevice", KeyValue.Key, KeyValue.Value)
                            End If
                        Next
                    End If


                Case LinkTableItems.ltiSourceName
                    ' value holds the index into the Sonos Zone Names
                    Dim ZoneUDN As String = GetZoneUDNByIndex(Val(Value))
                    Dim OldSourceInfoString As String = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
                    Dim Infos As String() = Split(OldSourceInfoString, ";")
                    Dim TTSFlag As Integer = 0
                    If UBound(Infos) > 0 Then
                        TTSFlag = Infos(1)
                    End If
                    ' now update the [LinkgroupZoneSource] section
                    WriteStringIniFile("LinkgroupZoneSource", LinkgroupName, ZoneUDN & ";" & TTSFlag)

                Case LinkTableItems.ltiLinkBox
                    Dim ZoneUDN As String = GetZoneUDNByIndex(CellIndex)
                    Dim LinkgroupDestinationInfoString As String = GetStringIniFile("LinkgroupZoneDestination", LinkgroupName, "")
                    Dim NewLinkgroupDestinationInfoString As String = ""
                    Dim DestGroup As String() = Split(LinkgroupDestinationInfoString, "|")
                    Dim Found As Boolean = False
                    Dim Checked As Integer = 0
                    If Value.ToUpper = "CHECKED" Then Checked = 1
                    If LinkgroupDestinationInfoString <> "" Then
                        For Each DestZone In DestGroup
                            'Log("ItemChange found DestZone = " & DestZone, LogType.LOG_TYPE_INFO)
                            If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                            Dim DestInfo() As String = Split(DestZone, ";")
                            'Log("ItemChange found DestInfo(0) = " & DestInfo(0) & " and ZoneUDN = " & ZoneUDN, LogType.LOG_TYPE_INFO)
                            If DestInfo(0) = ZoneUDN Then
                                Found = True
                                If UBound(DestInfo) < 3 Then
                                    ReDim Preserve DestInfo(3)
                                End If
                                DestInfo(3) = Checked.ToString
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestInfo(0) & ";" & DestInfo(1) & ";" & DestInfo(2) & ";" & DestInfo(3)
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            Else
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestZone.ToString
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            End If
                        Next
                    End If
                    If Not Found And (Checked <> 0) Then
                        ' add entry
                        If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                        NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & ZoneUDN & ";;0;" & Checked.ToString
                        'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                    End If
                    WriteStringIniFile("LinkgroupZoneDestination", LinkgroupName, NewLinkgroupDestinationInfoString)
                Case LinkTableItems.ltiVolumeBox
                    Dim ZoneUDN As String = GetZoneUDNByIndex(CellIndex)
                    Dim LinkgroupDestinationInfoString As String = GetStringIniFile("LinkgroupZoneDestination", LinkgroupName, "")
                    Dim NewLinkgroupDestinationInfoString As String = ""
                    Dim DestGroup As String() = Split(LinkgroupDestinationInfoString, "|")
                    Dim Found As Boolean = False
                    If LinkgroupDestinationInfoString <> "" Then
                        For Each DestZone In DestGroup
                            'Log("ItemChange found DestZone = " & DestZone, LogType.LOG_TYPE_INFO)
                            If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                            Dim DestInfo() As String = Split(DestZone, ";")
                            ' Log("ItemChange found DestInfo(0) = " & DestInfo(0) & " and ZoneUDN = " & ZoneUDN, LogType.LOG_TYPE_INFO)
                            If DestInfo(0) = ZoneUDN Then
                                Found = True
                                If UBound(DestInfo) < 3 Then
                                    ReDim Preserve DestInfo(3)
                                    DestInfo(3) = "1"
                                End If
                                DestInfo(1) = Value
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestInfo(0) & ";" & DestInfo(1) & ";" & DestInfo(2) & ";" & DestInfo(3)
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            Else
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestZone.ToString
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            End If
                        Next
                    End If
                    If Not Found Then
                        ' add entry
                        If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                        NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & ZoneUDN & ";" & Value & ";0;0"
                        'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                    End If
                    WriteStringIniFile("LinkgroupZoneDestination", LinkgroupName, NewLinkgroupDestinationInfoString)
                Case LinkTableItems.ltiMuteBox
                    Dim ZoneUDN As String = GetZoneUDNByIndex(CellIndex)
                    Dim LinkgroupDestinationInfoString As String = GetStringIniFile("LinkgroupZoneDestination", LinkgroupName, "")
                    Dim NewLinkgroupDestinationInfoString As String = ""
                    Dim DestGroup As String() = Split(LinkgroupDestinationInfoString, "|")
                    Dim Found As Boolean = False
                    Dim Checked As Integer = 0
                    If Value.ToUpper = "CHECKED" Then Checked = 1
                    If LinkgroupDestinationInfoString <> "" Then
                        For Each DestZone In DestGroup
                            'Log("ItemChange found DestZone = " & DestZone, LogType.LOG_TYPE_INFO)
                            If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                            Dim DestInfo() As String = Split(DestZone, ";")
                            'Log("ItemChange found DestInfo(0) = " & DestInfo(0) & " and ZoneUDN = " & ZoneUDN, LogType.LOG_TYPE_INFO)
                            If DestInfo(0) = ZoneUDN Then
                                Found = True
                                If UBound(DestInfo) < 3 Then
                                    ReDim Preserve DestInfo(3)
                                    DestInfo(3) = "1"
                                End If
                                DestInfo(2) = Checked.ToString
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestInfo(0) & ";" & DestInfo(1) & ";" & DestInfo(2) & ";" & DestInfo(3)
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            Else
                                NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & DestZone.ToString
                                'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                            End If
                        Next
                    End If
                    If Not Found Then
                        ' add entry
                        If NewLinkgroupDestinationInfoString <> "" Then NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & "|"
                        NewLinkgroupDestinationInfoString = NewLinkgroupDestinationInfoString & ZoneUDN & ";;" & Checked.ToString & ";0"
                        'Log("ItemChange new linkinfo = " & NewLinkgroupDestinationInfoString, LogType.LOG_TYPE_INFO)
                    End If
                    WriteStringIniFile("LinkgroupZoneDestination", LinkgroupName, NewLinkgroupDestinationInfoString)
                Case LinkTableItems.ltiTTS
                    ' value holds the index into the Sonos Zone Names
                    Dim ZoneSourceString As String = GetStringIniFile("LinkgroupZoneSource", LinkgroupName, "")
                    Dim ZoneUDN As String = ""
                    Try
                        Dim ZoneSourceParts As String() = Split(ZoneSourceString, ";")
                        ZoneUDN = ZoneSourceParts(0)
                    Catch ex As Exception

                    End Try

                    Dim Checked As Integer = 0
                    If Value.ToUpper = "CHECKED" Then Checked = 1
                    ' now update the [LinkgroupZoneSource] section
                    WriteStringIniFile("LinkgroupZoneSource", LinkgroupName, ZoneUDN & ";" & Checked.ToString)
                    'Log("ItemChange new SourceInfo = " & ZoneUDN & ";" & Checked.ToString, LogType.LOG_TYPE_INFO)
                Case LinkTableItems.ltiSpeakerDevice
                    Log("ItemChange is updating TTSSpeakDevice LingroupName = " & LinkgroupName & " and Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
                    Try
                        If Value <> "" Then
                            WriteStringIniFile("NewTTSSpeakDevice", LinkgroupName, Value)
                        Else
                            DeleteEntryIniFile("NewTTSSpeakDevice", LinkgroupName)
                        End If
                    Catch ex As Exception
                        Log("Error in Script looking for NewTTSSpeakerDevices with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try


                Case LinkTableItems.ltiDeleteBtn
                    ' need to update sections [LinkgroupNames] 
                    Dim LinkGroupStr As String = GetStringIniFile("LinkgroupNames", "Names", "")
                    Dim LinkGroupNames As String() = Split(LinkGroupStr, "|")
                    Dim NewLinkGroupStr As String = ""

                    For Each LinkgrpName As String In LinkGroupNames
                        If LinkgrpName <> LinkgroupName Then
                            If NewLinkGroupStr <> "" Then NewLinkGroupStr = NewLinkGroupStr & "|"
                            NewLinkGroupStr = NewLinkGroupStr & LinkgrpName
                        End If
                    Next
                    WriteStringIniFile("LinkgroupNames", "Names", NewLinkGroupStr)
                    ' need to update sections [LinkgroupZoneSource] 
                    DeleteEntryIniFile("LinkgroupZoneSource", LinkgroupName)
                    ' need to update sections [LinkgroupZoneDestination]
                    DeleteEntryIniFile("LinkgroupZoneDestination", LinkgroupName)
                    ' need to update section [NewTTSSpeakDevice]
                    DeleteEntryIniFile("NewTTSSpeakDevice", LinkgroupName)
                Case LinkTableItems.ltiAddBtn
                    ' need to update sections [LinkgroupNames] [LinkgroupZoneSource] [LinkgroupZoneDestination]
                    Dim LinkGroupStr As String = GetStringIniFile("LinkgroupNames", "Names", "")
                    Dim NewLinkgroupName As String = ""
                    If LinkGroupStr <> "" Then
                        Dim LinkGroupNames As String() = Split(LinkGroupStr, "|")
                        NewLinkgroupName = "NewTableEntry" & (UBound(LinkGroupNames) + 2).ToString
                        LinkGroupStr = LinkGroupStr & "|" & NewLinkgroupName
                    Else
                        NewLinkgroupName = "NewTableEntry1"
                        LinkGroupStr = NewLinkgroupName
                    End If
                    Dim ZoneUDN As String = GetZoneUDNByIndex(0)    ' get the first ZoneName
                    WriteStringIniFile("LinkgroupNames", "Names", LinkGroupStr)
                    WriteStringIniFile("LinkgroupZoneSource", NewLinkgroupName, ZoneUDN & ";0")
                    WriteStringIniFile("LinkgroupZoneDestination", NewLinkgroupName, "")
            End Select

        Catch ex As Exception
            Log("Error in ItemChange case selector with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeletePlayerClick(PlayerTableItem As Integer)
        If g_bDebug Then Log("DeletePlayerClick called with tableItem = " & PlayerTableItem.ToString, LogType.LOG_TYPE_INFO)
        PlayerTableItem = Trim(PlayerTableItem)
        ' go find the UDN
        Dim IndexCount As Integer = 0
        If PIReference.MyHSDeviceLinkedList.Count > 0 Then
            For Each HSDevice As MyUPnpDeviceInfo In PIReference.MyHSDeviceLinkedList
                If Not HSDevice Is Nothing And (IndexCount = PlayerTableItem) Then
                    Dim Player As HSPI = HSDevice.ZonePlayerControllerRef
                    Dim UDN As String = Player.GetUDN
                    PIReference.DeleteDevice(UDN)
                    If MainInstance <> "" Then
                        RemoveInstance(MainInstance & "-" & UDN)
                    Else
                        RemoveInstance(UDN)
                    End If
                    'PIReference.DeleteDevice(UDN)
                    Exit Sub
                End If
                IndexCount += 1
            Next
        End If
    End Sub

    Public Sub DeleteAllPlayersClick()
        If g_bDebug Then Log("DeleteAllPlayersClick called", LogType.LOG_TYPE_INFO)
        If PIReference.MyHSDeviceLinkedList.Count > 0 Then
            For Each HSDevice As MyUPnpDeviceInfo In PIReference.MyHSDeviceLinkedList
                If Not HSDevice Is Nothing Then
                    Dim Player As HSPI = HSDevice.ZonePlayerControllerRef
                    If Player IsNot Nothing Then
                        Dim DeviceUDN As String = Player.GetUDN
                        Player.DeleteWebLink(DeviceUDN, Player.ZonePlayerName)
                        If MainInstance <> "" Then
                            RemoveInstance(MainInstance & "-" & DeviceUDN)
                        Else
                            RemoveInstance(DeviceUDN)
                        End If
                    End If
                    DeleteIniSection(HSDevice.ZoneUDN)
                End If
            Next
        End If
        PIReference.MyHSDeviceLinkedList.Clear()
        DeleteIniSection("UPnP Devices UDN to Info")
        DeleteIniSection("UPnP HSRef to UDN")
        DeleteIniSection("UPnP UDN to HSRef")
        DeleteIniSection("Sonos Zonenames")
        DeleteIniSection("SonosMusicPage")
        WriteIntegerIniFile("Settings", "MasterHSDeviceRef", -1) ' reset the master code, next time Sonos initializes everything will be cleaned out.
        hs.DeleteIODevices(sIFACE_NAME, MainInstance)
        ' Force HomeSeer to save changes to devices and events so we can find our new device
        hs.SaveEventsDevices()
        MyShutDownRequest = True
    End Sub


    Public Sub ResetPingCountersClick()
        If g_bDebug Then Log("ResetPingCountersClick called", LogType.LOG_TYPE_INFO)
        If PIReference.MyHSDeviceLinkedList.Count > 0 Then
            For Each HSDevice As MyUPnpDeviceInfo In PIReference.MyHSDeviceLinkedList
                If Not HSDevice Is Nothing Then
                    Dim Player As HSPI = HSDevice.ZonePlayerControllerRef
                    Player.MissedPings = 0
                End If
            Next
        End If
    End Sub


    Private Function GetZoneNameByIndex(index As Integer) As String
        GetZoneNameByIndex = ""
        If g_bDebug Then Log("GetZoneNameByIndex called with Index = " & index.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim ZoneNameString As String = GetStringIniFile("Sonos Zonenames", "Names", "")
            Dim ZoneNames As String() = Split(ZoneNameString, ":|:")
            Dim ZoneInfos As String = ZoneNames(index)
            Dim ZoneNameInfos As String() = Split(ZoneInfos, ";:;")
            If g_bDebug Then Log("GetZoneNameByIndex called with Index = " & index.ToString & " and found Name = " & ZoneNameInfos(0), LogType.LOG_TYPE_INFO)
            Return ZoneNameInfos(0)
        Catch ex As Exception
            Log("Error in GetZoneNameByIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Function GetZoneUDNByIndex(index As Integer) As String
        GetZoneUDNByIndex = ""
        If g_bDebug Then Log("GetZoneUDNByIndex called with Index = " & index.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim ZoneUDNString As String = GetStringIniFile("Sonos Zonenames", "UDNs", "")
            Dim ZoneUDNs As String() = Split(ZoneUDNString, ":|:")
            Dim ZoneUDN As String = ZoneUDNs(index)
            If g_bDebug Then Log("GetZoneUDNByIndex called with Index = " & index.ToString & " and found UDN = " & ZoneUDN, LogType.LOG_TYPE_INFO)
            Return ZoneUDN
        Catch ex As Exception
            Log("Error in GetZoneUDNByIndex with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function



    Private Function GetHeadContent() As String
        Try
            Return hs.GetPageHeader(sIFACE_NAME, "", "", False, False, True, False, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function GetFooterContent() As String
        Try
            Return hs.GetPageFooter(False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function GetBodyContent() As String
        Try
            Return hs.GetPageHeader(StrConv(sIFACE_NAME, VbStrConv.ProperCase), "", "", False, False, False, True, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Private Function BuildLinkgroupEntryDeleteOverlay(HeaderText As String, ButtonSuffix As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("LinkTblOverlay" & ButtonSuffix.ToString, MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "Delete this Linkgroup Entry"
        ConfirmOverlay.label = "Delete"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("LinkTableConfirmform" & ButtonSuffix.ToString, MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovLnkTblSubmit_" & ButtonSuffix, "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovLnkTblCancel_" & ButtonSuffix, "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildPlayerDeleteOverlay(HeaderText As String, ButtonSuffix As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("PlayerDelOverlay" & ButtonSuffix.ToString, MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "Delete this Player"
        ConfirmOverlay.label = "Delete"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("PlayerDeleteConfirmform" & ButtonSuffix.ToString, MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovPlayerDelSubmit_" & ButtonSuffix, "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovPlayerDelCancel_" & ButtonSuffix, "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

    Private Function BuildAllPlayersDeleteOverlay(HeaderText As String) As String
        Dim ConfirmOverlay As New clsJQuery.jqOverlay("DellAllOverlay", MyPageName, False, "events_overlay")
        ConfirmOverlay.toolTip = "This will delete all players, including the root device. Upon completion the plugin will terminate and will have to be restarted manually or will be restarted by HS. Note any players on the network when the PI is restarted will be added again to HS"
        ConfirmOverlay.label = "Delete All Players"
        ConfirmOverlay.overlayHTML = clsPageBuilder.FormStart("DelAllPlayersConfirmform", MyPageName, "post")
        ConfirmOverlay.overlayHTML &= "<div>" & HeaderText & "<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovDelAllSubmit", "Submit", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovDelAllCancel", "Cancel", MyPageName, False)
        ConfirmOverlay.overlayHTML &= tbut2.Build & "<br /><br />"
        ConfirmOverlay.overlayHTML &= "</div>"
        ConfirmOverlay.overlayHTML &= clsPageBuilder.FormEnd
        Return ConfirmOverlay.Build
    End Function

End Class
