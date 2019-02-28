Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.IO.Path
Imports System.Drawing
Imports System.Xml
Imports System.Data.SQLite


<Serializable()> _
Partial Public Class HSPI
    Implements IMediaAPI_3


    Public Enum player_status_change
        SongChanged = 1         'raises whenever the current song changes
        PlayStatusChanged = 2   'raises when pause, stop, play, etc. pressed.
        PlayList = 3            'raises whenever the current playlist changes
        Library = 4             'raises when the library changes
        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        AlarmStart = 12          ' raised when the alarm goes off
        ConfigChange = 13        ' raised when the configuration of a device changes like alarm info being modified
        NextSong = 14            ' raised when the next song is about to start
        AnnouncementChange = 15           ' raised when the next song is about to start
    End Enum

    Public Enum player_state_values
        Playing = 1
        Stopped = 2
        Paused = 3
        Forwarding = 4
        Rewinding = 5
        Docked = 11         ' new state to support WD100 devices being docked
        Undocked = 12       ' new state to support WD100 devices being undocked
        AudioInTrue = 13    ' new state to indicating Line-In went Up (or connected)
        AudioInFalse = 14   ' new state to indicating Line-In went down (or disconnected)
        ZoneName = 15
        ReplicationState = 16
        UpdateHSServerOnly = 17
        Online = 18
        Offline = 19
        AnnouncementStart = 20
        AnnouncementStop = 21
    End Enum

    Public Enum repeat_modes
        repeat_off = 0
        repeat_one = 1
        repeat_all = 2
    End Enum

    Public Enum Shuffle_modes
        Shuffled = 1
        Ordered = 2
        Sorted = 3
    End Enum

    Public Enum QueueActions
        qaDontPlay
        qaPlayLast
        qaPlayNext
        qaPlayNow
    End Enum

    Public Enum MyLibraryTypes
        LibraryQueue = 1
        LibraryDB = 2
    End Enum

    Private TrackDescriptor As New track_desc
    Private MyCurrentTrack As String = ""
    Private MyCurrentArtist As String = ""
    Private MyCurrentAlbum As String = ""
    Private MyRadiostationName As String = ""
    Private MyNextTrack As String = ""
    Private MyNextArtist As String = ""
    Private MyNextAlbum As String = ""
    Private MyNextAlbumURI As String = NoArtPath
    Private MyCurrentPlayerState As player_state_values = player_state_values.Stopped
    Private MyCurrentArtworkURL As String = NoArtPath
    'Private MyVolume As Integer = 0
    Private MyMuteState As Boolean
    Private MyZoneSource As String = ""
    Private MyZoneSourceExt As String = ""
    Private MyCurrentURIMetaData As String = ""
    Private MyNextURIMetaData As String = ""
    Private MyPlayerPosition As Integer = -1
    Private MyTrackLength As Integer = 0
    Private LastRetrieveAlbumArtURL As String = ""
    Friend WithEvents MyMusicAPITimer As Timers.Timer
    Friend WithEvents MyZonePlayerTimer As Timers.Timer
    Private MyShuffleState As String = "NORMAL"
    Private MyPreviousShuffleState As String = "NORMAL"
    Private MyMusicService As String = ""
    Private ConnectToIPod As Boolean = False
    Private ConnectPlayer As Boolean = False
    Private MyDockediPodPlayerName As String = ""
    Private MyPlayerWentThroughPlayState As Boolean = False
    Private MyNbrOfTracksInQueue As Integer = 0
    Private MyTrackInQueueNbr As Integer = 0
    Public MyQueueHasChanged As Boolean = False
    Private MyPreviousAnnouncementPosition As Integer = 0
    Private MyAnnouncementPositionHasNotChangedNbrofTimes As Integer = 0
    Const MaxAllowedAnnouncementPositionHasNotChanged = 6
    Private CurrentLibEntry As New Lib_Entry
    Private NextLibEntry As New Lib_Entry
    Private CurrentLibKey As New Lib_Entry_Key
    Private NextLibKey As New Lib_Entry_Key

    Const MaxPlayerTOActionArray = 10
    '
    ' Timeout indexes
    Const TOReachable = 0
    Const TOPositionUpdate = 1
    '
    ' Timeout Values
    Const TOReachableValue = 10
    Const ToPositionUpdateValue = 1


    Public Sub InitMusicAPI()
        MyMusicAPITimer = New Timers.Timer
        MyMusicAPITimer.Interval = 1000 ' one second
        MyMusicAPITimer.AutoReset = True
        MyMusicAPITimer.Enabled = True
        Dim Index As Integer
        For Index = 0 To MaxPlayerTOActionArray
            MyPlayerTimeoutActionArray(Index) = 0
        Next
        MyZonePlayerTimer = New Timers.Timer
        MyZonePlayerTimer.Interval = 20000 ' 20 seconds '10000 ' 10 seconds
        MyZonePlayerTimer.AutoReset = True
        MyZonePlayerTimer.Enabled = True
        MyPlayerTimeoutActionArray(TOReachable) = TOReachableValue
        MyPlayerTimeoutActionArray(TOPositionUpdate) = ToPositionUpdateValue

        With CurrentLibKey
            .iKey = 1
            .Library = MyLibraryTypes.LibraryQueue
            .sKey = "1"
            .Title = ""
            .WhichKey = eKey_Type.eEither
        End With

        With NextLibKey
            .iKey = 1
            .Library = MyLibraryTypes.LibraryQueue
            .sKey = "1"
            .Title = ""
            .WhichKey = eKey_Type.eEither
        End With

        With CurrentLibEntry
            .Title = ""
            .Album = ""
            .Artist = ""
            .Cover_path = NoArtPath
            .Cover_Back_path = NoArtPath
            .Genre = ""
            .Key = CurrentLibKey
            .Kind = ""
            .LengthSeconds = 0
            .Lib_Media_Type = eLib_Media_Type.Music
            .Lib_Type = 1
            .PlayedCount = 0
            .Rating = 0
            .Year = 0
        End With

        With NextLibEntry
            .Title = ""
            .Album = ""
            .Artist = ""
            .Cover_path = NoArtPath
            .Cover_Back_path = NoArtPath
            .Genre = ""
            .Key = NextLibKey
            .Kind = ""
            .LengthSeconds = 0
            .Lib_Media_Type = eLib_Media_Type.Music
            .Lib_Type = 1
            .PlayedCount = 0
            .Rating = 0
            .Year = 0
        End With

        DeviceStatus = "Offline"
    End Sub

    Public Sub DestroyPlayer(Disposing As Boolean)
        Try
            If g_bDebug Then Log("Dispose called for zoneplayer = " & ZoneName & " and Disposing = " & Disposing.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        If Disposing Then
            ' Free other state (managed objects).
            Try
                CurrentLibEntry.Key = Nothing
                NextLibEntry.Key = Nothing
                CurrentLibEntry = Nothing
                NextLibEntry = Nothing
            Catch ex As Exception
            End Try
            Try
                MyMusicAPITimer.Enabled = False
                MyZonePlayerTimer.Enabled = False
                MyMusicAPITimer = Nothing
                MyZonePlayerTimer = Nothing
                myAVTransportCallback = Nothing
                myRenderingControlCallback = Nothing
                myContentDirectoryCallback = Nothing
                myAudioInCallback = Nothing
                myDevicePropertiesCallback = Nothing
                myAlarmClockCallback = Nothing
                myMusicServicesCallback = Nothing
                mySystemPropertiesCallback = Nothing
                myZonegroupTopologyCallback = Nothing
                myGroupManagementCallback = Nothing
                myConnectionManagerCallback = Nothing
                myQueueServiceCallback = Nothing
                myVirtualLineInCallBack = Nothing
                MediaServer = Nothing
                MediaRenderer = Nothing
                AudioIn = Nothing
                DeviceProperties = Nothing
                AVTransport = Nothing
                RenderingControl = Nothing
                ContentDirectory = Nothing
                AlarmClock = Nothing
                ZoneGroupTopology = Nothing
                MusicServices = Nothing
                QueueService = Nothing
                VirtualLineIn = Nothing
                MyHSPIControllerRef = Nothing
                MyWirelessDockSourcePlayer = Nothing
                MyWirelessDockDestinationPlayer = Nothing
                MyUPnPDevice = Nothing
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub MyMusicAPITimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyMusicAPITimer.Elapsed
        Dim Index As Integer
        'If g_bDebug Then Log( "MyMusicAPITimer_Elapsed called for zone " & ZoneName)
        For Index = 0 To MaxPlayerTOActionArray
            If MyPlayerTimeoutActionArray(Index) <> 0 Then
                Select Case Index
                    Case TOPositionUpdate
                        MyPlayerTimeoutActionArray(Index) = MyPlayerTimeoutActionArray(Index) - 1
                        If MyPlayerTimeoutActionArray(Index) <= 0 Then
                            UpdatePositionInfo()
                            MyPlayerTimeoutActionArray(Index) = ToPositionUpdateValue
                        End If
                End Select
            End If
        Next
        e = Nothing
        sender = Nothing
    End Sub

    Private Sub UpdatePositionInfo()
        If hs Is Nothing Then Exit Sub ' We're existing and HS is already disconnected
        Dim NewPlayerPosition As Integer = MyPlayerPosition
        'If g_bDebug Then Log("UpdatePositionInfo for ZoneName = " & ZoneName & " and ZoneSource = " & MyZoneSource & ", My ExtZoneSource = " & MyZoneSourceExt & ", Playerstate = " & MyCurrentPlayerState.ToString & ", PlayerPosition = " & MyPlayerPosition.ToString, LogType.LOG_TYPE_WARNING)
        If MyCurrentPlayerState = player_state_values.Playing Then
            If Not (MyZoneSourceExt.IndexOf("Stream Radio") = 0 Or MyZoneSourceExt.IndexOf("TV") = 0) Or NewPlayerPosition = -1 Then ' -1 is only at startup
                ' OK I'm going to fake it here a bit
                NewPlayerPosition = NewPlayerPosition + 1
            End If
        ElseIf MyCurrentPlayerState = player_state_values.Stopped Then
            NewPlayerPosition = 0
        End If
        If NewPlayerPosition = MyPlayerPosition Then Exit Sub
        MyPlayerPosition = NewPlayerPosition
        If HSRefTrackPos <> -1 Then
            Select Case MyHSTrackPositionFormat
                Case HSSTrackPositionSettings.TPSSeconds
                    hs.SetDeviceString(HSRefTrackPos, MyPlayerPosition.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, MyPlayerPosition, True)
                Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                    hs.SetDeviceString(HSRefTrackPos, ConvertSecondsToTimeFormat(MyPlayerPosition), True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, MyPlayerPosition, True)
                Case HSSTrackPositionSettings.TPSPercentage
                    Dim TrackPos As Integer = 0
                    If MyTrackLength <> 0 Then
                        TrackPos = MyPlayerPosition / MyTrackLength * 100
                    End If
                    hs.SetDeviceString(HSRefTrackPos, TrackPos.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, TrackPos, True)
            End Select
        End If
    End Sub

    Public Property Track As String
        Get
            Track = MyCurrentTrack
            'If g_bDebug And gIOEnabled Then Log( "Track Get for ZoneName = " & ZoneName & ". Track = " & MyCurrentTrack)
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property  ' We're existing and HS is already disconnected
            If MyCurrentTrack <> value And HSRefTrack <> -1 Then
                hs.SetDeviceString(HSRefTrack, value, True)
            End If
            MyCurrentTrack = value
            CurrentLibEntry.Title = value
            CurrentLibKey.Title = value
            If SuperDebug And gIOEnabled Then Log("Track Set for ZoneName = " & ZoneName & ". Track = " & MyCurrentTrack, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextTrack As String
        Get
            'Returns the name of the next track.
            'If g_bDebug Then Log( "NextTrack called for Zone - " & ZoneName & ". Value= " & MyNextTrack.ToString)
            NextTrack = MyNextTrack
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyNextTrack <> value And HSRefNextTrack <> -1 Then
                hs.SetDeviceString(HSRefNextTrack, value, True)
            End If
            MyNextTrack = value
            NextLibEntry.Title = value
            NextLibKey.Title = value
            If SuperDebug And gIOEnabled Then Log("NextTrack Set for ZoneName = " & ZoneName & ". NextTrack = " & MyNextTrack, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property Artist As String
        Get
            Artist = MyCurrentArtist
            'If g_bDebug And gIOEnabled Then Log( "Artist Get for ZoneName = " & ZoneName & ". Artist = " & MyCurrentArtist)
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyCurrentArtist <> value And HSRefArtist <> -1 Then
                hs.SetDeviceString(HSRefArtist, value, True)
            End If
            MyCurrentArtist = value
            CurrentLibEntry.Artist = value
            If SuperDebug And gIOEnabled Then Log("Artist Set for ZoneName = " & ZoneName & ". Artist = " & MyCurrentArtist, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextArtist As String
        Get
            'Returns the artist's name of the next track.
            NextArtist = MyNextArtist
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyNextArtist <> value And HSRefNextArtist <> -1 Then
                hs.SetDeviceString(HSRefNextArtist, value, True)
            End If
            MyNextArtist = value
            NextLibEntry.Artist = value
            If SuperDebug And gIOEnabled Then Log("NextArtist Set for ZoneName = " & ZoneName & ". NextArtist = " & MyNextArtist, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property Album As String
        Get
            Album = MyCurrentAlbum
            'If g_bDebug And gIOEnabled Then Log( "Album Get for ZoneName = " & ZoneName & ". Album = " & MyCurrentAlbum)
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyCurrentAlbum <> value And HSRefAlbum <> -1 Then
                hs.SetDeviceString(HSRefAlbum, value, True)
            End If
            MyCurrentAlbum = value
            CurrentLibEntry.Album = value
            If SuperDebug And gIOEnabled Then Log("Album Set for ZoneName = " & ZoneName & ". Album = " & MyCurrentAlbum, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextAlbum As String
        Get
            'Returns the album name of the next track.
            NextAlbum = MyNextAlbum
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyNextAlbum <> value And HSRefNextAlbum <> -1 Then
                hs.SetDeviceString(HSRefNextAlbum, value, True)
            End If
            MyNextAlbum = value
            NextLibEntry.Album = value
            If SuperDebug And gIOEnabled Then Log("NextAlbum Set for ZoneName = " & ZoneName & ". NextAlbum = " & MyNextAlbum, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property RadiostationName As String
        Get
            RadiostationName = MyRadiostationName
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyRadiostationName <> value And HSRefRadiostationName <> -1 Then
                hs.SetDeviceString(HSRefRadiostationName, value, True)
            End If
            MyRadiostationName = value
            If SuperDebug And gIOEnabled Then Log("RadiostationName Set for ZoneName = " & ZoneName & ". RadiostationName = " & MyRadiostationName, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public WriteOnly Property SetVolume As Integer
        Set(ByVal value As Integer)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyCurrentMasterVolumeLevel <> value Then
                If HSRefVolume <> -1 Then
                    If SuperDebug Then Log("SetVolume is setting HS Status Volume for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefVolume.ToString, LogType.LOG_TYPE_INFO)
                    hs.SetDeviceValueByRef(HSRefVolume, value, True)
                    If MyCurrentMasterVolumeLevel > value Then
                        DeviceTrigger("Sonos Volume Down")
                    ElseIf MyCurrentMasterVolumeLevel < value Then
                        DeviceTrigger("Sonos Volume Up")
                    End If
                End If
                If MyWirelessDockSourcePlayer IsNot Nothing Then
                    MyWirelessDockSourcePlayer.SetVolume = value
                End If
                If PlayBarMaster Then UpdatePlaybarSlaves("Volume", value)
            End If
            MyCurrentMasterVolumeLevel = value
        End Set
    End Property

    Public WriteOnly Property SetMuteState As Boolean
        ' needed to update mute state between linked zones
        Set(ByVal value As Boolean)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyCurrentMuteState <> value Then
                If HSRefMute <> -1 Then
                    If SuperDebug Then Log("SetMuteState is setting HS Status Mute for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefMute.ToString, LogType.LOG_TYPE_INFO)
                    If value Then
                        'hs.SetDeviceString(HSRefMute, "Muted", True)
                        hs.SetDeviceValueByRef(HSRefMute, msMuted, True)
                    Else
                        'hs.SetDeviceString(HSRefMute, "Unmuted", True)
                        hs.SetDeviceValueByRef(HSRefMute, msUnmuted, True)
                    End If
                End If
                If MyWirelessDockSourcePlayer IsNot Nothing Then
                    MyWirelessDockSourcePlayer.SetMuteState = value
                End If
                If PlayBarMaster Then UpdatePlaybarSlaves("Mute", value)
            End If
            MyCurrentMuteState = value
        End Set
    End Property

    Public WriteOnly Property SetLoudness As Boolean
        Set(value As Boolean)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If value <> MyCurrentLoudnessState Then
                If HSRefLoudness <> -1 Then
                    If SuperDebug Then Log("SetLoudnessState is setting HS Status Loundness for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefVolume.ToString, LogType.LOG_TYPE_INFO)
                    If value Then
                        hs.SetDeviceValueByRef(HSRefLoudness, lsLoudnessOn, True)
                        'hs.SetDeviceString(HSRefLoudness, "Loudness On", True)
                    Else
                        hs.SetDeviceValueByRef(HSRefLoudness, lsLoudnessOff, True)
                        'hs.SetDeviceString(HSRefLoudness, "Loudness Off", True)
                    End If
                End If
                If MyWirelessDockSourcePlayer IsNot Nothing Then
                    MyWirelessDockSourcePlayer.SetLoudness = value
                End If
                If PlayBarMaster Then UpdatePlaybarSlaves("Loudness", value)
                MyCurrentLoudnessState = value
            End If
        End Set
    End Property

    Public Property CurrentPlayerState As player_state_values
        Get
            CurrentPlayerState = MyCurrentPlayerState
            'If g_bDebug And gIOEnabled Then Log( "CurrentPlayerState Get for ZoneName = " & ZoneName & ". Value = " & MyCurrentPlayerState.ToString)
            'Log( "CurrentPlayerState Get for ZoneName = " & ZoneName & ". Value = " & MyCurrentPlayerState.ToString)
        End Get
        Set(ByVal value As player_state_values)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If MyCurrentPlayerState <> value And HSRefPlayState <> -1 Then
                If SuperDebug Then Log("CurrentPlayerState is setting HS Status for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefPlayState.ToString, LogType.LOG_TYPE_INFO)
                hs.SetDeviceValueByRef(HSRefPlayState, value, True)
            End If
            MyCurrentPlayerState = value
            If g_bDebug And gIOEnabled Then Log("CurrentPlayerState Set for ZoneName = " & ZoneName & ". Value = " & MyCurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public WriteOnly Property LinkedCurrentPlayerState As player_state_values
        Set(ByVal value As player_state_values)
            If MyCurrentPlayerState <> value And HSRefPlayState <> -1 Then
                If SuperDebug Then Log("LinkedCurrentPlayerState is setting HS Status for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefPlayState.ToString, LogType.LOG_TYPE_INFO)
                hs.SetDeviceValueByRef(HSRefPlayState, value, True)
            End If
            PlayChangeNotifyCallback(player_status_change.SongChanged, value) ' notify HS if they have the callback linked
            If value <> MyCurrentPlayerState Then PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, value)
            MyCurrentPlayerState = value
            If MyCurrentPlayerState = player_state_values.Playing Then MyPlayerWentThroughPlayState = True
            If g_bDebug And gIOEnabled Then Log("LinkedCurrentPlayerState Set for ZoneName = " & ZoneName & ". Value = " & MyCurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Private Sub UpdatePlaybarSlaves(WhatDevice As String, NewValue As Object)
        If g_bDebug Then Log("UpdatePlaybarSlaves called ZoneName - " & ZoneName & " for " & WhatDevice & " and new value = " & NewValue.ToString, LogType.LOG_TYPE_INFO)
        If MyTargetZoneLinkedList = "" Then Exit Sub
        Dim TargetZones As String() = Split(MyTargetZoneLinkedList, ";")
        If TargetZones Is Nothing Then Exit Sub
        For Each TargetZone As String In TargetZones
            If TargetZone <> "" Then
                Dim UpdatePlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(TargetZone)
                If UpdatePlayer IsNot Nothing Then
                    Select Case WhatDevice
                        Case "Volume"
                            UpdatePlayer.SetVolume = NewValue
                        Case "Loudness"
                            UpdatePlayer.SetLoudness = NewValue
                        Case "Mute"
                            UpdatePlayer.SetMuteState = NewValue
                        Case "Balance"
                            UpdatePlayer.SetBalance(NewValue)
                    End Select
                End If
            End If

        Next
    End Sub

    Public Property ArtworkURL As String
        Get
            ArtworkURL = MyCurrentArtworkURL
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If Not (value.ToLower().StartsWith("http://") Or value.ToLower().StartsWith("https://") Or value.ToLower().StartsWith("file:")) Then
                Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                'value = "http://" & hs.GetIPAddress & HTTPPort & value
            End If
            If MyCurrentArtworkURL <> value And HSRefArt <> -1 Then
                If SuperDebug Then Log("ArtworkURL is setting HS Status for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefArt.ToString, LogType.LOG_TYPE_INFO)
                hs.DeviceVSP_ClearAll(HSRefArt, True)
                hs.DeviceVGP_ClearAll(HSRefArt, True)
                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = 1000
                Pair.Status = value
                hs.DeviceVSP_AddPair(HSRefArt, Pair)
                Dim GraphicsPair As VGPair
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = value
                GraphicsPair.Set_Value = 1000
                hs.DeviceVGP_AddPair(HSRefArt, GraphicsPair)
                hs.SetDeviceValueByRef(HSRefArt, 1000, True)
                hs.SetDeviceString(HSRefArt, value, True)
            End If
            MyCurrentArtworkURL = value
            CurrentLibEntry.Cover_path = value
            If SuperDebug And gIOEnabled Then Log("ArtworkURL Set for ZoneName = " & ZoneName & ". Album = " & MyCurrentArtworkURL, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public Property NextArtworkURL As String
        Get
            NextArtworkURL = MyNextAlbumURI ' MyNextAlbumURI
        End Get
        Set(ByVal value As String)
            If hs Is Nothing Then Exit Property ' We're existing and HS is already disconnected
            If Not (value.ToLower().StartsWith("http://") Or value.ToLower().StartsWith("https://") Or value.ToLower().StartsWith("file:")) Then
                Dim HTTPPort As String = hs.GetINISetting("Settings", "gWebSvrPort", "")
                If HTTPPort <> "" Then HTTPPort = ":" & HTTPPort
                'value = "http://" & hs.GetIPAddress & HTTPPort & value
            End If
            If MyNextAlbumURI <> value And HSRefNextArt <> -1 Then
                If SuperDebug Then Log("NextArtworkURL is setting HS Status for ZoneName - " & ZoneName & " with Value = " & value.ToString & " and HSRef = " & HSRefNextArt.ToString, LogType.LOG_TYPE_INFO)
                hs.DeviceVSP_ClearAll(HSRefNextArt, True)
                hs.DeviceVGP_ClearAll(HSRefNextArt, True)
                Dim Pair As VSPair
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = 1000
                Pair.Status = value
                hs.DeviceVSP_AddPair(HSRefNextArt, Pair)
                Dim GraphicsPair As VGPair
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = value
                GraphicsPair.Set_Value = 1000
                hs.DeviceVGP_AddPair(HSRefNextArt, GraphicsPair)
                hs.SetDeviceValueByRef(HSRefNextArt, 1000, True)
                hs.SetDeviceString(HSRefNextArt, value, True)
            End If
            MyNextAlbumURI = value
            NextLibEntry.Cover_path = value
            If SuperDebug And gIOEnabled Then Log("NextArtworkURL Set for ZoneName = " & ZoneName & ". URL = " & MyNextAlbumURI, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property PlayerIconURL As String
        Get
            PlayerIconURL = MyIconURL
        End Get
    End Property

    Public Property ZoneSource As String
        Get
            ZoneSource = MyZoneSource
        End Get
        Set(ByVal value As String)
            MyZoneSource = value
        End Set
    End Property

    Public ReadOnly Property ZoneSourceExt As String
        Get
            ZoneSourceExt = MyZoneSourceExt
        End Get
    End Property

    Public ReadOnly Property CurrentTrackMetaData As String
        Get
            CurrentTrackMetaData = MyCurrentURIMetaData
        End Get
    End Property

    Public ReadOnly Property NextTrackMetaData As String
        Get
            NextTrackMetaData = MyNextURIMetaData
        End Get
    End Property

    Public ReadOnly Property EnqueuedTransportURI As String
        Get
            EnqueuedTransportURI = MyEnqueuedTransportURI
        End Get
    End Property

    Public ReadOnly Property EnqueuedTransportURIMetaData As String
        Get
            EnqueuedTransportURIMetaData = MyEnqueuedTransportURIMetaData
        End Get
    End Property

    Public ReadOnly Property ZonePlayerName As String
        Get
            ZonePlayerName = ZoneName
        End Get
    End Property

    Public ReadOnly Property InstanceName As String
        Get
            InstanceName = ZoneName
        End Get
    End Property

    Public Property DockediPodPlayerName As String
        Get
            DockediPodPlayerName = MyDockediPodPlayerName
        End Get
        Set(ByVal value As String)
            MyDockediPodPlayerName = value
        End Set
    End Property

    Public ReadOnly Property LineInputConnected As Boolean
        Get
            LineInputConnected = MyLineInputConnected
        End Get
    End Property


    Public Property SonosPlayerPosition As Integer 'As Integer 'Implements MediaCommon.MusicAPI.PlayerPosition
        Get ' The position of the player in the current track expressed as seconds.
            ' the HST plugin calls this a lot, I think it uses it to figure out when tracks have changed
            'If g_bDebug And gIOEnabled Then Log( "Get PlayerPosition called for Zone - " & ZoneName & ". ZoneSource= " & MyZoneSource)
            SonosPlayerPosition = MyPlayerPosition ' this is automatically updated by the own timer
            'If g_bDebug Then Log( "Get PlayerPosition called for Zone - " & ZoneName & ". Value = " & PlayerPosition.ToString)
            'Log( "Get PlayerPosition called for Zone - " & ZoneName & ". Value = " & PlayerPosition.ToString)
        End Get
        Set(ByVal value As Integer) ' Sets the position of the player in the current track - parameter is expressed as seconds.         
            If g_bDebug Then Log("Set SonosPlayerPosition called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Dim Time As String
            Time = ConvertSecondsToTimeFormat(value)
            If Not (MyZoneSource = "Linked" Or MyZoneSource = "Input") Then
                ' this won't work and will generate an error 
                SeekTime(Time)
            End If
        End Set
    End Property

    Public Sub SetPlayerPosition(ByVal Position As Integer)
        ' I call this in response to updates received from the controllers. I cannot call "PlayerPosition because then I create a loop, instruction SONOS to go 
        ' to where it just reported it was
        If MyPlayerPosition <> Position And HSRefTrackPos <> -1 Then
            Select Case MyHSTrackPositionFormat
                Case HSSTrackPositionSettings.TPSSeconds
                    hs.SetDeviceString(HSRefTrackPos, Position.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, Position, True)
                Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                    hs.SetDeviceString(HSRefTrackPos, ConvertSecondsToTimeFormat(Position), True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, Position, True)
                Case HSSTrackPositionSettings.TPSPercentage
                    Dim TrackPos As Integer = 0
                    If MyTrackLength <> 0 Then
                        TrackPos = MyPlayerPosition / MyTrackLength * 100
                    End If
                    hs.SetDeviceString(HSRefTrackPos, TrackPos.ToString, True)
                    hs.SetDeviceValueByRef(HSRefTrackPos, TrackPos, True)
            End Select
        End If
        MyPlayerPosition = Position
    End Sub

    Public Property MusicService As String
        Get
            MusicService = MyMusicService
        End Get
        Set(ByVal value As String)
            MyMusicService = value
        End Set
    End Property

    Public Sub SetTrackLength(ByVal TrackLength As Integer)
        If MyTrackLength <> TrackLength Then
            If HSRefTrackLength <> -1 Then
                If g_bDebug Then Log("SetTrackLength is setting HS Device TrackLength for ZoneName - " & ZoneName & " with Value = " & TrackLength.ToString & " and HSRef = " & HSRefTrackLength.ToString, LogType.LOG_TYPE_INFO)
                Select Case MyHSTrackLengthFormat
                    Case HSSTrackLengthSettings.TLSSeconds
                        hs.SetDeviceString(HSRefTrackLength, TrackLength.ToString, True)
                    Case HSSTrackLengthSettings.TLSHoursMinutesSeconds
                        hs.SetDeviceString(HSRefTrackLength, ConvertSecondsToTimeFormat(TrackLength), True)
                End Select
                hs.SetDeviceValueByRef(HSRefTrackLength, TrackLength, True)
                If HSRefTrackPos <> -1 Then
                    If g_bDebug Then Log("SetTrackLength is setting HS Status Max Position TrackPosition for ZoneName - " & ZoneName & " and HSRef = " & HSRefTrackPos.ToString, LogType.LOG_TYPE_INFO)
                    ' update the slider control pair
                    Dim VSVGPair As VSPair
                    VSVGPair = hs.DeviceVSP_Get(HSRefTrackPos, 0, ePairStatusControl.Both) ' use value = 0 to be within the range
                    VSVGPair.Render_Location.Column = 2
                    VSVGPair.Render_Location.Row = 1
                    If VSVGPair.Render = HomeSeerAPI.Enums.CAPIControlType.ValuesRangeSlider Then
                        If g_bDebug Then Log("SetTrackLength set Pair for ZoneName - " & ZoneName & " Old Max Range = " & VSVGPair.RangeEnd.ToString & " New Max Range = " & TrackLength.ToString, LogType.LOG_TYPE_INFO)
                        If MyHSTrackPositionFormat <> HSSTrackPositionSettings.TPSPercentage Then
                            VSVGPair.RangeEnd = TrackLength
                        Else
                            VSVGPair.RangeEnd = 100
                        End If
                        hs.DeviceVSP_ClearAny(HSRefTrackPos, 0)
                        hs.DeviceVSP_AddPair(HSRefTrackPos, VSVGPair)
                    End If
                End If
            End If
            If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' Update player zone
                MyWirelessDockDestinationPlayer.SetTrackLength(TrackLength)
            End If
        End If
        MyTrackLength = TrackLength
        CurrentLibEntry.LengthSeconds = TrackLength
    End Sub

    Public Function PlayerState() As player_state_values 'Implements MediaCommon.MusicAPI.PlayerState
        'Returns the current state of the player using the following Enum values:
        'Public Enum player_state_values
        '   playing = 1
        '   stopped = 2
        '   paused = 3
        '   forwarding = 4
        '   rewinding = 5
        'End Enum
        'If g_bDebug Then Log( "PlayerState called for Zone - " & ZoneName & ". State = " & MyCurrentPlayerState.ToString)
        PlayerState = MyCurrentPlayerState
        'Log( "PlayerState called for Zone - " & ZoneName & ". State = " & MyCurrentPlayerState.ToString)
    End Function

    Public ReadOnly Property CurrentStreamTitle As String 'Implements MediaCommon.MusicAPI.CurrentStreamTitle
        Get
            'The title of the currently playing music stream (e.g. from an Internet music source)
            'If g_bDebug Then Log( "CurrentStreamTitle called for Zone - " & ZoneName & "Value= " & MyCurrentTrack.ToString)
            CurrentStreamTitle = MyCurrentTrack
            'Log( "CurrentStreamTitle called for Zone - " & ZoneName & "Value= " & MyCurrentTrack.ToString)
        End Get
    End Property

    Public ReadOnly Property CurrentArtworkFile(Optional ByVal sPath As String = "") As String 'Implements MediaCommon.MusicAPI.CurrentArtworkFile
        Get
            ' implementation is slightly different and probably reverse. I return path here wheras in CurrentAlbumArtPath I download file and return path
            CurrentArtworkFile = MyCurrentArtworkURL
            'If g_bDebug Then Log( "CurrentArtworkFile called for Zone - " & ZoneName & "Value= " & MyCurrentArtworkURL.ToString)
        End Get
    End Property

    Public ReadOnly Property CurrentAlbumArtPath As String
        Get
            CurrentAlbumArtPath = MyCurrentArtworkURL
        End Get
    End Property

    Public ReadOnly Property CurrentTrack As String 'Implements IMediaAPI_3.
        Get
            'Returns the name of the currently playing track.
            'If g_bDebug Then Log( "CurrentTrack called for Zone - " & ZoneName & ". Value= " & MyCurrentTrack.ToString)
            CurrentTrack = MyCurrentTrack
            'Log( "CurrentTrack called for Zone - " & ZoneName & ". Value= " & MyCurrentTrack.ToString)
        End Get
    End Property

    Public ReadOnly Property CurrentAlbum As String 'Implements MediaCommon.MusicAPI.CurrentAlbum
        Get
            'Returns the album name of the currently playing track.
            'If g_bDebug Then Log( "CurrentAlbum called for Zone - " & ZoneName & "Value= " & MyCurrentAlbum.ToString)
            CurrentAlbum = MyCurrentAlbum
        End Get
    End Property

    Public ReadOnly Property CurrentArtist As String 'Implements MediaCommon.MusicAPI.CurrentArtist
        Get
            'Returns the artist's name of the currently playing track.
            'If g_bDebug Then Log( "CurrentArtist called for Zone - " & ZoneName & "Value= " & MyCurrentArtist.ToString)
            CurrentArtist = MyCurrentArtist
        End Get
    End Property

    Public ReadOnly Property CurrentPlayMode As String
        Get
            'If g_bDebug Then Log( "CurrentPlayMode called for Zone - " & ZoneName & "Value= " & MyShuffleState.ToString)
            CurrentPlayMode = MyShuffleState
        End Get
    End Property

    Public ReadOnly Property CurrentTrackDuration As String
        Get
            CurrentTrackDuration = MyTrackLength
        End Get
    End Property

    Public ReadOnly Property CurrentTrackDescription As String
        Get
            CurrentTrackDescription = ""
        End Get
    End Property

    Public Property HasAnnouncementStarted As Boolean
        Get
            HasAnnouncementStarted = MyPlayerWentThroughPlayState
        End Get
        Set(value As Boolean)
            MyPlayerWentThroughPlayState = value
            If Not value Then ' the HasAnnouncementStarted flag is being reset, also reset the position monitoring functions added in V3.0.0.18
                MyPreviousAnnouncementPosition = 0
                MyAnnouncementPositionHasNotChangedNbrofTimes = 0
            End If
        End Set
    End Property

    Private Sub UpdateBalance()
        If MyLeftVolume < 100 Then
            MyBalance = 100 - MyLeftVolume
        ElseIf MyRightVolume < 100 Then
            MyBalance = -100 + MyRightVolume
        Else
            MyBalance = 0
        End If
        If HSRefBalance <> -1 Then
            If g_bDebug Then Log("UpdateBalance is setting HS Status for ZoneName - " & ZoneName & " with Value = " & MyBalance.ToString & " and HSRef = " & HSRefBalance.ToString, LogType.LOG_TYPE_INFO)
            'hs.SetDeviceString(HSRefBalance, MyBalance.ToString, True)
            hs.SetDeviceValueByRef(HSRefBalance, MyBalance, True)
        End If
        If MyWirelessDockSourcePlayer IsNot Nothing Then
            MyWirelessDockSourcePlayer.SetHSBalance = MyBalance
        End If
    End Sub

    Public WriteOnly Property SetHSBalance As Integer ' used to update the WD100 HS device
        Set(value As Integer)
            If HSRefBalance <> -1 Then
                If g_bDebug Then Log("SetHSBalance is setting HS Status for ZoneName - " & ZoneName & " with Value = " & MyBalance.ToString & " and HSRef = " & HSRefBalance.ToString, LogType.LOG_TYPE_INFO)
                hs.SetDeviceValueByRef(HSRefBalance, value, True)
            End If
        End Set
    End Property

    Private Sub SetBalance(inBalance As Integer)
        If g_bDebug Then Log("SetBalance called for ZoneName - " & ZoneName & " with Value = " & inBalance.ToString, LogType.LOG_TYPE_INFO)
        If inBalance < -100 Or inBalance > 100 Then Exit Sub
        If inBalance < 0 Then
            If MyLeftVolume <> 100 Then
                SetVolumeLevel("LF", 100)
            End If
            If MyRightVolume <> 100 + inBalance Then
                SetVolumeLevel("RF", 100 + inBalance)
            End If
        ElseIf inBalance > 0 Then
            If MyRightVolume <> 100 Then
                SetVolumeLevel("RF", 100)
            End If
            If MyLeftVolume <> 100 - inBalance Then
                SetVolumeLevel("LF", 100 - inBalance)
            End If
        Else
            If MyLeftVolume <> 100 Then
                SetVolumeLevel("LF", 100)
            End If
            If MyRightVolume <> 100 Then
                SetVolumeLevel("RF", 100)
            End If
        End If
    End Sub


    Public Sub DeviceTrigger(ByVal triggerEvent As String)
        If g_bDebug Then Log("DeviceTrigger called for Zone - " & ZoneName & " with Trigger = " & triggerEvent.ToString, LogType.LOG_TYPE_INFO)
        Dim TrigsToCheck() As IPlugInAPI.strTrigActInfo
        'Dim TC As IPlugInAPI.strTrigActInfo

        Dim strTrig As strTrigger = Nothing
        Try
            If MainInstance <> "" Then
                TrigsToCheck = callback.GetTriggersInst(sIFACE_NAME, MainInstance)
            Else
                TrigsToCheck = callback.GetTriggers(sIFACE_NAME)
            End If
        Catch ex As Exception
            TrigsToCheck = Nothing
        End Try
        If TrigsToCheck IsNot Nothing AndAlso TrigsToCheck.Count > 0 Then
            For Each TC As IPlugInAPI.strTrigActInfo In TrigsToCheck
                If SuperDebug Then Log("DeviceTrigger found Trigger: EvRef=" & TC.evRef.ToString & ", Trig/SubTrig=" & TC.TANumber.ToString & "/" & TC.SubTANumber.ToString & ", UID=" & TC.UID.ToString, LogType.LOG_TYPE_INFO)
                'Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = Nothing
                'TrigsToCheck = callback.TriggerMatches(sIFACE_NAME, Info.TANumber, Info.SubTANumber)
                'callback.TriggerFire(IFACE_NAME, Info)
                If Not (TC.DataIn Is Nothing) Then
                    Dim trigger As New trigger
                    DeSerializeObject(TC.DataIn, trigger)
                    Dim Command As String = ""
                    Dim PlayerUDN As String = ""
                    Dim Linkgroup As String = ""
                    For Each sKey In trigger.Keys
                        'If g_bDebug Then Log("TriggerConfigured found sKey = " & sKey.ToString & " and Value = " & trigger(sKey), LogType.LOG_TYPE_INFO)
                        Select Case True
                            Case InStr(sKey, "PlayerListTrigger") > 0 AndAlso trigger(sKey) <> ""
                                If trigger(sKey) <> UDN Then
                                    Exit For ' not for this player
                                End If
                                PlayerUDN = trigger(sKey)
                                If Command <> "" Then
                                    callback.TriggerFire(sIFACE_NAME, TC)
                                    If g_bDebug Then Log("DeviceTrigger called TriggerFire for Zone - " & ZoneName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                    Exit For
                                End If
                            Case InStr(sKey, "CommandListTrigger") > 0 AndAlso trigger(sKey) <> ""
                                'If g_bDebug Then Log("DeviceTrigger for Zone - " & ZoneName & " command = " & trigger(sKey) & " while looking for triggerEvent = " & triggerEvent, LogType.LOG_TYPE_INFO)
                                If trigger(sKey) <> triggerEvent Then
                                    Exit For ' not right state
                                End If
                                Command = trigger(sKey)
                                If PlayerUDN <> "" Then
                                    callback.TriggerFire(sIFACE_NAME, TC)
                                    If g_bDebug Then Log("DeviceTrigger called TriggerFire for Zone - " & ZoneName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                    Exit For
                                ElseIf Linkgroup <> "" Then
                                    If AnnouncementLink IsNot Nothing Then
                                        If Linkgroup <> AnnouncementLink.LinkGroupName Then
                                            Exit For
                                        End If
                                        callback.TriggerFire(sIFACE_NAME, TC)
                                        If g_bDebug Then Log("DeviceTrigger called TriggerFire for Zone - " & ZoneName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                        Exit For
                                    Else
                                        Exit For
                                    End If
                                End If
                            Case InStr(sKey, "LinkgroupListTrigger") > 0 AndAlso trigger(sKey) <> ""
                                Linkgroup = trigger(sKey)
                                If AnnouncementLink IsNot Nothing Then
                                    If Linkgroup <> AnnouncementLink.LinkGroupName Then
                                        Exit For
                                    End If
                                Else
                                    Exit For
                                End If
                                If Command <> "" Then
                                    callback.TriggerFire(sIFACE_NAME, TC)
                                    If g_bDebug Then Log("DeviceTrigger called TriggerFire for Zone - " & ZoneName & " with Trigger = " & Command, LogType.LOG_TYPE_INFO)
                                    Exit For
                                End If
                        End Select
                    Next
                End If
            Next
        End If
    End Sub



    Public Sub PlayChangeNotifyCallback(ByVal ChangeType As player_status_change, ByVal ChangeValue As player_state_values, Optional SendDeviceTrigger As Boolean = True)
        ' Raised by a Music plug-in whenever various music plug-in status changes.  
        'HSTouch and other applications can add an event handler for this event in your plug-in to be notified of changes in the status.
        ' Public Enum player_status_change
        '        SongChanged = 1           raised whenever the current song changes
        '        PlayStatusChanged = 2     raised when pause, stop, play, etc. pressed.
        '        PlayList = 3              raised whenever the current playlist changes
        '        Library = 4               raised when the library changes
        '        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        '        AlarmStart = 12          ' raised when the alarm goes off
        '        ConfigChange = 13        ' raised when the configuration of a device changes like alarm info being modified
        '        NextSong = 14
        '    End Enum
        If g_bDebug Then Log("PlayChangeNotifyCallback called for Zone - " & ZoneName & " with ChangeType = " & ChangeType.ToString & " and Changevalue = " & ChangeValue.ToString & " and SendDeviceTrigger = " & SendDeviceTrigger.ToString, LogType.LOG_TYPE_INFO)
        'Log( "PlayChangeNotifyCallback called for Zone - " & ZoneName & " with ChangeType = " & ChangeType.ToString & " and Changevalue = " & ChangeValue.ToString)
        If gInterfaceStatus <> ERR_NONE Then
            'If g_bDebug Then Log( "Warning PlayChangeNotifyCallback called for Zone - " & ZoneName & " before plugin is initialized. Nothing sent")
            Exit Sub ' no updates to be sent until completely intialized. Else the multizone API is hosed.
        End If
        Dim Parms(2)
        Dim TriggerEvent As String = ""
        ' trigger Names
        '   Track Change
        '   Player Stop
        '   Player Paused
        '   "Player Start Playing
        Parms(0) = APIInstance.ToString
        Parms(1) = APIName
        If ChangeType = player_status_change.SongChanged Then
            ReDim Preserve Parms(6)
            Parms(2) = MyCurrentTrack
            Parms(3) = MyCurrentArtist
            Parms(4) = MyCurrentAlbum
            Parms(5) = CurrentAlbumArtPath
            Parms(6) = 0
            TriggerEvent = "Sonos Track Change"
        ElseIf ChangeType = player_status_change.PlayStatusChanged Then
            Parms(2) = ChangeValue.ToString
            If ChangeValue = player_state_values.Playing Then
                TriggerEvent = "Sonos Player Start Playing"
            ElseIf ChangeValue = player_state_values.Paused Then
                TriggerEvent = "Sonos Player Paused"
            ElseIf ChangeValue = player_state_values.Stopped Then
                TriggerEvent = "Sonos Player Stop"
            End If
        ElseIf ChangeType = player_status_change.DeviceStatusChanged Then
            Parms(2) = ChangeValue
            If ChangeValue = player_state_values.Docked Then
                TriggerEvent = "Sonos Player Docked"
            ElseIf ChangeValue = player_state_values.Undocked Then
                TriggerEvent = "Sonos Player Undocked"
            ElseIf ChangeValue = player_state_values.AudioInTrue Then
                TriggerEvent = "Sonos Player Line-in Connected"
            ElseIf ChangeValue = player_state_values.AudioInFalse Then
                TriggerEvent = "Sonos Player Line-in Disconnected"
            ElseIf ChangeValue = player_state_values.Online Then
                TriggerEvent = "Sonos Player Online"
            ElseIf ChangeValue = player_state_values.Offline Then
                TriggerEvent = "Sonos Player Offline"
            End If
        ElseIf ChangeType = player_status_change.AlarmStart Then
            TriggerEvent = "Sonos Player Alarm Start"
        ElseIf ChangeType = player_status_change.ConfigChange Then
            TriggerEvent = "Sonos Player Config Change"
        ElseIf ChangeType = player_status_change.NextSong Then
            ReDim Preserve Parms(6)
            Parms(2) = MyNextTrack
            Parms(3) = MyNextArtist
            Parms(4) = MyNextAlbum
            Parms(5) = MyNextAlbumURI
            Parms(6) = 0
            TriggerEvent = "Sonos Next Track Change"
        ElseIf ChangeType = player_status_change.AnnouncementChange Then
            Parms(2) = ChangeValue
            If ChangeValue = player_state_values.AnnouncementStart Then
                TriggerEvent = "Sonos Announcement Start"
            ElseIf ChangeValue = player_state_values.AnnouncementStop Then
                TriggerEvent = "Sonos Announcement Stop"
            End If

        End If
        Try
            'Log("RaiseGenericEventCB called with ChangeType = " & ChangeType.ToString & " and APIName = " & APIName & " and instance = " & instance, LogType.LOG_TYPE_WARNING)
            If ChangeType <= player_status_change.Library Then
                '  callback.RaiseGenericEventCB(ChangeType.ToString, Parms, APIName, instance)
            End If
        Catch ex As Exception
            Log("Error in raising generic callback event in PlayChangeNotifyCallback with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'RaiseEvent PlayChangeNotify(ChangeType, Me) 'tobefixed dcor
        Catch ex As Exception
            Log("Error in raising event in PlayChangeNotifyCallback with error : " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Parms = Nothing

        If TriggerEvent <> "" And SendDeviceTrigger Then DeviceTrigger(TriggerEvent)

    End Sub


    Public Property PlayerMute As Boolean 'Implements MediaCommon.MusicAPI.PlayerMute
        Get ' (Property Get) Boolean Returns True if the player is muted, False if it is not.
            'If g_bDebug Then Log( "Get PlayerMute called for Zone - " & ZoneName & " State = " & MyCurrentMuteState.ToString)
            PlayerMute = MyCurrentMuteState
        End Get
        Set(ByVal value As Boolean)
            If g_bDebug Then Log("Get PlayerMute called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            If value Then
                SonosMute()
            Else
                UnMute()
            End If
        End Set
    End Property

    Public Sub SonosMute()
        'Mutes the player.  Has no affect if the player is already muted.
        If g_bDebug Then Log("SonosMute called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            If ZoneModel = "WD100" And Not MyWirelessDockDestinationPlayer Is Nothing Then
                MyWirelessDockDestinationPlayer.SetMute("Master", True)
            Else
                SetMute("Master", True)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UnMute() 'Implements MediaCommon.MusicAPI.UnMute
        'UnMutes the player.  Has no affect if the player is not muted.
        If g_bDebug Then Log("UnMute called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            If ZoneModel = "WD100" And Not MyWirelessDockDestinationPlayer Is Nothing Then
                MyWirelessDockDestinationPlayer.SetMute("Master", False)
            Else
                SetMute("Master", False)
            End If
        Catch ex As Exception
        End Try
    End Sub


    Public Sub ShuffleToggle() 'Implements MediaCommon.MusicAPI.ShuffleToggle
        'Toggles through the 3 states for playlist shuffling: Shuffle, Order, Sort
        If g_bDebug Then Log("ShuffleToggle called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim CurrentShuffleState = GetPlayMode()
        Select Case CurrentShuffleState
            Case "Normal"
                PlayModeShuffle()
            Case "Shuffle No Repeat"
                PlayModeRepeatAll()
            Case "Repeat All"
                PlayModeNormal()
            Case "Shuffle"
                PlayModeShuffleNoRepeat()
            Case "Unknown"
                PlayModeNormal()
        End Select
    End Sub

    Public Sub SetShuffleState(ByVal ShuffleState As String)
        ' this is a real time update
        If MyPreviousShuffleState <> ShuffleState Then
            'PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False) ' notify HS if they have the callback linked
        End If
        If MyPreviousShuffleState <> ShuffleState And HSRefShuffle <> -1 Then
            If g_bDebug Then Log("SetShuffleState is setting HS Status for ZoneName - " & ZoneName & " with Value = " & ShuffleState.ToString & " and HSRef = " & HSRefShuffle.ToString, LogType.LOG_TYPE_INFO)
            Try
                Dim dv_Shuffle As Scheduler.Classes.DeviceClass
                dv_Shuffle = hs.GetDeviceByRef(HSRefShuffle)
                Dim dv_Repeat As Scheduler.Classes.DeviceClass
                dv_Repeat = hs.GetDeviceByRef(HSRefRepeat)
                Select Case ShuffleState.ToUpper
                    Case "NORMAL"
                        'hs.SetDeviceString(HSRefShuffle, "Ordered", True)
                        hs.SetDeviceValueByRef(HSRefShuffle, ssNoShuffle, True)
                        'dv_Shuffle.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "ordered.gif"
                        ' hs.SetDeviceString(HSRefRepeat, "RepeatOff", True)
                        hs.SetDeviceValueByRef(HSRefRepeat, rsnoRepeat, True)
                        'dv_Repeat.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "repeatoff.gif"
                    Case "SHUFFLE_NOREPEAT"
                        'hs.SetDeviceString(HSRefShuffle, "Shuffled", True)
                        hs.SetDeviceValueByRef(HSRefShuffle, ssShuffled, True)
                        'dv_Shuffle.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "shuffled.gif"
                        'hs.SetDeviceString(HSRefRepeat, "RepeatOff", True)
                        hs.SetDeviceValueByRef(HSRefRepeat, rsnoRepeat, True)
                        'dv_Repeat.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "repeatoff.gif"
                    Case "REPEAT_ALL"
                        'hs.SetDeviceString(HSRefShuffle, "Ordered", True)
                        hs.SetDeviceValueByRef(HSRefShuffle, ssNoShuffle, True)
                        'dv_Shuffle.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "ordered.gif"
                        'hs.SetDeviceString(HSRefRepeat, "RepeatAll", True)
                        hs.SetDeviceValueByRef(HSRefRepeat, rsRepeat, True)
                        'dv_Repeat.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "repeatall.gif"
                    Case "SHUFFLE"
                        'hs.SetDeviceString(HSRefShuffle, "Shuffled", True)
                        hs.SetDeviceValueByRef(HSRefShuffle, ssShuffled, True)
                        'dv_Shuffle.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "shuffled.gif"
                        'hs.SetDeviceString(HSRefRepeat, "RepeatAll", True)
                        hs.SetDeviceValueByRef(HSRefRepeat, rsRepeat, True)
                        'dv_Repeat.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "repeatall.gif"
                    Case Else
                        'hs.SetDeviceString(HSRefShuffle, "Ordered", True)
                        hs.SetDeviceValueByRef(HSRefShuffle, ssNoShuffle, True)
                        'dv_Shuffle.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "ordered.gif"
                        'hs.SetDeviceString(HSRefRepeat, "RepeatOff", True)
                        hs.SetDeviceValueByRef(HSRefRepeat, rsnoRepeat, True)
                        'dv_Repeat.Image(hs) = "http://" & hs.GetIPAddress & ImagesPath & "repeatoff.gif"
                End Select
                If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
                    MyWirelessDockDestinationPlayer.SetShuffleState(ShuffleState)
                End If
            Catch ex As Exception
                Log("Error in SetShuffleState is setting HS Status for ZoneName - " & ZoneName & " with Value = " & ShuffleState.ToString & " and HSRef = " & HSRefShuffle.ToString & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        MyShuffleState = ShuffleState
        MyPreviousShuffleState = ShuffleState ' Previous shufflestate was introduced because MyShufflestate had to be set immediately to avoid a repeat/shuffle change at the same time before Sonos had responded
    End Sub

    Public Property SonosShuffle As Integer 'Implements MediaCommon.MusicAPI.Shuffle
        Get '(Property Get) 	  	Short Integer 	Returns the current shuffle status: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            If MyWirelessDockSourcePlayer IsNot Nothing Then
                Return MyWirelessDockSourcePlayer.SonosShuffle
                Exit Property
            End If
            If ZoneIsASlave Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
                Try
                    Return LinkedZone.SonosShuffle
                Catch ex As Exception
                    SonosShuffle = Shuffle_modes.Ordered '2
                    If g_bDebug Then Log("Get SonosShuffle called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Property
            End If
            SonosShuffle = Shuffle_modes.Ordered '2
            Select Case MyShuffleState
                Case "NORMAL"
                    SonosShuffle = Shuffle_modes.Ordered ' 2
                Case "SHUFFLE_NOREPEAT"
                    SonosShuffle = Shuffle_modes.Shuffled '1
                Case "REPEAT_ALL"
                    SonosShuffle = Shuffle_modes.Ordered '2
                Case "SHUFFLE"
                    SonosShuffle = Shuffle_modes.Shuffled '1
                Case Else
                    SonosShuffle = Shuffle_modes.Ordered '2
            End Select
            'If g_bDebug Then Log( "Get Shuffle called for Zone - " & ZoneName & " with Value : " & Shuffle.ToString)
        End Get
        Set(ByVal value As Integer) '(Property Set) 	Short Integer 	  	Sets the shuffle status to the indicated value: 1 = Shuffled, 2 = Ordered, 3 = Sorted
            If g_bDebug Then Log("SonosShuffle set called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            If MyWirelessDockSourcePlayer IsNot Nothing Then
                MyWirelessDockSourcePlayer.SonosShuffle = value
                Exit Property
            End If
            If ZoneIsASlave Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
                Try
                    LinkedZone.SonosShuffle = value
                Catch ex As Exception
                    If g_bDebug Then Log("Set SonosShuffle called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Property
            End If
            Dim RepeatState As repeat_modes = SonosRepeat()
            Select Case value
                Case Shuffle_modes.Shuffled '1 ' Shuffled
                    If RepeatState = repeat_modes.repeat_all Then
                        If PlayModeShuffle() = "OK" Then MyShuffleState = "SHUFFLE" ' set the states here, prevents wrong setting when shuffle/repeat are set in one action
                    Else
                        If PlayModeShuffleNoRepeat() = "OK" Then MyShuffleState = "SHUFFLE_NOREPEAT"
                    End If
                Case Shuffle_modes.Ordered '2 ' Ordered
                    If RepeatState = repeat_modes.repeat_all Then
                        If PlayModeRepeatAll() = "OK" Then MyShuffleState = "REPEAT_ALL"
                    Else
                        If PlayModeNormal() = "OK" Then MyShuffleState = "NORMAL"
                    End If
                Case Shuffle_modes.Sorted ' 3' Sorted
                    If RepeatState = repeat_modes.repeat_all Then
                        If PlayModeRepeatAll() = "OK" Then MyShuffleState = "REPEAT_ALL"
                    Else
                        If PlayModeNormal() = "OK" Then MyShuffleState = "NORMAL"
                    End If
            End Select
            'PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
        End Set
    End Property

    Public ReadOnly Property ShuffleStatus As String 'Implements MediaCommon.MusicAPI.ShuffleStatus
        Get
            'Returns the current shuffle mode as a string value: Shuffled, Ordered, Sorted, or Unknown
            ShuffleStatus = "Unknown"
            Select Case MyShuffleState.ToUpper
                Case "NORMAL"
                    ShuffleStatus = "Ordered"
                Case "SHUFFLE_NOREPEAT"
                    ShuffleStatus = "Shuffled"
                Case "REPEAT_ALL"
                    ShuffleStatus = "Ordered"
                Case "SHUFFLE"
                    ShuffleStatus = "Shuffled"
                Case Else
                    ShuffleStatus = "Ordered"
            End Select
            'If g_bDebug Then Log( "Get ShuffleStatus called for Zone - " & ZoneName & " with Value : " & ShuffleStatus)
        End Get
    End Property


    Public ReadOnly Property RepeatStatus As String 'Implements MediaCommon.MusicAPI.RepeatStatus
        Get
            'If g_bDebug Then Log( "Get RepeatStatus called for Zone - " & ZoneName & " with Value : " & Repeat.ToString)
            RepeatStatus = SonosRepeat.ToString
        End Get
    End Property

    Public Property SonosRepeat As repeat_modes
        '(Enum) 	Returns the current repeat setting using the following Enum values:
        ' Public Enum repeat_modes
        '      repeat_off = 0
        '      repeat_one = 1
        '     repeat_all = 2
        ' End Enum
        Get
            If MyWirelessDockSourcePlayer IsNot Nothing Then
                Return MyWirelessDockSourcePlayer.SonosRepeat
                Exit Property
            End If
            If ZoneIsASlave Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
                Try
                    Return LinkedZone.SonosRepeat
                Catch ex As Exception
                    SonosRepeat = repeat_modes.repeat_off
                    If g_bDebug Then Log("Get SonosRepeat called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Property
            End If
            If MyShuffleState = "REPEAT_ALL" Or MyShuffleState = "SHUFFLE" Then
                SonosRepeat = repeat_modes.repeat_all
            Else
                SonosRepeat = repeat_modes.repeat_off
            End If
            'If g_bDebug Then Log( "Get SonosRepeat called for Zone - " & ZoneName & " with Value : " & Repeat.ToString)
        End Get
        Set(ByVal value As repeat_modes)
            If g_bDebug Then Log("Set SonosRepeat called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            If MyWirelessDockSourcePlayer IsNot Nothing Then
                MyWirelessDockSourcePlayer.SonosRepeat = value
                Exit Property
            End If
            If ZoneIsASlave Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
                Try
                    LinkedZone.SonosRepeat = value
                Catch ex As Exception
                    If g_bDebug Then Log("Set SonosRepeat called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Property
            End If
            'Dim ShuffleState As String = Shuffle
            If value = repeat_modes.repeat_all Then
                Select Case MyShuffleState
                    Case "NORMAL"
                        If PlayModeRepeatAll() = "OK" Then MyShuffleState = "REPEAT_ALL"
                    Case "SHUFFLE_NOREPEAT"
                        If PlayModeShuffle() = "OK" Then MyShuffleState = "SHUFFLE"
                End Select
            Else ' no repeat_off or repeat_one
                Select Case MyShuffleState
                    Case "REPEAT_ALL"
                        If PlayModeNormal() = "OK" Then MyShuffleState = "NORMAL"
                    Case "SHUFFLE"
                        If PlayModeShuffleNoRepeat() = "OK" Then MyShuffleState = "SHUFFLE_NOREPEAT"
                End Select
            End If
            'PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False)
        End Set
    End Property


    Public Sub SonosPlay()
        'Starts the player playing the currently loaded HomeSeer playlist.
        If g_bDebug Then Log("SonosPlay called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Play")
            Catch ex As Exception
                If g_bDebug Then Log("SonosPlay called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            SetTransportState("Play")
        Catch ex As Exception
            If g_bDebug Then Log("SonosPlay called for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SonosPause()
        'Toggles the state of the pause function of the player.
        If g_bDebug Then Log("SonosPause called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Pause")
            Catch ex As Exception
                If g_bDebug Then Log("SonosPause called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            SetTransportState("Pause")
        Catch ex As Exception
            If g_bDebug Then Log("SonosPause called for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub TogglePause()
        'Toggles the state of the pause function of the player.
        If g_bDebug Then Log("TogglePause called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.TogglePause()
            Catch ex As Exception
                If g_bDebug Then Log("Error in TogglePause for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        Try
            If PlayState = "Pause" Then
                SetTransportState("Play")
            ElseIf PlayState = "Play" Then
                SetTransportState("Pause")
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in TogglePause for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub TogglePlay()
        'Toggles the state of the play function of the player.
        If g_bDebug Then Log("TogglePlay called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.TogglePlay()
            Catch ex As Exception
                If g_bDebug Then Log("Error in TogglePlay for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        Try
            If PlayState = "Play" Then
                SetTransportState("Pause")
            Else
                SetTransportState("Play")
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in TogglePlay for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Public Sub PlayIfPaused() 'Implements MediaCommon.MusicAPI.PlayIfPaused
        'If the current state of the player is paused, the player will be resumed.
        If g_bDebug Then Log("PlayIfPaused called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        ' states are Play Stop Pause Next Previous
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Play")
            Catch ex As Exception
                If g_bDebug Then Log("PlayIfPaused called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        If PlayState = "Stop" Or PlayState = "Pause" Then
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Try
                    LinkedZone.SetTransportState("Play")
                Catch ex As Exception
                    If g_bDebug Then Log("PlayIfPaused called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Sub
            End If
            Try
                SetTransportState("Play")
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub PauseIfPlaying() 'Implements MediaCommon.MusicAPI.PauseIfPlaying
        'If the current state of the player is playing, the player will be paused.
        If g_bDebug Then Log("PauseIfPlaying called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        ' states are Play Stop Pause Next Previous
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Pause")
            Catch ex As Exception
                If g_bDebug Then Log("PauseIfPlaying called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim PlayState As String = ""
        Try
            PlayState = GetTransportState()
        Catch ex As Exception
        End Try
        If PlayState = "Play" Then
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI  'HSMusicAPI
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Try
                    LinkedZone.SetTransportState("Pause")
                Catch ex As Exception
                    If g_bDebug Then Log("PauseIfPlaying called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Exit Sub
            End If
            Try
                SetTransportState("Pause")
            Catch ex As Exception
            End Try
        End If
    End Sub

    Public Sub StopPlay() 'Implements MediaCommon.MusicAPI.StopPlay
        'Stops the player.
        If g_bDebug Then Log("StopPlay called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Stop")
            Catch ex As Exception
                If g_bDebug Then Log("StopPlay called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            SetTransportState("Stop")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub TrackNext() 'Implements MediaCommon.MusicAPI.TrackNext
        'Causes the player to jump to the next track in the playlist and begin playing it.
        If g_bDebug Then Log("TrackNext called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Next")
            Catch ex As Exception
                If g_bDebug Then Log("TrackNext called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            SetTransportState("Next")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub TrackPrev() 'Implements MediaCommon.MusicAPI.TrackPrev
        'Causes the player to start playing from the previous track in the playlist.
        If g_bDebug Then Log("TrackPrev called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState("Previous")
            Catch ex As Exception
                If g_bDebug Then Log("TrackPrev called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            SetTransportState("Previous")
        Catch ex As Exception
        End Try
    End Sub

    Public Property PlayerVolume As Integer
        Get ' Returns the current volume setting of the player from 0 to 100.
            PlayerVolume = MyCurrentMasterVolumeLevel
            'If g_bDebug Then Log( "Get Volume called for Zone - " & ZoneName & " with value = " & Volume.ToString)
        End Get
        Set(ByVal value As Integer) '(Property Set) 	Integer 	  	Sets the volume of the player to the level indicated by the parameter, in the range 0-100.
            Dim Result
            If g_bDebug Then Log("Set PlayerVolume called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Try
                If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
                    MyWirelessDockDestinationPlayer.PlayerVolume = value
                    Exit Property
                Else
                    Result = SetVolumeLevel("Master", value)
                End If
                If Result <> "OK" Then
                    If g_bDebug Then Log("PlayerVolume called for Zone - " & ZoneName & " but ended in Error: " & Result.ToString, LogType.LOG_TYPE_ERROR)
                End If
            Catch ex As Exception
                Log("PlayerVolume called for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Set
    End Property

    Public Sub VolumeUp() 'Implements MediaCommon.MusicAPI.VolumeUp
        'Increases the volume of the player by 5%
        If g_bDebug Then Log("VolumeUp called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            If ZoneModel = "WD100" And Not MyWirelessDockSourcePlayer Is Nothing Then
                MyWirelessDockSourcePlayer.ChangeVolumeLevel("Master", VolumeStep)
            Else
                ChangeVolumeLevel("Master", VolumeStep)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub VolumeDown() 'Implements MediaCommon.MusicAPI.VolumeDown
        'Decreases the volume of the player by 5%
        If g_bDebug Then Log("VolumeDown called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            If ZoneModel = "WD100" And Not MyWirelessDockSourcePlayer Is Nothing Then
                MyWirelessDockSourcePlayer.ChangeVolumeLevel("Master", -VolumeStep)
            Else
                ChangeVolumeLevel("Master", -VolumeStep)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub SetNbrOfTracks(NbrOfTracks As Integer)
        If SuperDebug Then Log("SetNbrOfTracks called for Zone - " & ZoneName & " with NbrOfTracks = " & NbrOfTracks.ToString, LogType.LOG_TYPE_INFO)
        If MyNbrOfTracksInQueue <> NbrOfTracks And ZoneSource = "Tracks" Then
            'PlayChangeNotifyCallback(player_status_change.PlayList, player_state_values.UpdateHSServerOnly, False) ' notify HS if they have the callback linked
            MyQueueHasChanged = True
            MyNbrOfTracksInQueue = NbrOfTracks
        End If
    End Sub

    Public Sub SetTrackNbr(TrackNbr As Integer)
        If SuperDebug Then Log("SetTrackNbr called for Zone - " & ZoneName & " with TrackNbr = " & TrackNbr.ToString, LogType.LOG_TYPE_INFO)
        If MyTrackInQueueNbr <> TrackNbr And ZoneSource = "Tracks" Then
            'PlayChangeNotifyCallback(player_status_change.PlayList, player_state_values.UpdateHSServerOnly, False) ' notify HS if they have the callback linked
            MyQueueHasChanged = True
            MyTrackInQueueNbr = TrackNbr
        End If
        CurrentLibEntry.Key.iKey = TrackNbr
        CurrentLibKey.iKey = TrackNbr
        CurrentLibKey.sKey = TrackNbr.ToString
    End Sub

    Public Sub SetHSPlayerInfo()
        Dim TransportInfo As String = "<table><tr><td>"
        If ArtworkURL <> "" Then
            TransportInfo &= "<img src='" & ArtworkURL & "'"    ' added ' quotes in v023
            If ArtworkVSize <> 0 Then
                TransportInfo = TransportInfo & " height=""" & ArtworkVSize.ToString & """ "
                If ArtworkHSize <> 0 Then
                    TransportInfo = TransportInfo & " width=""" & ArtworkHSize.ToString & """ "
                End If
            End If
            TransportInfo &= ">"
        End If
        If MyZoneIsPairSlave Then
            TransportInfo = TransportInfo & "</td><td><p>" & CurrentPlayerState.ToString & "</p><p>Paired to " & MyHSPIControllerRef.GetZoneNamebyUDN(MyZonePairMasterZoneUDN) & "</p><p>" & Track & "</p><p>" & Artist & "</p><p>" & Album & "</p></td></tr></table>"
        ElseIf MyZoneIsPlaybarSlave Then
            TransportInfo = TransportInfo & "</td><td><p>" & CurrentPlayerState.ToString & "</p><p>Paired to " & MyHSPIControllerRef.GetZoneNamebyUDN(MyZonePlayBarUDN) & "</p><p>" & Track & "</p><p>" & Artist & "</p><p>" & Album & "</p></td></tr></table>"
        ElseIf ZoneSource.ToUpper = "LINKED" Then
            TransportInfo = TransportInfo & "</td><td><p>" & CurrentPlayerState.ToString & "</p><p>Linked to " & MyHSPIControllerRef.GetZoneNamebyUDN(MySourceLinkedZone) & "</p><p>" & Track & "</p><p>" & Artist & "</p><p>" & Album & "</p></td></tr></table>"
        Else
            TransportInfo = TransportInfo & "</td><td><p>" & CurrentPlayerState.ToString & "</p><p>" & MyZoneSourceExt & "</p><p>" & Track & "</p><p>" & Artist & "</p><p>" & Album & "</p></td></tr></table>"
        End If
        hs.SetDeviceString(HSRefPlayer, TransportInfo, True)    ' changed from false to true in v.19 because a complain that it wasn't triggering easytrigger PI
        If SuperDebug Then Log("SetHSPlayerInfo updated HS ZonePlayer " & ZoneName & ". HS Code = " & HSRefPlayer & ". Info = " & TransportInfo, LogType.LOG_TYPE_INFO)
    End Sub

    Public WriteOnly Property VolumeLeft As Integer
        Set(ByVal value As Integer) '(Property Set) 	Integer 	  	Sets the Left volume of the player to the level indicated by the parameter, in the range 0-100.
            Dim Result
            If g_bDebug Then Log("Set Volume Left called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Try
                If ZoneModel = "WD100" And Not MyWirelessDockDestinationPlayer Is Nothing Then
                    Result = MyWirelessDockDestinationPlayer.SetVolumeLevel("LF", value)
                Else
                    Result = SetVolumeLevel("LF", value)
                End If
                If Result <> "OK" Then
                    If g_bDebug Then Log("Volume Left called for Zone - " & ZoneName & " but ended in Error: " & Result.ToString, LogType.LOG_TYPE_ERROR)
                End If
            Catch ex As Exception
                Log("Volume Left called for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Set
    End Property

    Public WriteOnly Property VolumeRight As Integer
        Set(ByVal value As Integer) '(Property Set) 	Integer 	  	Sets the Left volume of the player to the level indicated by the parameter, in the range 0-100.
            Dim Result
            If g_bDebug Then Log("Set Volume Right called for Zone - " & ZoneName & " with Value : " & value.ToString, LogType.LOG_TYPE_INFO)
            Try
                If ZoneModel = "WD100" And Not MyWirelessDockDestinationPlayer Is Nothing Then
                    Result = MyWirelessDockDestinationPlayer.SetVolumeLevel("RF", value)
                Else
                    Result = SetVolumeLevel("RF", value)
                End If
                If Result <> "OK" Then
                    If g_bDebug Then Log("Volume Right called for Zone - " & ZoneName & " but ended in Error: " & Result.ToString, LogType.LOG_TYPE_ERROR)
                End If
            Catch ex As Exception
                Log("Volume Right called for Zone - " & ZoneName & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Set
    End Property

    Public Sub PlayFavorite(Favorite As String, Optional ByVal ClearPlayerQueue As Boolean = False, Optional QueueAction As QueueActions = QueueActions.qaDontPlay)
        If g_bDebug Then Log("PlayFavorite called for Zone - " & ZoneName & " with Favorite = " & Favorite & " and ClearPlayerQueue = " & ClearPlayerQueue.ToString & " and QueueAction = " & QueueAction.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.PlayFavorite(Favorite, ClearPlayerQueue, QueueAction)
            Catch ex As Exception
                If g_bDebug Then Log("PlayFavorite called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If ZoneModel = "WD100" Then Exit Sub
        PlayMusicOnZone(CurrentAppPath & MusicDBPath, "", Me, "", "", "", "", "", "", "", "", Favorite, ClearPlayerQueue, QueueAction)

    End Sub

    Public Sub PlayMusic(ByVal Artist As String, Optional ByVal Album As String = "", Optional ByVal PlayList As String = "", Optional ByVal Genre As String = "", Optional ByVal Track As String = "", Optional ByVal StartWithArtist As String = "", Optional ByVal StartWithTrack As String = "", Optional ByVal TrackMatch As String = "", Optional ByVal AudioBook As String = "", Optional ByVal PodCast As String = "", Optional ByVal ClearPlayerQueue As Boolean = False, Optional QueueAction As QueueActions = QueueActions.qaDontPlay)
        'This causes the player to create a HomeSeer playlist matching the criteria provided and begin playing it.  
        'At least one parameter must be provided, although it may be delivered as a null string ("") if it is not desired to specify the artist.
        'Examples:
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("Phil Collins")                  Plays music by Phil Collins
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("", "", "", "Rock")             Plays music in the Rock genre
        '         hs.Plugin("Media Player").MusicAPI.PlayMusic("", "", "My Top Rated")   Plays music from the 'My Top Rated' playlist.
        If g_bDebug Then Log("PlayMusic called for Zone - " & ZoneName & " with Artist=" & Artist & " and Album=" & Album & " and Playlist=" & PlayList & " and Genre=" & Genre & " and Track=" & Track & " and StartWithArtist=" & StartWithArtist & " and StartWithTrack=" & StartWithTrack & " and TrackMatch=" & TrackMatch & "and Audiobook = " & AudioBook & " and Podcast = " & PodCast & " and ClearPlayerQueue = " & ClearPlayerQueue.ToString & " and QueueAction = " & QueueAction.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.PlayMusic(Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, AudioBook, PodCast, ClearPlayerQueue, QueueAction)
            Catch ex As Exception
                If g_bDebug Then Log("PlayMusic called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If

        ' for a WD100 zone this can be called as follows:
        ' Artist=All Albums and Album=All Tracks
        ' This is messing up queries into the DB
        If Artist = "All Albums" Then Artist = "" ' we don't have this stored in the DB

        If ZoneModel = "WD100" Then
            ' this is a complete different DB and model
            PlayMusicOnWirelessDock(Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, AudioBook, PodCast)
            Exit Sub
        End If
        If PlayList <> "" Then
            ' another special case. This can be a Playlist, a track from a playlist or a radiostation
            If Mid(PlayList, 1, 14) = "RadioStation: " Then
                PlayMusicOnZone(CurrentAppPath & MusicDBPath, "", Me, Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, "", ClearPlayerQueue, QueueAction)
            ElseIf Mid(PlayList, 1, 10) = "Pandora - " _
                Or Mid(PlayList, 1, 9) = "Sirius - " _
                Or Mid(PlayList, 1, 11) = "SiriusXM - " _
                Or Mid(PlayList, 1, 9) = "LastFM - " _
                   Or Mid(PlayList, 1, 10) = "Learned - " _
                Or Mid(PlayList, 1, 11) = "Rhapsody - " Then
                PlayRadioStation(PlayList)
            Else
                PlayMusicOnZone(CurrentAppPath & MusicDBPath, "", Me, Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, "", ClearPlayerQueue, QueueAction)
            End If
        Else
            PlayMusicOnZone(CurrentAppPath & MusicDBPath, "", Me, Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, "", ClearPlayerQueue, QueueAction)
        End If
    End Sub

    Public Sub PlayMusicOnWirelessDock(ByVal Artist As String, Optional ByVal Album As String = "", Optional ByVal PlayList As String = "", Optional ByVal Genre As String = "", Optional ByVal Track As String = "", Optional ByVal StartWithArtist As String = "", Optional ByVal StartWithTrack As String = "", Optional ByVal TrackMatch As String = "", Optional ByVal AudioBook As String = "", Optional ByVal PodCast As String = "")

        If g_bDebug Then Log("PlayMusicOnWirelessDock called for Zone - " & ZoneName & " with Artist=" & Artist & " and Album=" & Album & " and Playlist=" & PlayList & " and Genre=" & Genre & " and Track=" & Track & " and StartWithArtist=" & StartWithArtist & " and StartWithTrack=" & StartWithTrack & " and TrackMatch=" & TrackMatch & "and Audiobook = " & AudioBook & " and Podcast = " & PodCast & " and IPodPlayer = " & MyDockediPodPlayerName, LogType.LOG_TYPE_INFO)
        ' the DB to be opened depend on which ZonePlayer is docked. The iPod name can be found in MyDockediPodPlayerName
        If MyDockediPodPlayerName = "" Then Exit Sub ' we have nothing to open
        Dim DestZonePlayerName As String = ""
        If MyWirelessDockDestinationPlayer Is Nothing Then
            If DestinationZone = "" Then Exit Sub ' no default zone, we can't play anything
            DestZonePlayerName = GetZoneByUDN(DestinationZone)
            If DestZonePlayerName = "" Then
                If g_bDebug Then Log("ERROR in PlayMusicOnWirelessDock for Zone = " & ZoneName & ": cannot retrieve destination zone for Wireless dock", LogType.LOG_TYPE_ERROR)
                Exit Sub
            Else
                ' retrieve the dest Player
                Try
                    MyWirelessDockDestinationPlayer = MyHSPIControllerRef.GetMusicAPI(DestZonePlayerName)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in PlayMusicOnWirelessDock for Zone = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
                If MyWirelessDockDestinationPlayer Is Nothing Then Exit Sub
            End If
        End If

        If Not MakeiPodBrowseable() Then Exit Sub
        ' (ByVal TrackName As String, ByVal iPodPlayerName As String, ByVal ZoneUDN As String, ByVal SourceSonosPlayer As HSMusicAPI, ByVal AudioBook As Boolean)
        If AudioBook <> "" Then
            MyWirelessDockDestinationPlayer.PlayAudioBookPodCast(AudioBook, MyDockediPodPlayerName, "x-sonos-dock:" & GetUDN() & "/", Me, True)
        ElseIf PodCast <> "" Then
            MyWirelessDockDestinationPlayer.PlayAudioBookPodCast(PodCast, MyDockediPodPlayerName, "x-sonos-dock:" & GetUDN() & "/", Me, False)
        Else
            MyWirelessDockDestinationPlayer.PlayMusicOnZone(CurrentAppPath & DockedPlayersDBPath & MyDockediPodPlayerName.ToString & ".sdb", "x-sonos-dock:" & GetUDN() & "/", Me, Artist, Album, PlayList, Genre, Track, StartWithArtist, StartWithTrack, TrackMatch, "", True, QueueActions.qaPlayNow)
        End If
    End Sub

    Public Sub PlayLineInput(ByVal InputZoneName As String)
        If g_bDebug Then Log("PlayLineInput called for Zone - " & ZoneName & " with zone =" & InputZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.PlayLineInput(InputZoneName)
            Catch ex As Exception
                If g_bDebug Then Log("PlayFavorite called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim SonosPlayer As HSPI  'HSMusicAPI
        SonosPlayer = MyHSPIControllerRef.GetMusicAPI(InputZoneName)
        If Not SonosPlayer Is Nothing Then
            If SonosPlayer.ZoneModel <> "WD100" Then
                PlayURI("x-rincon-stream:" & SonosPlayer.GetUDN, "") ' set to line input
                SetTransportState("Play", True)
            End If
        End If
    End Sub

    Public Sub ResumeFromPause() 'Implements MediaCommon.MusicAPI.ResumeFromPause
        Log("ResumeFromPause called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            SetTransportState("Play")
        Catch ex As Exception
        End Try
    End Sub

    Public Function GetCurrentPlaylist() As track_desc()
        '  Array of track_desc 	
        'Returns the last HomeSeer playlist created when the player began playing whether by event action, control web page, or script command. 
        ' The return is an array of type track_desc, which is a class defined as follows:
        'Public Class track_desc
        '        Public name As String
        '        Public artist As String
        '        Public album As String
        '        Public length As String
        'End Class
        GetCurrentPlaylist = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        If g_bDebug Then Log("GetCurrentPlaylist called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim xmlData As XmlDocument = New XmlDocument
            Dim TotalMatches As Integer
            Dim ArrayIndex As Integer = 0
            Dim MetaData As String = ""
            Dim Queue() As track_desc

            Try

                Dim InArg(5)
                Dim OutArg(3)

                InArg(0) = "Q:0"
                InArg(1) = "BrowseDirectChildren"
                InArg(2) = "*"
                InArg(3) = "0"
                InArg(4) = "1"
                InArg(5) = ""

                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in GetCurrentPlaylist for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Queue = Nothing
                    Exit Function
                End Try

                TotalMatches = OutArg(2)

                If g_bDebug Then Log("GetCurrentPlaylist found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
                If TotalMatches < 1 Then
                    Queue = Nothing
                    Exit Function
                End If

                ReDim Queue(TotalMatches - 1)
                '
                For ArrayIndex = 1 To TotalMatches
                    InArg(0) = "Q:0/" & ArrayIndex.ToString
                    InArg(1) = "BrowseMetadata"
                    InArg(2) = "*"
                    InArg(3) = "0"
                    InArg(4) = "1"
                    InArg(5) = ""
                    Try
                        ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                    Catch ex As Exception
                        Log("ERROR in GetCurrentPlaylist/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit For
                    End Try
                    MetaData = OutArg(0)
                    Try
                        xmlData.LoadXml(MetaData)
                    Catch ex As Exception
                        Log("ERROR in GetCurrentPlaylist for zoneplayer = " & ZoneName & " loading XML with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit For
                    End Try
                    Dim QueueElement As New track_desc
                    Try
                        QueueElement.album = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
                    Catch ex As Exception
                        QueueElement.album = ""
                    End Try
                    Try
                        QueueElement.artist = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
                    Catch ex As Exception
                        QueueElement.artist = ""
                    End Try
                    Try
                        QueueElement.name = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        QueueElement.name = ""
                    End Try
                    Try
                        QueueElement.length = GetSeconds(xmlData.GetElementsByTagName("res").Item(0).Attributes("Duration").Value).ToString
                    Catch ex As Exception
                        QueueElement.length = "0"
                    End Try
                    Queue(ArrayIndex - 1) = QueueElement
                    'If g_bDebug Then Log( "GetCurrentPlaylist for Zone - " & ZoneName & " add Tittle " & Queue(ArrayIndex - 1).name & " add Ablum " & Queue(ArrayIndex - 1).album & " add Artist " & Queue(ArrayIndex - 1).artist & " add Length " & Queue(ArrayIndex - 1).length.ToString)
                Next
            Catch ex As Exception
                Log("Error in GetCurrentPlaylist for Zone - " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Queue = Nothing
                xmlData = Nothing
                Exit Function
            End Try
            If g_bDebug Then Log("GetCurrentPlaylist for Zone - " & ZoneName & " returned " & (UBound(Queue) + 1).ToString & " elements", LogType.LOG_TYPE_INFO)
            GetCurrentPlaylist = Queue
            xmlData = Nothing
        Catch ex As Exception
            Log("Error in GetCurrentPlaylist called for Zone - " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Public Function DBGetTracks(ByVal artist As String, ByVal album As String, ByVal genre As String, Optional ByVal Track As String = "", Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of tracks matching the parameters provided by Artist, Album, and Genre.  If empty strings are provided as parameters, then all tracks from the library are returned.  
        'Note, only the track names are returned, duplicates included.  
        If g_bDebug Then Log("DBGetTracks called for Zone - " & ZoneName & " with Artist=" & artist & " Album=" & album & " Genre=" & genre & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        Dim MyTracks As String()
        MyTracks = {""}
        DBGetTracks = MyTracks
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String
        If genre <> "" Then
            QueryString = "SELECT * FROM Tracks WHERE Genre='" & PrepareForQuery(genre) & "'"
        Else
            QueryString = "SELECT * FROM Tracks"
            If artist <> "" And album <> "" Then
                QueryString = QueryString & " WHERE Artist = '" & PrepareForQuery(artist) & "' AND Album='" & PrepareForQuery(album) & "'"
            ElseIf artist <> "" Then
                QueryString = QueryString & " WHERE Artist = '" & PrepareForQuery(artist) & "'"
            ElseIf album <> "" Then
                QueryString = QueryString & " WHERE Album='" & PrepareForQuery(album) & "'"
            End If
        End If
        If album <> "" Then QueryString = QueryString & " ORDER BY TrackNo"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("Error DBGetTracks for zoneplayer " & ZoneName & " unable to open DB with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("Error DBGetTracks for " & ZoneName & " with Query=" & QueryString, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            Index = 0
            While SQLreader.Read()
                If SQLreader("ParentID").ToString = "A:TRACKS" Then
                    tempstring = SQLreader("Name").ToString
                    'Log( "Record #" & Index.ToString & " with value " & tempstring)
                    If SQLreader("Album").ToString <> "All Tracks" Then ' this is how the WD100 does it
                        'If g_bDebug Then Log( "GetTracks for " & ZoneName & " found " & tempstring)
                        ReDim Preserve MyTracks(Index)
                        MyTracks(Index) = tempstring
                        Index = Index + 1
                    End If
                ElseIf genre <> "" And SQLreader("ParentID").ToString = "A:GENRE" Then

                    ' Go find the tracks

                    Dim Tracks As System.Array
                    Dim LoopIndex As Integer = 0

                    Try
                        Tracks = GetMatchingTracks(artist, album, genre, Track, iPodDBName)
                    Catch ex As Exception
                        Tracks = Nothing
                        Log("DBGetTracks unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                    Try
                        LoopIndex = 0
                        While LoopIndex < UBound(Tracks)
                            'Log( "GetTracks Track Info (0) = " & Tracks(LoopIndex + 0).ToString)
                            'Log( "GetTracks Track Info (1) = " & Tracks(LoopIndex + 1).ToString)
                            'Log( "GetTracks Track Info (2) = " & Tracks(LoopIndex + 2).ToString)
                            ReDim Preserve MyTracks(Index)
                            MyTracks(Index) = Tracks(LoopIndex + 0).ToString
                            Index = Index + 1
                            LoopIndex = LoopIndex + 4 ' changed from 3 to 4 in v.78
                        End While
                    Catch ex As Exception
                        Log("DBGetTracks unable read record  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                End If
            End While
            If album = "" Then
                Array.Sort(MyTracks)
            End If
            DBGetTracks = MyTracks
        Catch ex As Exception
            Log("Error DBGetTracks for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("DBGetTracks to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("DBGetTracks called for Zone - " & ZoneName & " returned " & Index.ToString & " tracks ", LogType.LOG_TYPE_INFO)
    End Function


    Public Function AllAlbums() As System.Collections.SortedList
        AllAlbums = Nothing
        ' this is something undocumented but used by the MusicPage.aspx
        If g_bDebug Then Log("AllAlbums called", LogType.LOG_TYPE_INFO)
        Dim AlbumListLists
        AlbumListLists = GetAlbums("", "")
        Dim SList As New System.Collections.SortedList
        If AlbumListLists Is Nothing Then
            If g_bDebug Then Log("AllAlbums called but found no entries", LogType.LOG_TYPE_INFO)
            SList.Add(0, "")
            AllAlbums = SList
            Exit Function
        End If
        Dim KeyIndex As Integer = 0
        While KeyIndex <= UBound(AlbumListLists, 1)
            SList.Add(KeyIndex.ToString, AlbumListLists(KeyIndex).ToString)
            KeyIndex = KeyIndex + 1
        End While
        AllAlbums = SList
        SList = Nothing
        AlbumListLists = Nothing
    End Function

    Public Function GetAlbums(ByVal artist As String, ByVal genre As String, Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of albums matching the parameters provided by Artist and Genre.  If empty strings are provided as parameters, then all tracks from the library are returned.  
        'Note, only the album names are returned, duplicates removed, in sorted order.
        If g_bDebug Then Log("GetAlbums called for Zone - " & ZoneName & " with Artist=" & artist & " Genre=" & genre & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        Dim MyAlbums As String()
        MyAlbums = {""}
        GetAlbums = MyAlbums
        Dim ConnectionString As String
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String
        If genre <> "" Then
            QueryString = "SELECT * FROM Tracks WHERE Genre='" & PrepareForQuery(genre) & "'"
        Else
            QueryString = "SELECT * FROM Tracks"
            If artist <> "" Then
                QueryString = QueryString & " WHERE Artist = '" & PrepareForQuery(artist) & "'"
            End If
        End If
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetAlbums unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                If (SQLreader("ParentID").ToString = "A:ALBUM") Or (SQLreader("ParentID").ToString = "A:TRACKS") Then
                    If (artist <> "" And artist = SQLreader("Artist").ToString) Or (artist = "") Then
                        tempstring = SQLreader("Album").ToString
                        'Log( "Record #" & Index.ToString & " with value " & tempstring)
                        If MyAlbums IsNot Nothing Then
                            Dim found As Boolean = False
                            For Each AlbumEntry In MyAlbums
                                If AlbumEntry = tempstring Then
                                    found = True
                                    Exit For
                                End If
                            Next
                            If Not found Then
                                ReDim Preserve MyAlbums(Index)
                                MyAlbums(Index) = tempstring
                                Index = Index + 1
                            End If
                        Else
                            ReDim Preserve MyAlbums(Index)
                            MyAlbums(Index) = tempstring
                            Index = Index + 1
                        End If
                    End If
                ElseIf genre <> "" And SQLreader("ParentID").ToString = "A:GENRE" Then
                    Dim Artists
                    Dim ArtistIndex As Integer = 0
                    Dim Albums
                    Try
                        Artists = GetTracks(SQLreader("Id").ToString, False, True)
                    Catch ex As Exception
                        Artists = Nothing
                        Log("Error GetAlbums for zoneplayer " & ZoneName & "unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        While ArtistIndex < UBound(Artists, 1)
                            If g_bDebug Then Log("GetAlbums Artist #" & ArtistIndex.ToString & " for Artist (0) with value " & Artists(ArtistIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                            'If g_bDebug Then Log( "GetAlbums Artist #" & ArtistIndex.ToString & " for Artist (1) with value " & Artists(ArtistIndex, 1).ToString)
                            'If g_bDebug Then Log( "GetAlbums Artist #" & ArtistIndex.ToString & " for Artist (2) with value " & Artists(ArtistIndex, 2).ToString) 'name
                            'If g_bDebug Then Log( "GetAlbums Artist #" & ArtistIndex.ToString & " for Artist (3) with value " & Artists(ArtistIndex, 3).ToString)
                            'If g_bDebug Then Log( "GetAlbums Artist #" & ArtistIndex.ToString & " for Artist (6) with value " & Artists(ArtistIndex, 6).ToString) ' ID
                            If Artists(ArtistIndex, 0) = artist Or (artist = "" And Artists(ArtistIndex, 0) <> "All" And Artists(ArtistIndex, 0) <> "All Albums") Then
                                ' now pull up its list of albums
                                Albums = GetTracks(Artists(ArtistIndex, 6), False, True) ' this is the track id so we can pull up album

                                Dim AlbumIndex As Integer = 0
                                For AlbumIndex = 0 To UBound(Albums, 1) - 1

                                    If g_bDebug Then Log("GetAlbums Album #" & AlbumIndex.ToString & " for Album (0) with value " & Albums(AlbumIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                                    'If g_bDebug Then Log( "GetAlbums Album #" & AlbumIndex.ToString & " for Album (1) with value " & Albums(AlbumIndex, 1).ToString)
                                    'If g_bDebug Then Log( "GetAlbums Album #" & AlbumIndex.ToString & " for Album (2) with value " & Albums(AlbumIndex, 2).ToString) 'name
                                    'If g_bDebug Then Log( "GetAlbums Album #" & AlbumIndex.ToString & " for Album (3) with value " & Albums(AlbumIndex, 3).ToString)
                                    'If g_bDebug Then Log( "GetAlbums Album #" & AlbumIndex.ToString & " for Album (6) with value " & Albums(AlbumIndex, 6).ToString) ' ID
                                    If Albums(AlbumIndex, 0) <> "All" And Albums(AlbumIndex, 0) <> "All Tracks" Then
                                        ReDim Preserve MyAlbums(Index)
                                        MyAlbums(Index) = Albums(AlbumIndex, 0).ToString
                                        If g_bDebug Then Log("GetAlbums for zoneplayer " & ZoneName & " Index = " & Index.ToString & " and Album found = " & Albums(AlbumIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                                        Index = Index + 1
                                    End If
                                Next
                            End If
                            ArtistIndex = ArtistIndex + 1
                        End While
                    Catch ex As Exception
                        Log("Error GetAlbums for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            End While
            Array.Sort(MyAlbums)
            GetAlbums = MyAlbums
        Catch ex As Exception
            Log("Error GetAlbums for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetAlbums to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetAlbums called for Zone - " & ZoneName & " returned " & Index.ToString & " albums ", LogType.LOG_TYPE_INFO)
    End Function


    Public Function AllArtists() As System.Collections.SortedList
        AllArtists = Nothing
        ' this is something undocumented but used by the MusicPage.aspx
        If g_bDebug Then Log("AllArtists called", LogType.LOG_TYPE_INFO)
        Dim AllArtistsLists
        AllArtistsLists = GetArtists("", "")
        Dim SList As New System.Collections.SortedList
        If AllArtistsLists Is Nothing Then
            If g_bDebug Then Log("AllArtists called but found no entries", LogType.LOG_TYPE_INFO)
            SList.Add(0, "")
            AllArtists = SList
            Exit Function
        End If
        Dim KeyIndex As Integer = 0
        While KeyIndex <= UBound(AllArtistsLists, 1)
            SList.Add(KeyIndex.ToString, AllArtistsLists(KeyIndex).ToString)
            KeyIndex = KeyIndex + 1
        End While
        AllArtists = SList
        SList = Nothing
        AllArtistsLists = Nothing
    End Function

    Public Function GetArtists(ByVal album As String, ByVal genre As String, Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of artists matching the parameters provided by Album and Genre.  If empty strings are provided as parameters, then all tracks from the library are returned.  
        'Note, only the artist names are returned, duplicates removed, in sorted order.
        If g_bDebug Then Log("GetArtists called for Zone - " & ZoneName & " Album=" & album & " Genre=" & genre & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        Dim MyLib As String()
        MyLib = {""}
        GetArtists = MyLib
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String
        If genre <> "" Then
            QueryString = "SELECT * FROM Tracks WHERE Genre='" & PrepareForQuery(genre) & "'"
        Else
            QueryString = "SELECT * FROM Tracks"
            If album <> "" Then
                QueryString = QueryString & " WHERE Album='" & PrepareForQuery(album) & "'"
            End If
        End If
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetArtists unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                If SQLreader("ParentID").ToString = "A:ALBUMARTIST" Then
                    If (album = "" And genre = "" And SQLreader("Album").ToString = "All Albums") Or (album <> "" And SQLreader("Album").ToString = album) Then
                        tempstring = SQLreader("Artist").ToString
                        'Log( "Record #" & Index.ToString & " with value " & tempstring)
                        ReDim Preserve MyLib(Index)
                        MyLib(Index) = tempstring
                        Index = Index + 1
                    End If
                ElseIf genre <> "" And SQLreader("ParentID").ToString = "A:GENRE" Then
                    Dim Artists
                    Dim ArtistIndex As Integer = 0
                    Try
                        Artists = GetTracks(SQLreader("Id").ToString, False, True)
                    Catch ex As Exception
                        Artists = Nothing
                        Log("Error GetArtists for zoneplayer " & ZoneName & " unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        While ArtistIndex < UBound(Artists, 1)
                            'If g_bDebug Then Log( "GetArtists Record #" & ArtistIndex.ToString & " for Artists (0) with value " & Artists(ArtistIndex, 0).ToString)
                            'If g_bDebug Then Log( "GetArtists Record #" & ArtistIndex.ToString & " for Artists (1) with value " & Artists(ArtistIndex, 1).ToString)
                            'If g_bDebug Then Log( "GetArtists Record #" & ArtistIndex.ToString & " for Artists (2) with value " & Artists(ArtistIndex, 2).ToString) 'name
                            'If g_bDebug Then Log( "GetArtists Record #" & ArtistIndex.ToString & " for Artists (3) with value " & Artists(ArtistIndex, 3).ToString)
                            'If g_bDebug Then Log( "GetArtists Record #" & ArtistIndex.ToString & " for Artists (6) with value " & Artists(ArtistIndex, 6).ToString) ' ID
                            If Artists(ArtistIndex, 0) <> "All" And Artists(ArtistIndex, 0) <> "All Albums" Then
                                ReDim Preserve MyLib(Index)
                                MyLib(Index) = Artists(ArtistIndex, 0).ToString
                                Index = Index + 1
                                If g_bDebug Then Log("GetArtists for zoneplayer " & ZoneName & " Artist = " & Artists(ArtistIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                            End If
                            ArtistIndex = ArtistIndex + 1
                        End While
                    Catch ex As Exception
                        Log("Error in GetArtists for zoneplayer " & ZoneName & " unable to read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            End While
            Array.Sort(MyLib)
            GetArtists = MyLib
        Catch ex As Exception
            Log("GetArtists unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetArtists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetArtists called for Zone - " & ZoneName & " returned " & Index.ToString & " artists ", LogType.LOG_TYPE_INFO)
    End Function


    Public Function AllGenres() As System.Collections.SortedList
        AllGenres = Nothing
        ' this is something undocumented but used by the MusicPage.aspx
        If g_bDebug Then Log("AllGenres called", LogType.LOG_TYPE_INFO)
        Dim AllGenresLists
        AllGenresLists = GetGenres()
        Dim SList As New System.Collections.SortedList
        If AllGenresLists Is Nothing Then
            If g_bDebug Then Log("AllGenres called but found no entries", LogType.LOG_TYPE_INFO)
            SList.Add(0, "")
            AllGenres = SList
            Exit Function
        End If
        Dim KeyIndex As Integer = 0
        While KeyIndex <= UBound(AllGenresLists, 1)
            SList.Add(KeyIndex.ToString, AllGenresLists(KeyIndex).ToString)
            KeyIndex = KeyIndex + 1
        End While
        AllGenres = SList
        SList = Nothing
        AllGenresLists = Nothing
    End Function

    Public Function GetGenres(Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of all Genres names in the system. Because HSTouch doesn't provide a way to pull up radio stations, I'm extracting them with Genres
        If g_bDebug Then Log("GetGenres called for Zone - " & ZoneName & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        GetGenres = {""}
        Dim MyGenres As String()
        MyGenres = {""}
        Dim ConnectionString As String
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            If Not MakeiPodBrowseable() Then Exit Function
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = "SELECT * FROM Tracks WHERE Name ='All Genres'"
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetGenres unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                tempstring = SQLreader("Genre").ToString
                'Log( "Record #" & Index.ToString & " with value " & tempstring)
                ReDim Preserve MyGenres(Index)
                MyGenres(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyGenres)
            GetGenres = MyGenres
        Catch ex As Exception
            Log("Error GetGenres for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetGenres to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetGenres called for Zone - " & ZoneName & " returned " & Index.ToString & " genres ", LogType.LOG_TYPE_INFO)
    End Function


    Public Function AllPlayLists(Optional ByVal IncludeRadioStations As Boolean = True) As System.Collections.SortedList
        ' this is something undocumented but used by the ActionUI
        Dim PlayLists
        PlayLists = GetPlaylists(IncludeRadioStations) 'MyControllerRef.GetAvailablePlaylists()
        If PlayLists Is Nothing Then
            If g_bDebug Then Log("AllPlayLists called but found no entries", LogType.LOG_TYPE_INFO)
            PlayLists.Add(0, "")
            AllPlayLists = PlayLists
            Exit Function
        End If
        Dim SList As New System.Collections.SortedList
        If g_bDebug Then Log("AllPlayLists called and found " & UBound(PlayLists).ToString & " entries", LogType.LOG_TYPE_INFO)
        Dim KeyIndex As Integer = 0
        Try
            While KeyIndex <= UBound(PlayLists)
                SList.Add(KeyIndex.ToString, PlayLists(KeyIndex).ToString)
                KeyIndex = KeyIndex + 1
            End While
        Catch ex As Exception
            If g_bDebug Then Log("AllPlayLists called, found " & KeyIndex.ToString & " entries but returned error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        AllPlayLists = SList
        SList = Nothing
        PlayLists = Nothing
    End Function

    Public Function GetPlaylists(Optional ByVal iPodDBName As String = "", Optional ByVal IncludeRadioStations As Boolean = True) As System.Array
        'Returns a list of all playlist names in the system.
        If g_bDebug Then Log("GetPlaylists called for Zone - " & ZoneName & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        GetPlaylists = {""}
        Dim MyPlayLists As String()
        MyPlayLists = {""}
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            If Not MakeiPodBrowseable() Then Exit Function
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String
        If IncludeRadioStations Then
            QueryString = "SELECT * FROM Tracks WHERE Name='All Playlists' or Name='All RadioStations'"
        Else
            QueryString = "SELECT * FROM Tracks WHERE Name='All Playlists'"
        End If
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetPlaylists unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                tempstring = SQLreader("Artist").ToString
                'Log( "Record #" & Index.ToString & " with value " & tempstring)
                If tempstring <> "All Tracks" Then
                    ReDim Preserve MyPlayLists(Index)
                    MyPlayLists(Index) = tempstring
                    Index = Index + 1
                End If
            End While
            Array.Sort(MyPlayLists)
            GetPlaylists = MyPlayLists
        Catch ex As Exception
            Log("Error GetPlaylists for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetPlaylists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If ZoneModel = "WD100" Then Exit Function ' we're done here

        If Not IncludeRadioStations Then Exit Function

        ' Now look for "learned" Radio Stations like Pandora

        ConnectionString = "Data Source=" & CurrentAppPath & RadioStationsDBPath
        QueryString = "SELECT * FROM RadioStations"
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            If g_bDebug Then Log("Error GetPlaylists for zoneplayer " & ZoneName & " unable to open RadioStation DB with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            While SQLreader.Read()
                tempstring = SQLreader("Name").ToString
                'Log( "Record #" & Index.ToString & " with value " & tempstring)
                ReDim Preserve MyPlayLists(Index)
                MyPlayLists(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyPlayLists)
            GetPlaylists = MyPlayLists
        Catch ex As Exception
            Log("Error GetPlaylists for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetPlaylists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetPlaylists called for Zone - " & ZoneName & " returned " & Index.ToString & " playlists ", LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetCurrentPlaylistTracks() As String()
        ' this is something undocumented but used by the ActionUI
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                Return LinkedZone.GetCurrentPlaylistTracks()
            Catch ex As Exception
                If g_bDebug Then Log("GetCurrentPlaylistTracks called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                Return {""}
            End Try
            Exit Function
        End If
        Return GetQueue()


        Dim QueueInformation As String() = Nothing
        QueueInformation = GetQueue()
        If QueueInformation Is Nothing Then
            If g_bDebug Then Log("GetCurrentPlaylistTracks called but found no entries", LogType.LOG_TYPE_INFO)
            GetCurrentPlaylistTracks = {""}
            Exit Function
        ElseIf QueueInformation.Count = 0 Then
            If g_bDebug Then Log("GetCurrentPlaylistTracks called but found no entries", LogType.LOG_TYPE_INFO)
            GetCurrentPlaylistTracks = {""}
            Exit Function
        End If
        Dim SList() As String = {""}
        If g_bDebug Then Log("GetCurrentPlaylistTracks called and found " & (UBound(QueueInformation, 1) + 1).ToString & " entries", LogType.LOG_TYPE_INFO)
        Dim KeyIndex As Integer = 0
        ReDim Preserve SList(QueueInformation.Count - 1)
        Try
            While KeyIndex < QueueInformation.Count
                SList(KeyIndex) = QueueInformation(KeyIndex)
                KeyIndex = KeyIndex + 1
            End While
        Catch ex As Exception
            If g_bDebug Then Log("GetCurrentPlaylistTracks called, found " & KeyIndex.ToString & " entries but returned error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        GetCurrentPlaylistTracks = SList
        SList = Nothing
        QueueInformation = Nothing
    End Function

    Public Sub ClearCurrentPlayList()
        ClearQueue()
    End Sub

    Public Sub AddTrackToCurrentPlaylist(ByVal Artist As String, ByVal Album As String, ByVal Track As String)
        PlayMusic(Artist, Album, "", "", Track, "", "", "", "", "", False, QueueActions.qaDontPlay)
    End Sub

    Public Function GetPlaylistTracks(ByVal playlist_name As String, Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of all tracks in the given playlist name.  If an empty string is provided as the parameter, then an empty array is returned, otherwise all tracks in the given playlist are returned with duplicates.
        Dim MyPlayListTracks As String()
        If g_bDebug Then Log("GetPlaylistTracks called for Zone - " & ZoneName & " with Playlist=" & playlist_name & " and iPodDBName = " & iPodDBName, LogType.LOG_TYPE_INFO)
        MyPlayListTracks = {""}
        GetPlaylistTracks = MyPlayListTracks
        Dim ConnectionString As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            If Not MakeiPodBrowseable() Then Exit Function
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodDBName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String = "SELECT * FROM Tracks WHERE Name = 'All Playlists' OR Name = 'All RadioStations'"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetPlaylistTracks unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                If SQLreader("Artist").ToString = playlist_name Then
                    If SQLreader("Name").ToString = "All RadioStations" Then
                        ' this is a radiostation, just return its name
                        ReDim Preserve MyPlayListTracks(Index)
                        MyPlayListTracks(Index) = playlist_name
                        Index = Index + 1
                    Else
                        Dim Tracks
                        Dim TrackIndex As Integer = 0
                        Try
                            Tracks = GetTracks(SQLreader("Id").ToString, True, True) ' changed in v.80 because tracks wouldn't show up 
                        Catch ex As Exception
                            Tracks = Nothing
                            Log("Error GetPlaylistTracks for zoneplayer " & ZoneName & " unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            While TrackIndex < UBound(Tracks, 1)
                                'If g_bDebug then Log( "Record #" & Index.ToString & " with value " & Tracks(TrackIndex, 0).ToString)
                                ReDim Preserve MyPlayListTracks(Index)
                                MyPlayListTracks(Index) = Tracks(TrackIndex, 0).ToString
                                Index = Index + 1
                                TrackIndex = TrackIndex + 1
                            End While
                        Catch ex As Exception
                            ' probably no tracks stored
                            ' Log( "GetPlaylistTracks unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                End If
            End While
            GetPlaylistTracks = MyPlayListTracks
        Catch ex As Exception
            Log("Error GetPlaylistTracks for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetPlaylistTracks to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetPlaylistTracks called for Zone - " & ZoneName & " returned " & Index.ToString & " tracks ", LogType.LOG_TYPE_INFO)
    End Function

    Public Function AllRadioStationLists(Optional ByRef IncludeLearned As Boolean = True) As System.Collections.SortedList
        ' this is something undocumented but used by the MusicPage.aspx
        If g_bDebug Then Log("AllRadioStationLists called", LogType.LOG_TYPE_INFO)
        Dim RadioStationLists
        RadioStationLists = LibGetRadioStationlists(IncludeLearned) 'MyControllerRef.GetAvailableRadioStations()
        Dim SList As New System.Collections.SortedList
        If RadioStationLists Is Nothing Then
            If g_bDebug Then Log("AllRadioStationLists called but found no entries", LogType.LOG_TYPE_INFO)
            SList.Add(0, "")
            AllRadioStationLists = SList
            Exit Function
        End If
        Dim KeyIndex As Integer = 0
        While KeyIndex <= UBound(RadioStationLists, 1)
            SList.Add(KeyIndex.ToString, RadioStationLists(KeyIndex, 0).ToString)
            KeyIndex = KeyIndex + 1
        End While
        AllRadioStationLists = SList
        SList = Nothing
        RadioStationLists = Nothing
    End Function

    Public Function LibGetRadioStationlists(Optional ByRef IncludeLearned As Boolean = True) As System.Array
        'Returns a list of all RadioStations in the system.
        If g_bDebug Then Log("LibGetRadioStationlists called for Zone - " & ZoneName & " and IncludeLearned = " & IncludeLearned.ToString, LogType.LOG_TYPE_INFO)
        LibGetRadioStationlists = {""}
        Dim MyRadioStationLists As String()
        MyRadioStationLists = {""}
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" Then Exit Function
            If Not MakeiPodBrowseable() Then Exit Function
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & MyDockediPodPlayerName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String = "SELECT * FROM Tracks WHERE Name='All RadioStations'"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("LibGetRadioStationlists unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                tempstring = SQLreader("Artist").ToString
                'Log( "Record #" & Index.ToString & " with value " & tempstring)
                ReDim Preserve MyRadioStationLists(Index)
                MyRadioStationLists(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyRadioStationLists)
            LibGetRadioStationlists = MyRadioStationLists
        Catch ex As Exception
            Log("Error in LibGetRadioStationlists for zone " & ZoneName & "  unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("LibGetRadioStationlists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If ZoneModel = "WD100" Then Exit Function ' we're done here

        If Not IncludeLearned Then Exit Function

        ' Now look for "learned" Radio Stations like Pandora

        ConnectionString = "Data Source=" & CurrentAppPath & RadioStationsDBPath
        QueryString = "SELECT * FROM RadioStations"
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            If g_bDebug Then Log("Error in LibGetRadioStationlists for zone " & ZoneName & "  unable to open RadioStation DB with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            While SQLreader.Read()
                tempstring = SQLreader("Name").ToString
                ReDim Preserve MyRadioStationLists(Index)
                MyRadioStationLists(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyRadioStationLists)
            LibGetRadioStationlists = MyRadioStationLists
        Catch ex As Exception
            Log("Error in LibGetRadioStationlists for zone " & ZoneName & "  unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("LibGetRadioStationlists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("LibGetRadioStationlists called for Zone - " & ZoneName & " returned " & Index.ToString & " radiostations ", LogType.LOG_TYPE_INFO)
    End Function

    Public Function LibGetAudiobookslists(ByVal iPodPlayerName As String) As System.Array
        'Returns a list of all Audiobooks for this player
        If g_bDebug Then Log("LibGetAudiobookslists called for Zone - " & ZoneName & " with iPodName = " & iPodPlayerName, LogType.LOG_TYPE_INFO)
        LibGetAudiobookslists = {""}
        If iPodPlayerName = "" Then Exit Function
        'If Not MakeiPodBrowseable() Then Exit Function
        Dim MyAudiobookLists As String()
        MyAudiobookLists = {""}
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0

        ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodPlayerName & ".sdb"
        Dim QueryString As String = "SELECT * FROM Tracks WHERE `ParentID`='A:AUDIOBOOK'"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetPlaylistTracks unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                tempstring = SQLreader("Name").ToString
                If g_bDebug Then Log("Record #" & Index.ToString & " with value " & tempstring, LogType.LOG_TYPE_INFO)
                ReDim Preserve MyAudiobookLists(Index)
                MyAudiobookLists(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyAudiobookLists)
            LibGetAudiobookslists = MyAudiobookLists
        Catch ex As Exception
            Log("Error in LibGetAudiobookslists for zone " & ZoneName & "  unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("LibGetAudiobookslists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("LibGetAudiobookslists called for Zone - " & ZoneName & " returned " & Index.ToString & " audiobooks ", LogType.LOG_TYPE_INFO)
    End Function

    Public Function LibGetPodcastlists(ByVal iPodPlayerName As String) As System.Array
        'Returns a list of all Podcasts for this player
        If g_bDebug Then Log("LibGetPodcastlists called for Zone - " & ZoneName & " with iPodName = " & iPodPlayerName, LogType.LOG_TYPE_INFO)
        LibGetPodcastlists = {""}
        If iPodPlayerName = "" Then Exit Function
        'If Not MakeiPodBrowseable() Then Exit Function
        Dim MyPodcastLists As String()
        MyPodcastLists = {""}
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0

        ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodPlayerName & ".sdb"
        Dim QueryString As String = "SELECT * FROM Tracks WHERE `ParentID`='A:PODCAST'"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("LibGetPodcastlists unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                tempstring = SQLreader("Name").ToString
                If g_bDebug Then Log("Record #" & Index.ToString & " with value " & tempstring, LogType.LOG_TYPE_INFO)
                ReDim Preserve MyPodcastLists(Index)
                MyPodcastLists(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyPodcastLists)
            LibGetPodcastlists = MyPodcastLists
        Catch ex As Exception
            Log("Error in LibGetPodcastlists for zone " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("LibGetPodcastlists to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("LibGetPodcastlists called for Zone - " & ZoneName & " returned " & Index.ToString & " podcasts ", LogType.LOG_TYPE_INFO)
    End Function

    Public Function LibGetObjectslist(ObjectID As String) As System.Array
        'Returns a list of all Podcasts for this player
        If g_bDebug Then Log("LibGetObjectslist called for Zone - " & ZoneName & " with ObjectID = " & ObjectID, LogType.LOG_TYPE_INFO)
        LibGetObjectslist = {""}
        'If Not MakeiPodBrowseable() Then Exit Function
        Dim MyObjectList As String()
        MyObjectList = {""}
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0

        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" Then Exit Function
            If Not MakeiPodBrowseable() Then Exit Function
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & MyDockediPodPlayerName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String = "SELECT * FROM Tracks WHERE `ParentID`='" & ObjectID & "'"
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("LibGetObjectslist unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            Index = 0
            While SQLreader.Read()
                'tempstring = SQLreader("Name").ToString
                tempstring = SQLreader("Artist").ToString
                If g_bDebug Then Log("Record #" & Index.ToString & " with value " & tempstring, LogType.LOG_TYPE_INFO)
                ReDim Preserve MyObjectList(Index)
                MyObjectList(Index) = tempstring
                Index = Index + 1
            End While
            Array.Sort(MyObjectList)
            LibGetObjectslist = MyObjectList
        Catch ex As Exception
            Log("Error in LibGetObjectslist for zone " & ZoneName & " with ObjectID = " & ObjectID & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("LibGetObjectslist to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("LibGetObjectslist called for Zone - " & ZoneName & " with ObjectID = " & ObjectID & " returned " & Index.ToString & " Objects ", LogType.LOG_TYPE_INFO)
    End Function


    Public Function LibGetiPodNameLists() As System.Array
        'Returns a list of all iPods that have a Music DB
        LibGetiPodNameLists = {""}
        Dim MyiPodLists() As String = {""}
        Dim DBiPodNamesString As String = ""
        Dim iPodDBNames
        DBiPodNamesString = GetStringIniFile("iPod Player Names", "iPod Players", "")
        iPodDBNames = Split(DBiPodNamesString, ":|:")
        Dim Index As Integer = 0
        For Each iPodDBName In iPodDBNames
            If (File.Exists(CurrentAppPath & DockedPlayersDBPath & iPodDBName & ".sdb")) Then
                ' Only players with a DB will be returned
                ReDim Preserve MyiPodLists(Index)
                MyiPodLists(Index) = iPodDBName
                Index = Index + 1
            End If
        Next
        If g_bDebug Then Log("LibGetiPodNameLists called for Zone - " & ZoneName & " and found " & Index.ToString & " iPod Player Names", LogType.LOG_TYPE_INFO)
        LibGetiPodNameLists = MyiPodLists
    End Function

    Public Function LibGetActiveLineInputlists() As System.Array
        LibGetActiveLineInputlists = {""}
        Dim MyLineInputLists() As String = {""}
        Dim ZoneNameString As String = ""
        ZoneNameString = GetStringIniFile("Sonos Zonenames", "Names", "")
        Dim ZoneNames As Object
        Dim ZoneInfos As Object
        Dim Index As Integer = 0
        ZoneNames = Split(ZoneNameString, ":|:")
        Try
            For Each ZoneInfos In ZoneNames
                Dim ZoneNameInfos As Object
                ZoneNameInfos = Split(ZoneInfos, ";:;")
                If ZoneNameInfos(1) <> "WD100" Then
                    Dim SonosPlayer As HSPI  'HSMusicAPIHSMusicAPI
                    SonosPlayer = MyHSPIControllerRef.GetMusicAPI(ZoneNameInfos(0))
                    If SonosPlayer.LineInputConnected Then
                        ReDim Preserve MyLineInputLists(Index)
                        MyLineInputLists(Index) = ZoneNameInfos(0)
                        Index = Index + 1
                    End If
                End If
            Next
            If g_bDebug Then Log("LibGetActiveLineInputlists called for Zone - " & ZoneName & " and found " & Index.ToString & " Player Names", LogType.LOG_TYPE_INFO)
            LibGetActiveLineInputlists = MyLineInputLists
        Catch ex As Exception
            Log("Error in LibGetActiveLineInputlists for Zone - " & ZoneName & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function LibGetStereoPlayerlist(PlayerType As String) As System.Array
        LibGetStereoPlayerlist = {""}
        Dim MyPlayerList() As String = {""}
        Dim ZoneNameString As String = ""
        ZoneNameString = GetStringIniFile("Sonos Zonenames", "Names", "")
        Dim ZoneNames As Object
        Dim ZoneInfos As Object
        Dim Index As Integer = 0
        ZoneNames = Split(ZoneNameString, ":|:")
        Try
            For Each ZoneInfos In ZoneNames
                Dim ZoneNameInfos As Object
                ZoneNameInfos = Split(ZoneInfos, ";:;")
                If ZoneNameInfos(1) = PlayerType And ZoneNameInfos(0) <> ZonePlayerName Then
                    '  If ZoneNameInfos(1) = "S5" And ZoneNameInfos(0) <> ZonePlayerName Then
                    ReDim Preserve MyPlayerList(Index)
                    MyPlayerList(Index) = ZoneNameInfos(0)
                    Index = Index + 1
                End If
            Next
            If g_bDebug Then Log("LibGetStereoPlayerlist called for Zone - " & ZoneName & " and found " & Index.ToString & " " & PlayerType & " Player Names", LogType.LOG_TYPE_INFO)
            LibGetStereoPlayerlist = MyPlayerList
        Catch ex As Exception
            Log("Error in LibGetStereoPlayerlist for Zone - " & ZoneName & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Public Function LibGetLineInputlists() As System.Array
        LibGetLineInputlists = {""}
        Dim MyLineInputLists() As String = {""}
        Dim ZoneNameString As String = ""
        ZoneNameString = GetStringIniFile("Sonos Zonenames", "Names", "")
        Dim ZoneNames As Object
        Dim ZoneInfos As Object
        Dim Index As Integer = 0
        ZoneNames = Split(ZoneNameString, ":|:")
        Try
            For Each ZoneInfos In ZoneNames
                Dim ZoneNameInfos As Object
                ZoneNameInfos = Split(ZoneInfos, ";:;")
                If ZoneNameInfos(1) <> "WD100" Then
                    ReDim Preserve MyLineInputLists(Index)
                    MyLineInputLists(Index) = ZoneNameInfos(0)
                    Index = Index + 1
                End If
            Next
            If g_bDebug Then Log("LibGetLineInputlists called for Zone - " & ZoneName & " and found " & Index.ToString & " Player Names", LogType.LOG_TYPE_INFO)
            LibGetLineInputlists = MyLineInputLists
        Catch ex As Exception
            Log("Error in LibGetLineInputlists for Zone - " & ZoneName & " with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub SonosSkipToTrack(ByVal track_num As Integer)
        '(Overloaded) 	Integer 	  	Jumps the player to the track number in the current HomeSeer playlist given in the Integer parameter.  Track numbers less than 0 or greater than the number of entries in the playlist are ignored.
        If g_bDebug Then Log("SonosSkipToTrack(integer) called for Zone - " & ZoneName & " with value = " & track_num.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SonosSkipToTrack(track_num)
            Catch ex As Exception
                If g_bDebug Then Log("SonosSkipToTrack called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Try
            PlayFromQueue("Q:")
            SeekTrack(track_num + 1) ' dcor not sure this is needed ??
            PlayIfPaused()
        Catch ex As Exception
            Log("Error in SonosSkipToTrack with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SonosSkipToTrack(ByVal track_name As String)
        '(Overloaded) 	String 	  	Jumps the player to track matching the track name provided as the parameter.  If the track name does not match any entries in the current playlist, nothing happens.
        If g_bDebug Then Log("SonosSkipToTrack(String) called for Zone - " & ZoneName & " with value = " & track_name.ToString, LogType.LOG_TYPE_INFO)
        Dim QueueInformation() As String = Nothing
        QueueInformation = GetCurrentPlaylistTracks()
        If QueueInformation Is Nothing Then Exit Sub
        Dim Index As Integer
        For Index = 0 To UBound(QueueInformation)
            If QueueInformation(Index) = track_name Then
                ' we found it, use the index to start the track
                PlayFromQueue("Q:")
                SeekTrack(Index + 1) ' Sonos starts at 1 not 0 as an Index
                PlayIfPaused()
                Exit Sub
            End If
        Next
        Log("SonosSkipToTrack(String) called for Zone - " & ZoneName & " but couldn't find = " & track_name.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub StartPlay() 'Implements MediaCommon.MusicAPI.StartPlay
        'Like Play, this command starts the player playing the currently loaded HomeSeer playlist, but StartPlay always starts at playlist entry 0 (the beginning).
        If g_bDebug Then Log("StartPlay called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            SetTransportState("Play")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub LoadPlayLists()
        'HomeSeer automatically detects changes to the media playlists (tracks added/removed) and regenerates its internal list of playlist tracks.  
        'This procedure may be used if desired to regenerate the list of playlist tracks manually.
        'Note: Special playlists such as "Recently Played" are updated automatically as well, but nothing will be logged in HomeSeer to indicate this 
        'since these playlists are updated frequently by the player.
        If g_bDebug Then Log("LoadPlayLists called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub Link(MasterZoneUDN As String, Optional AddGroup As Boolean = False)
        If g_bDebug Then Log("Link called for Zone - " & ZoneName & " and MasterZoneUDN = " & MasterZoneUDN, LogType.LOG_TYPE_INFO)
        If Mid(MasterZoneUDN, 1, 5) = "uuid:" Then
            Mid(MasterZoneUDN, 1, 5) = "     " ' remove the uuid:
        End If
        MasterZoneUDN = Trim(MasterZoneUDN)
        ' Go find the zone where this belongs to 
        If AddGroup Then
            Dim SonosPlayer As HSPI = Nothing
            Dim GroupSourceUDN As String
            SonosPlayer = MyHSPIControllerRef.GetAPIByUDN(MasterZoneUDN)
            If SonosPlayer IsNot Nothing Then
                GroupSourceUDN = SonosPlayer.GetZoneSourceUDN()
                If GroupSourceUDN <> "" Then
                    ' player is linked reroute to GroupSourceUDN
                    MasterZoneUDN = GroupSourceUDN
                End If
                SonosPlayer = Nothing
            Else
                If g_bDebug Then Log("Error in Link with MasterZoneUDN : " & MasterZoneUDN & "  but didn't find MasterPlayer", LogType.LOG_TYPE_ERROR)
            End If
        End If
        If ZoneModel <> "WD100" Then
            PlayURI("x-rincon:" & MasterZoneUDN, "") ' group
        Else
            PlayURI("x-sonos-dock:" & MasterZoneUDN, "") ' group
        End If
    End Sub

    Public Sub Unlink()
        If g_bDebug Then Log("Unlink called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        ' Unlink by pointing the player to its queue
        If ZoneIsLinked Then
            PlayURI("x-rincon-queue:" & GetUDN() & "#0", "")
        ElseIf MyZoneIsSourceForLinkedZone Then
            BecomeCoordinatorOfStandaloneGroup()
        End If
    End Sub

    Public Sub RadioStationPrev()
        If g_bDebug Then Log("RadioStationPrev called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim LastRadioStationIndex As Integer = 0
        LastRadioStationIndex = GetIntegerIniFile("ListIndexes", "Radiostations", 0)
        Dim IncludeLearned As Boolean = True
        IncludeLearned = GetBooleanIniFile("ListIndexes", "Include Learned Radiostations", False)
        Dim AllRadioStations As String()
        AllRadioStations = LibGetRadioStationlists(IncludeLearned)
        If g_bDebug Then Log("RadioStationPrev called for Zone - " & ZoneName & " found " & UBound(AllRadioStations, 1) & " Radiostations", LogType.LOG_TYPE_INFO)
        If (LastRadioStationIndex <= 0) Or (LastRadioStationIndex > UBound(AllRadioStations, 1)) Then
            LastRadioStationIndex = UBound(AllRadioStations, 1)
        Else
            LastRadioStationIndex = LastRadioStationIndex - 1
        End If
        WriteIntegerIniFile("ListIndexes", "Radiostations", LastRadioStationIndex)
        If g_bDebug Then Log("RadioStationPrev called for Zone - " & ZoneName & " is going to play RadioStation " & (AllRadioStations(LastRadioStationIndex)) & " with Index = " & LastRadioStationIndex, LogType.LOG_TYPE_INFO)
        PlayRadioStation(AllRadioStations(LastRadioStationIndex))
    End Sub

    Public Sub RadioStationNext()
        If g_bDebug Then Log("RadioStationNext called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim LastRadioStationIndex As Integer = 0
        LastRadioStationIndex = GetIntegerIniFile("ListIndexes", "Radiostations", 0)
        Dim IncludeLearned As Boolean = True
        IncludeLearned = GetBooleanIniFile("ListIndexes", "Include Learned Radiostations", False)
        Dim AllRadioStations As String()
        AllRadioStations = LibGetRadioStationlists(IncludeLearned)
        If g_bDebug Then Log("RadioStationNext called for Zone - " & ZoneName & " found " & UBound(AllRadioStations, 1) & " Radiostations", LogType.LOG_TYPE_INFO)
        If UBound(AllRadioStations, 1) > LastRadioStationIndex Then
            LastRadioStationIndex = LastRadioStationIndex + 1
        Else
            LastRadioStationIndex = 0
        End If
        WriteIntegerIniFile("ListIndexes", "Radiostations", LastRadioStationIndex)
        If g_bDebug Then Log("RadioStationNext called for Zone - " & ZoneName & " is going to play RadioStation " & (AllRadioStations(LastRadioStationIndex)) & " with Index = " & LastRadioStationIndex, LogType.LOG_TYPE_INFO)
        PlayRadioStation(AllRadioStations(LastRadioStationIndex))
    End Sub

    Public Sub PlaylistPrev()
        If g_bDebug Then Log("PlaylistPrev called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim LastPlayListIndex As Integer = 0
        LastPlayListIndex = GetIntegerIniFile("ListIndexes", "Playlists", 0)
        'Dim IncludeLearned As Boolean = False
        'IncludeLearned = objIniFile.GetBoolean("ListIndexes", "Include Learned Radiostations", False)
        Dim AllPlaylists As String()
        AllPlaylists = GetPlaylists("", False)
        If g_bDebug Then Log("PlaylistPrev called for Zone - " & ZoneName & " found " & UBound(AllPlaylists, 1) & " Playlists", LogType.LOG_TYPE_INFO)
        'Dim TempIndex As Integer
        'For TempIndex = 0 To UBound(AllPlaylists)
        'If g_bDebug Then Log( "      Playlist = " & AllPlaylists(TempIndex))
        'Next
        If (LastPlayListIndex <= 0) Or (LastPlayListIndex > UBound(AllPlaylists, 1)) Then
            LastPlayListIndex = UBound(AllPlaylists, 1)
        Else
            LastPlayListIndex = LastPlayListIndex - 1
        End If
        WriteIntegerIniFile("ListIndexes", "Playlists", LastPlayListIndex)
        If g_bDebug Then Log("PlaylistPrev called for Zone - " & ZoneName & " is going to play Playlist " & (AllPlaylists(LastPlayListIndex)) & " with Index = " & LastPlayListIndex, LogType.LOG_TYPE_INFO)
        PlayMusic("", "", AllPlaylists(LastPlayListIndex), "", "", "", "", "", "", "", True, QueueActions.qaPlayNow)
    End Sub

    Public Sub PlaylistNext()
        If g_bDebug Then Log("PlaylistNext called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim LastPlayListIndex As Integer = 0
        LastPlayListIndex = GetIntegerIniFile("ListIndexes", "Playlists", 0)
        'Dim IncludeLearned As Boolean = False
        'IncludeLearned = objIniFile.GetBoolean("ListIndexes", "Include Learned Radiostations", False)
        Dim AllPlaylists As String()
        AllPlaylists = GetPlaylists("", False)
        If g_bDebug Then Log("PlaylistNext called for Zone - " & ZoneName & " found " & UBound(AllPlaylists, 1) & " Playlists", LogType.LOG_TYPE_INFO)
        'Dim TempIndex As Integer
        'For TempIndex = 0 To UBound(AllPlaylists)
        'If g_bDebug Then Log( "      Playlist = " & AllPlaylists(TempIndex))
        'Next
        If UBound(AllPlaylists, 1) > LastPlayListIndex Then
            LastPlayListIndex = LastPlayListIndex + 1
        Else
            LastPlayListIndex = 0
        End If
        WriteIntegerIniFile("ListIndexes", "Playlists", LastPlayListIndex)
        If g_bDebug Then Log("PlaylistNext called for Zone - " & ZoneName & " is going to play Playlist " & (AllPlaylists(LastPlayListIndex)) & " with Index = " & LastPlayListIndex, LogType.LOG_TYPE_INFO)
        PlayMusic("", "", AllPlaylists(LastPlayListIndex), "", "", "", "", "", "", "", True, QueueActions.qaPlayNow)
    End Sub

    Public Function PlayListCreateNew(ByVal playlist_name As String) As String
        'Creates a new playlist with the name of the specified string. Returns a null string if successful or an error if unsuccessful.
        PlayListCreateNew = ""
        If g_bDebug Then Log("PlayListCreateNew called for Zone - " & ZoneName & " and Playlist = " & playlist_name, LogType.LOG_TYPE_INFO)
    End Function

    Public Function PlayListDeletePlaylist(ByVal playlist_name As String) As String
        'Deletes a playlist with the name of the specified string. Returns a null string if successful or an error if unsuccessful.
        PlayListDeletePlaylist = ""
        If g_bDebug Then Log("PlayListDeletePlayList called for Zone - " & ZoneName & " with Name=" & playlist_name, LogType.LOG_TYPE_INFO)
    End Function

    Public Function PlayListDeleteTrack(ByVal playlist_name As String, ByVal track_name As String, Optional ByVal Album As String = "") As String
        'Deletes a track from the specified playlist that has the specified name. The Album may be used to narrow the search. Returns a null string if successful or an error if unsuccessful.
        PlayListDeleteTrack = ""
        If g_bDebug Then Log("PlayListDeleteTrack called for Zone - " & ZoneName & " for Playlist=" & playlist_name & " and Track= " & track_name & " and Album= " & Album, LogType.LOG_TYPE_INFO)
    End Function

    Public Function PlayListAddTrack(ByVal playlist_name As String, ByVal track_name As String, Optional ByVal Album As String = "") As String
        'Adds a track into the specified playlist that has the specified name. The Album may be used to narrow the search (this is useful if you have multiple tracks that share the same name). Returns a null string if successful or an error if unsuccessful.
        PlayListAddTrack = "OK"
        If g_bDebug Then Log("PlayListAddTrack called for Zone - " & ZoneName & " and Playlist = " & playlist_name & " and Track = " & track_name & " and Album = " & Album, LogType.LOG_TYPE_INFO)
    End Function


    Public ReadOnly Property APIInstance As Integer
        Get
            'Log( "Get APIInstance called for Zone - " & ZoneName & ". Index = " & MyHSTMusicIndex.Message, LogType.LOG_TYPE_ERROR)
            APIInstance = MyHSTMusicIndex
        End Get
    End Property

    Public ReadOnly Property APIName As String
        Get
            'Log( " GET APIInstance called for Zone - " & ZoneName)
            APIName = ZoneName
        End Get
    End Property


    Public Sub CheckLibChange()
        If g_bDebug And gIOEnabled Then Log("CheckLibChange called for ZoneName = " & ZoneName, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub CheckPlaylistChange()
        If g_bDebug And gIOEnabled Then Log("CheckPlaylistChange called for ZoneName = " & ZoneName, LogType.LOG_TYPE_INFO)
    End Sub

    Public Property LastPlayInfo() As LastMusic
        Get
            If g_bDebug And gIOEnabled Then Log("LastPlayInfo Get called for ZoneName = " & ZoneName, LogType.LOG_TYPE_INFO)
            LastPlayInfo = Nothing
        End Get
        Set(ByVal value As LastMusic)
            If g_bDebug And gIOEnabled Then Log("LastPlayInfo Set called for ZoneName = " & ZoneName, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    Public ReadOnly Property LinkedZoneSource As String
        Get
            LinkedZoneSource = MySourceLinkedZone
        End Get
    End Property

    Public ReadOnly Property LinkedZoneDestination As String()
        Get
            LinkedZoneDestination = GetZoneDestination()
        End Get
    End Property

    Dim returnEntryKey As New HomeSeerAPI.Lib_Entry_Key

    Public Function CurrentlyPlaying() As HomeSeerAPI.Lib_Entry_Key Implements HomeSeerAPI.IMediaAPI.CurrentlyPlaying
        If SuperDebug Or DCORMEDIAAPITrace Then Log("CurrentlyPlaying called for Zone - " & ZoneName & " returning Key = " & TranslateLibEntryKey(CurrentLibKey), LogType.LOG_TYPE_INFO, LogColorNavy)
        Return CurrentLibKey
    End Function

    Public Function CurrentPlayList() As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.CurrentPlayList
        Dim ReturnList As HomeSeerAPI.Lib_Entry_Key() = Nothing
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.CurrentPlayList()
            Catch ex As Exception
                'If g_bDebug Then Log("CurrentPlayList called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        If SuperDebug Or DCORMEDIAAPITrace Then Log("CurrentPlayList called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Dim PlaylistTracks As String() = GetCurrentPlaylistTracks()
        If PlaylistTracks IsNot Nothing Then
            ReDim ReturnList(UBound(PlaylistTracks))
            Dim Index As Integer = 0
            For Each Track As String In PlaylistTracks
                Dim EntryKey As New HomeSeerAPI.Lib_Entry_Key
                EntryKey.iKey = Index + 1
                EntryKey.Library = MyLibraryTypes.LibraryQueue
                EntryKey.Title = Track
                EntryKey.sKey = (Index + 1).ToString
                EntryKey.WhichKey = eKey_Type.eEither
                ReturnList(Index) = EntryKey
                Index = Index + 1
            Next
            If SuperDebug Or DCORMEDIAAPITrace Then Log("CurrentPlayList called for Zone - " & ZoneName & " and returned = " & Index.ToString & " entries", LogType.LOG_TYPE_INFO, LogColorNavy)
        End If
        Return ReturnList
    End Function

    Public Function CurrentPlayListAdd(TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.CurrentPlayListAdd
        If g_bDebug Then Log("CurrentPlayListAdd called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return Nothing
    End Function

    Public Sub CurrentPlayListClear() Implements HomeSeerAPI.IMediaAPI.CurrentPlayListClear
        If g_bDebug Then Log("CurrentPlayListClear called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public Function CurrentPlayListCount() As Integer Implements HomeSeerAPI.IMediaAPI.CurrentPlayListCount
        If g_bDebug Then Log("CurrentPlayListCount called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return MyNbrOfTracksInQueue
    End Function

    Public Function CurrentPlayListRange(Start As Integer, Count As Integer) As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.CurrentPlayListRange
        If g_bDebug Then Log("CurrentPlayListRange called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return Nothing
    End Function

    Public Function CurrentPlayListSet(TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.CurrentPlayListSet
        If g_bDebug Then Log("CurrentPlayListSet called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return Nothing
    End Function

    Public Function LibGetAlbums(artist As String, genre As String, Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetAlbums
        If g_bDebug Then Log("LibGetAlbums called for Zone - " & ZoneName & " with Artist = " & artist & ", Genre = " & genre & ", Lib_type = " & Lib_Type.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return GetAlbums(artist, genre)
    End Function

    Public Function LibGetArtists(album As String, genre As String, Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetArtists
        If g_bDebug Then Log("LibGetArtists called for Zone - " & ZoneName & " with Album = " & album & ", Genre = " & genre & ", Lib_type = " & Lib_Type.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return GetArtists(album, genre)
    End Function

    Public Function LibGetEntry(Key As HomeSeerAPI.Lib_Entry_Key) As HomeSeerAPI.Lib_Entry Implements HomeSeerAPI.IMediaAPI.LibGetEntry
        LibGetEntry = Nothing
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.LibGetEntry(Key)
            Catch ex As Exception
                If g_bDebug Then Log("LibGetEntry called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        If SuperDebug Or DCORMEDIAAPITrace Then Log("LibGetEntry called for Zone - " & ZoneName & ", " & TranslateLibEntryKey(Key), LogType.LOG_TYPE_INFO, LogColorNavy)
        Dim insKey As String = ""
        Dim inTitle As String = ""
        If Key.sKey IsNot Nothing Then
            insKey = Key.sKey
        End If
        If Key.Title IsNot Nothing Then
            inTitle = Key.Title
        End If
        If Key.Library = MyLibraryTypes.LibraryQueue Then
            ' assume this to be the current queue
            Select Case Key.WhichKey
                Case eKey_Type.eEither, eKey_Type.eNumber
                    Return GetQueueElement(Key.iKey, inTitle)
                Case eKey_Type.eString
                    Dim KeyID As Integer = 0
                    Try
                        KeyID = Val(insKey)
                    Catch ex As Exception
                    End Try
                    Return GetQueueElement(KeyID, inTitle)
            End Select
        ElseIf Key.Library = MyLibraryTypes.LibraryDB Then
            Return Nothing
        End If
    End Function

    Public Function LibGetGenres(Lib_Type As UShort) As String() Implements HomeSeerAPI.IMediaAPI.LibGetGenres
        If g_bDebug Then Log("LibGetGenres called for Zone - " & ZoneName & " with Lib_type = " & Lib_Type.ToString, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return GetGenres("") 'Nothing
    End Function

    Public Function LibGetLibrary() As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibrary
        If g_bDebug Then Log("LibGetLibrary called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibrarybyLibType(Lib_Type As UShort) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibrarybyLibType
        If g_bDebug Then Log("LibGetLibrarybyLibType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryCount() As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCount
        If g_bDebug Then Log("LibGetLibraryCount called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryCountbyEntryType(EntryType As HomeSeerAPI.eLib_Media_Type) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyEntryType
        If g_bDebug Then Log("LibGetLibraryCountbyEntryType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryCountbyLibType(Lib_Type As UShort) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyLibType
        If g_bDebug Then Log("LibGetLibraryCountbyLibType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryCountbyType(Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) As Integer Implements HomeSeerAPI.IMediaAPI.LibGetLibraryCountbyType
        If g_bDebug Then Log("LibGetLibraryCountbyType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryRange(Start As Integer, Count As Integer) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRange
        If g_bDebug Then Log("LibGetLibraryRange called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyEntryType(Start As Integer, Count As Integer, EntryType As HomeSeerAPI.eLib_Media_Type) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyEntryType
        If g_bDebug Then Log("LibGetLibraryRangebyEntryType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyLibType(Start As Integer, Count As Integer, LibType As UShort) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyLibType
        If g_bDebug Then Log("LibGetLibraryRangebyLibType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryRangebyType(Start As Integer, Count As Integer, LibType As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryRangebyType
        If g_bDebug Then Log("LibGetLibraryRangebyType called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetLibraryTypes() As HomeSeerAPI.Lib_Type() Implements HomeSeerAPI.IMediaAPI.LibGetLibraryTypes
        If g_bDebug Then Log("LibGetLibraryTypes called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorOrange)
        Return Nothing
    End Function

    Public Function LibGetPlaylists(Optional Lib_Type As UShort = 0) As HomeSeerAPI.Playlist_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetPlaylists
        If g_bDebug Then Log("LibGetPlaylists called for Zone - " & ZoneName & " with Lib_Type = " & Lib_Type.ToString, LogType.LOG_TYPE_INFO, LogColorOrange)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.LibGetPlaylists(Lib_Type)
            Catch ex As Exception
                If g_bDebug Then Log("LibGetPlaylists called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        Dim ReturnList As HomeSeerAPI.Playlist_Entry() = Nothing
        Dim Playlists As String() = GetPlaylists()
        If Playlists IsNot Nothing Then
            ReDim ReturnList(UBound(Playlists))
            Dim Index As Integer = 0
            For Each Playlist As String In Playlists
                Dim EntryKey As New HomeSeerAPI.Playlist_Entry
                EntryKey.Length = 2
                EntryKey.Lib_Type = MyLibraryTypes.LibraryDB
                EntryKey.Playlist_Key = Index + 1
                EntryKey.Playlist_Name = Playlist
                ReturnList(Index) = EntryKey
                Index = Index + 1
            Next
            If g_bDebug Then Log("LibGetPlaylists called for Zone - " & ZoneName & " and returned = " & Index.ToString & " entries", LogType.LOG_TYPE_INFO, LogColorOrange)
        End If
        Return ReturnList
    End Function

    Public Function LibGetPlaylistTracks(Playlist As HomeSeerAPI.Playlist_Entry) As HomeSeerAPI.Lib_Entry() Implements HomeSeerAPI.IMediaAPI.LibGetPlaylistTracks
        If g_bDebug Then Log("LibGetPlaylistTracks called for Zone - " & ZoneName & " with Playlist = " & Playlist.Playlist_Name, LogType.LOG_TYPE_INFO, LogColorOrange)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.LibGetPlaylistTracks(Playlist)
            Catch ex As Exception
                If g_bDebug Then Log("LibGetPlaylistTracks called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        Dim ReturnList As HomeSeerAPI.Lib_Entry() = Nothing
        Dim PlaylistsTracks As System.Array = GetPlaylistTracks(Playlist.Playlist_Name)
        If PlaylistsTracks IsNot Nothing Then
            ReDim ReturnList(UBound(PlaylistsTracks))
            Dim Index As Integer = 0
            For Each PlaylistsTrack As String In PlaylistsTracks
                Dim LibEntry As New HomeSeerAPI.Lib_Entry
                LibEntry.Album = "album_" & Index.ToString
                LibEntry.Artist = "artist_" & Index.ToString
                LibEntry.Cover_Back_path = NoArtPath
                LibEntry.Cover_path = NoArtPath
                LibEntry.Genre = "genre_" & Index.ToString
                Dim LibEntryKey As New HomeSeerAPI.Lib_Entry_Key
                LibEntryKey.iKey = Index + 1
                LibEntryKey.Library = MyLibraryTypes.LibraryQueue ' changed from 0
                LibEntryKey.sKey = ""  'Index.ToString
                LibEntryKey.WhichKey = eKey_Type.eEither
                LibEntryKey.Title = PlaylistsTrack
                LibEntry.Key = LibEntryKey
                LibEntry.Kind = ""
                LibEntry.LengthSeconds = 0
                LibEntry.Lib_Media_Type = eLib_Media_Type.Music
                LibEntry.Lib_Type = MyLibraryTypes.LibraryQueue
                LibEntry.PlayedCount = 0
                LibEntry.Rating = 0
                LibEntry.Title = PlaylistsTrack
                LibEntry.Year = 0
                ReturnList(Index) = LibEntry
                Index = Index + 1
            Next
            If g_bDebug Then Log("LibGetPlaylists called for Zone - " & ZoneName & " and returned = " & Index.ToString & " entries", LogType.LOG_TYPE_INFO, LogColorOrange)
        End If
        Return ReturnList
    End Function

    Public Function LibGetTracks(artist As String, album As String, genre As String, Lib_Type As UShort) As HomeSeerAPI.Lib_Entry_Key() Implements HomeSeerAPI.IMediaAPI.LibGetTracks
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.LibGetTracks(artist, album, genre, Lib_Type)
            Catch ex As Exception
                If g_bDebug Then Log("LibGetTracks called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        If g_bDebug Then Log("LibGetTracks called for Zone - " & ZoneName & " with Artist = " & artist & ", Album = " & album & ", Genre = " & genre & ", Lib_type = " & Lib_Type.ToString, LogType.LOG_TYPE_INFO, LogColorOrange)
        Dim ReturnList As HomeSeerAPI.Lib_Entry_Key() = Nothing
        Dim Tracks As String() = DBGetTracks(artist, album, genre)
        If Tracks IsNot Nothing Then
            ReDim ReturnList(UBound(Tracks))
            Dim Index As Integer = 0
            For Each Track As String In Tracks
                Dim EntryKey As New HomeSeerAPI.Lib_Entry_Key
                EntryKey.iKey = Index + 1
                EntryKey.Library = MyLibraryTypes.LibraryDB
                EntryKey.Title = Track
                EntryKey.sKey = (Index + 1).ToString
                EntryKey.WhichKey = eKey_Type.eString
                ReturnList(Index) = EntryKey
                Index = Index + 1
            Next
            If g_bDebug Then Log("LibGetTracks called for Zone - " & ZoneName & " and returned = " & Index.ToString & " entries", LogType.LOG_TYPE_INFO, LogColorNavy)
        End If
        Return ReturnList
    End Function

    Public ReadOnly Property LibLoading As Boolean Implements HomeSeerAPI.IMediaAPI.LibLoading
        Get
            If g_bDebug Then Log("LibLoading called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
            Return False
        End Get
    End Property

    Public Sub Play(Key As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.Play
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Play(Key)
            Catch ex As Exception
                If g_bDebug Then Log("Play called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim InfoString As String = TranslateLibEntryKey(Key)
        If g_bDebug Then Log("Play called for Zone - " & ZoneName & " and Key = " & InfoString, LogType.LOG_TYPE_INFO, LogColorNavy)
        SonosPlay()
    End Sub

    Public Sub PlayGenre(GenreName As String, Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type) Implements HomeSeerAPI.IMediaAPI.PlayGenre
        If g_bDebug Then Log("PlayGenre called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public Sub PlayGenreAt(GenreName As String, Lib_Type As UShort, EntryType As HomeSeerAPI.eLib_Media_Type, Start_Track As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.PlayGenreAt
        If g_bDebug Then Log("PlayGenreAt called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public Function Playlist_Add(Playlist As HomeSeerAPI.Playlist_Entry) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add
        If g_bDebug Then Log("Playlist_Add called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Add_Track(Playlist As HomeSeerAPI.Playlist_Entry, TrackKey As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add_Track
        If g_bDebug Then Log("Playlist_Add_Track called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Add_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Add_Tracks
        If g_bDebug Then Log("Playlist_Add_Tracks called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Add_Matched_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, MatchInfo As HomeSeerAPI.Play_Match_Info) As Boolean Implements HomeSeerAPI.IMediaAPI_3.Playlist_Add_Matched_Tracks
        If g_bDebug Then Log("Playlist_Add_Matched_Tracks called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Delete(Playlist As HomeSeerAPI.Playlist_Entry) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete
        If g_bDebug Then Log("Playlist_Delete called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Delete_Track(Playlist As HomeSeerAPI.Playlist_Entry, TrackKey As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete_Track
        If g_bDebug Then Log("Playlist_Delete_Track called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Function Playlist_Delete_Tracks(Playlist As HomeSeerAPI.Playlist_Entry, TrackKeys() As HomeSeerAPI.Lib_Entry_Key) As Boolean Implements HomeSeerAPI.IMediaAPI.Playlist_Delete_Tracks
        If g_bDebug Then Log("Playlist_Delete_Tracks called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Return True
    End Function

    Public Sub PlayPlaylist(Playlist As HomeSeerAPI.Playlist_Entry) Implements HomeSeerAPI.IMediaAPI.PlayPlaylist
        If g_bDebug Then Log("PlayPlaylist called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public Sub PlayPlaylistAt(Playlist As HomeSeerAPI.Playlist_Entry, Start_Key As HomeSeerAPI.Lib_Entry_Key) Implements HomeSeerAPI.IMediaAPI.PlayPlaylistAt
        If g_bDebug Then Log("PlayPlaylistAt called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public Sub PlayMatch(MatchInfo As HomeSeerAPI.Play_Match_Info) Implements HomeSeerAPI.IMediaAPI_2.PlayMatch
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.PlayMatch(MatchInfo)
            Catch ex As Exception
                If g_bDebug Then Log("PlayMatch called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        Dim InfoString As String = " "
        Dim inIDs As Integer()
        Dim inLibType As Integer = MatchInfo.L0_Lib_Type
        Dim inGenre As String = ""
        Dim inArtist As String = ""
        Dim inAlbum As String = ""
        Dim inTitle As String = ""
        Dim inLibMediaType As String = ""
        Dim inPlaylist As String = ""
        Try
            If MatchInfo.IDs IsNot Nothing Then
                If MatchInfo.IDs.Length > 0 Then
                    inIDs = MatchInfo.IDs
                    Dim Index As Integer = 0
                    For Each ID In MatchInfo.IDs
                        InfoString &= " ID" & Index.ToString & " = " & ID.ToString
                        Index += 1
                    Next
                End If
            End If
            InfoString &= " L0_Lib_Type = " & MatchInfo.L0_Lib_Type.ToString
            If MatchInfo.L1_Genre IsNot Nothing Then
                InfoString &= ", L1_Genre = " & MatchInfo.L1_Genre
                inGenre = MatchInfo.L1_Genre
            End If
            If MatchInfo.L2_Artist IsNot Nothing Then
                InfoString &= ", L2_Artist = " & MatchInfo.L2_Artist
                inArtist = MatchInfo.L2_Artist
            End If
            If MatchInfo.L3_Album IsNot Nothing Then
                InfoString &= ", L3_Album = " & MatchInfo.L3_Album
                inAlbum = MatchInfo.L3_Album
            End If
            If MatchInfo.L4_Title IsNot Nothing Then
                InfoString &= ", L4_Title = " & MatchInfo.L4_Title
                inTitle = MatchInfo.L4_Title
            End If
            InfoString &= ", Lib_Media_Type = " & MatchInfo.Lib_Media_Type.ToString
            If MatchInfo.Playlist IsNot Nothing Then
                InfoString &= ", Playlist = " & MatchInfo.Playlist
                inPlaylist = MatchInfo.Playlist
            End If
        Catch ex As Exception
            Log("Error in PlayMatch for Zone - " & ZoneName & " with Error = " & ex.ToString, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("PlayMatch called for Zone - " & ZoneName & InfoString, LogType.LOG_TYPE_INFO, LogColorNavy)
        PlayMusic(inArtist, inAlbum, inPlaylist, inGenre, inTitle, "", "", "", "", "", False, QueueActions.qaPlayNow)
    End Sub

    Public Sub AdjustVolume(Amount As Integer, Optional Direction As HomeSeerAPI.eVolumeDirection = HomeSeerAPI.eVolumeDirection.Absolute) Implements HomeSeerAPI.IMediaAPI_3.AdjustVolume
        If g_bDebug Then Log("AdjustVolume called for Zone - " & ZoneName & " Amount = " & Amount.ToString & ", Direction = " & Direction.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        Select Case Direction
            Case eVolumeDirection.Absolute
                PlayerVolume = Amount
            Case eVolumeDirection.Down
                PlayerVolume = Volume - Amount
            Case eVolumeDirection.Up
                PlayerVolume = Volume + Amount
            Case Else
                Log("Error in AdjustVolume called for Zone - " & ZoneName & " with unknown Direction and Amount = " & Amount.ToString & ", Direction = " & Direction.ToString, LogType.LOG_TYPE_ERROR)
        End Select
    End Sub

    Public Sub Halt() Implements HomeSeerAPI.IMediaAPI_3.Halt
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Halt()
            Catch ex As Exception
                If g_bDebug Then Log("Halt called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If g_bDebug Then Log("Halt called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        StopPlay()
    End Sub

    Public Sub Mute(Mode As HomeSeerAPI.mute_modes) Implements HomeSeerAPI.IMediaAPI_3.Mute
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Mute(Mode)
            Catch ex As Exception
                If g_bDebug Then Log("Mute called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If SuperDebug Or DCORMEDIAAPITrace Then Log("Mute called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Select Case Muted
            Case mute_modes.muted
                PlayerMute = True
            Case mute_modes.not_muted
                PlayerMute = False
        End Select
    End Sub

    Public ReadOnly Property Muted As HomeSeerAPI.mute_modes Implements HomeSeerAPI.IMediaAPI_3.Muted
        Get
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI
                Try
                    LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                    Return LinkedZone.Muted
                Catch ex As Exception
                    If g_bDebug Then Log("Muted called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Return mute_modes.not_muted
            End If
            If SuperDebug Or DCORMEDIAAPITrace Then Log("Muted called for Zone - " & ZoneName & " Mute state = " & PlayerMute.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
            If PlayerMute Then
                Return mute_modes.muted
            Else
                Return mute_modes.not_muted
            End If
        End Get
    End Property

    Public Sub Pause() Implements HomeSeerAPI.IMediaAPI_3.Pause
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Pause()
            Catch ex As Exception
                If g_bDebug Then Log("Pause called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If g_bDebug Then Log("Pause called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        SonosPause()
    End Sub

    Public ReadOnly Property PlayerPosition As Integer Implements HomeSeerAPI.IMediaAPI_3.PlayerPosition
        Get
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI
                Try
                    LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                    Return LinkedZone.PlayerPosition
                Catch ex As Exception
                    If g_bDebug Then Log("PlayerPosition called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Return 0
            End If
            If SuperDebug Or DCORMEDIAAPITrace Then Log("PlayerPosition called for Zone - " & ZoneName & " and returned = " & SonosPlayerPosition, LogType.LOG_TYPE_INFO, LogColorNavy)
            Return SonosPlayerPosition
        End Get
    End Property

    Public Sub Repeat(Mode As HomeSeerAPI.repeat_modes) Implements HomeSeerAPI.IMediaAPI_3.Repeat
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Repeat(Mode)
            Catch ex As Exception
                If g_bDebug Then Log("Repeat called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If SuperDebug Or DCORMEDIAAPITrace Then Log("Repeat called for Zone - " & ZoneName & " with Repeat = " & Mode.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
        Select Case Mode
            Case repeat_modes.repeat_all
                SonosRepeat = repeat_modes.repeat_all
            Case repeat_modes.repeat_one
                SonosRepeat = repeat_modes.repeat_one
            Case Else
                SonosRepeat = repeat_modes.repeat_off
        End Select
    End Sub

    Public ReadOnly Property Repeating As HomeSeerAPI.repeat_modes Implements HomeSeerAPI.IMediaAPI_3.Repeating
        Get
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI
                Try
                    LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                    Return LinkedZone.Repeating
                Catch ex As Exception
                    If g_bDebug Then Log("Repeating called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Return repeat_modes.repeat_off
            End If
            If SuperDebug Or DCORMEDIAAPITrace Then Log("Repeating called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
            Select Case SonosRepeat
                Case repeat_modes.repeat_all
                    Return repeat_modes.repeat_all
                Case repeat_modes.repeat_one
                    Return repeat_modes.repeat_one
                Case Else
                    Return repeat_modes.repeat_off
            End Select
        End Get
    End Property

    Public Sub SelectTrack(TrackKey As Integer, Optional PlaylistKey As Integer = -1) Implements HomeSeerAPI.IMediaAPI_3.SelectTrack
        If g_bDebug Then Log("SelectTrack called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.SelectTrack(TrackKey)
            Catch ex As Exception
                If g_bDebug Then Log("SelectTrack called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        ' still needs implementation!!
    End Sub

    Public Sub Shuffle(Mode As HomeSeerAPI.shuffle_modes) Implements HomeSeerAPI.IMediaAPI_3.Shuffle
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.Shuffle(Mode)
            Catch ex As Exception
                If g_bDebug Then Log("Shuffle called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If SuperDebug Then Log("Shuffle called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
        Select Case Mode
            Case HomeSeerAPI.shuffle_modes.not_shuffled
                SonosShuffle = Shuffle_modes.Ordered
            Case HomeSeerAPI.shuffle_modes.shuffled
                SonosShuffle = Shuffle_modes.Shuffled
            Case HomeSeerAPI.shuffle_modes.sorted
                SonosShuffle = Shuffle_modes.Sorted
        End Select
    End Sub

    Public ReadOnly Property Shuffled As HomeSeerAPI.shuffle_modes Implements HomeSeerAPI.IMediaAPI_3.Shuffled
        Get
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI
                Try
                    LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                    Return LinkedZone.Shuffled
                Catch ex As Exception
                    If g_bDebug Then Log("Shuffled called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Return HomeSeerAPI.shuffle_modes.not_shuffled
            End If
            If SuperDebug Or DCORMEDIAAPITrace Then Log("Shuffled called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)
            Select Case SonosShuffle
                Case Shuffle_modes.Ordered
                    Return HomeSeerAPI.shuffle_modes.not_shuffled
                Case Shuffle_modes.Shuffled
                    Return HomeSeerAPI.shuffle_modes.shuffled
                Case Else
                    Return HomeSeerAPI.shuffle_modes.sorted
            End Select
        End Get
    End Property

    Public Sub SkipToTrack(TrackName As String) Implements HomeSeerAPI.IMediaAPI_3.SkipToTrack
        If g_bDebug Then Log("SkipToTrack called for Zone - " & ZoneName & " with Trackname = " & TrackName, LogType.LOG_TYPE_INFO, LogColorNavy)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.SkipToTrack(TrackName)
            Catch ex As Exception
                If g_bDebug Then Log("SkipToTrack called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        SonosSkipToTrack(TrackName)
    End Sub

    Public Sub SkipTracks(SkipValue As Integer) Implements HomeSeerAPI.IMediaAPI_3.SkipTracks
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                LinkedZone.SkipTracks(SkipValue)
            Catch ex As Exception
                If g_bDebug Then Log("SkipTracks called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Sub
        End If
        If g_bDebug Then Log("SkipTracks called for Zone - " & ZoneName & " with SkipValue = " & SkipValue.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
    End Sub

    Public ReadOnly Property State As HomeSeerAPI.player_state_values Implements HomeSeerAPI.IMediaAPI_3.State
        Get
            If ZoneIsLinked Then
                Dim LinkedZone As HSPI
                Try
                    LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                    Return LinkedZone.State
                Catch ex As Exception
                    'If g_bDebug Then Log("State called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                Return HomeSeerAPI.player_state_values.stopped
            End If
            If SuperDebug Or DCORMEDIAAPITrace Then Log("State called for Zone - " & ZoneName & " and State = " & PlayerState.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
            Select Case PlayerState()
                Case player_state_values.Playing
                    Return HomeSeerAPI.player_state_values.playing
                Case player_state_values.Paused
                    Return HomeSeerAPI.player_state_values.paused
                Case player_state_values.Stopped
                    Return HomeSeerAPI.player_state_values.stopped
                Case player_state_values.Forwarding
                    Return HomeSeerAPI.player_state_values.forwarding
                Case player_state_values.Rewinding
                    Return HomeSeerAPI.player_state_values.rewinding
                Case Else
                    Return HomeSeerAPI.player_state_values.stopped
            End Select
        End Get
    End Property

    Public ReadOnly Property Volume As Integer Implements HomeSeerAPI.IMediaAPI_3.Volume
        Get ' Returns the current volume setting of the player from 0 to 100.
            'If SuperDebug  Or DCORMEDIAAPITrace Then Log("Get Volume called for Zone - " & ZoneName & " returning value = " & MyCurrentMasterVolumeLevel.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
            Volume = MyCurrentMasterVolumeLevel
        End Get
    End Property

    Public Function GetMatch(QData As HomeSeerAPI.Query_Object) As HomeSeerAPI.Response_Object Implements HomeSeerAPI.IMediaAPI_3.GetMatch
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
                Return LinkedZone.GetMatch(QData)
            Catch ex As Exception
                If g_bDebug Then Log("GetMatch called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Return Nothing
        End If
        If g_bDebug Then Log("GetMatch called for Zone - " & ZoneName, LogType.LOG_TYPE_INFO, LogColorNavy)

        Try
            If QData.QueryData IsNot Nothing Then
                If QData.QueryData.Length > 0 Then
                    Dim Index As Integer = 0
                    Dim responsestring As String = ""
                    For Each inQueryData As HomeSeerAPI.Query_Data In QData.QueryData
                        If inQueryData.Data IsNot Nothing Then
                            If inQueryData.Data.Length > 0 Then
                                Dim DataIndex As Integer = 0
                                For Each DElemenent As String In inQueryData.Data
                                    responsestring = " Data" & Index.ToString & "_" & DataIndex.ToString & " = " & DElemenent
                                    If g_bDebug Then Log("GetMatch called for Zone - " & ZoneName & " with QueryData = " & responsestring & ", Datatype = " & inQueryData.DataType.ToString & ", isRange = " & inQueryData.IsRange.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
                                    DataIndex += 1
                                Next
                            End If
                        End If
                        Index += 1
                    Next
                End If
            End If
            If QData.ResponseData IsNot Nothing Then
                If QData.ResponseData.Length > 0 Then
                    Dim Index As Integer = 0
                    For Each inResponseData As HomeSeerAPI.eQRDataType In QData.ResponseData
                        If g_bDebug Then Log("GetMatch called for Zone - " & ZoneName & " with ResponseData_" & Index.ToString & " = " & inResponseData.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
                        Index += 1
                    Next
                End If
            End If
        Catch ex As Exception

        End Try
        Try
            Dim Response As New HomeSeerAPI.Response_Object
            Dim ResponseStr As New HomeSeerAPI.Response_Data
            ResponseStr.AdditionalInfo = Nothing
            ResponseStr.Key = 1
            ResponseStr.Title = ""
            Response.Data(0) = ResponseStr
            Response.DataTypes(0) = eQRDataType.Albums
            Return Response

        Catch ex As Exception

        End Try
        Return Nothing

    End Function

    Private Function TranslateLibEntryKey(Key As HomeSeerAPI.Lib_Entry_Key) As String
        TranslateLibEntryKey = ""
        Try
            TranslateLibEntryKey = "Key.iKey = " & Key.iKey.ToString
            If Key.sKey IsNot Nothing Then
                TranslateLibEntryKey &= ", Key.sKey = " & Key.sKey.ToString
            Else
                TranslateLibEntryKey &= ", Key.sKey = "
            End If
            TranslateLibEntryKey &= ", Key.Library = " & Key.Library.ToString
            If Key.Title IsNot Nothing Then
                TranslateLibEntryKey &= ", Key.Title = " & Key.Title.ToString
            Else
                TranslateLibEntryKey &= ", Key.Title = "
            End If
            TranslateLibEntryKey &= ", Key.WhichKey = " & Key.WhichKey.ToString
        Catch ex As Exception
        End Try
    End Function

End Class



<Serializable()> _
Public Class LastMusic
    Public Sub New()

    End Sub
    Public Album As String
    Public Artist As String
    Public Genre As String
    Public iMode As Integer
    Public Playlist As String
    Public Track As String
    Public WasPlaying As Boolean
End Class

Public Class track_desc
    Public name As String
    Public artist As String
    Public album As String
    Public length As String
End Class
