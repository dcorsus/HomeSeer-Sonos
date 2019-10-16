Imports HomeSeerAPI
Imports Scheduler
Imports System.Web.UI.WebControls
Imports System.Web.UI
Imports System.Text
Imports System.Web
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D

Class PlayerControl
    Inherits clsPageBuilder

    Private PIReference As HSPI = Nothing
    Private MyZoneUDN As String = ""
    Private MusicAPI As HSPI = Nothing   'HSMusicAPI = Nothing
    Private ar() As String
    Private arplaylist() As String
    Private ZoneName As String = ""
    Private ZoneModel As String = ""
    Private iPodPlayerName As String = ""
    Private MyPageName As String = ""

    Private LblTrackName As String = ""
    Private LblArtistName As String = ""
    Private LblAlbumName As String = ""
    Private LblRadiostationName As String = ""
    Private LblNextTrackName As String = ""
    Private LblNextArtistName As String = ""
    Private LblNextAlbumName As String = ""
    Private LblDuration As Integer = 0
    Private LblDescr As String = ""
    Private LblWhatsLoaded As String = ""
    Private LblLinkState As String = ""
    Private ButRepeatImagePath As String = ""
    Private ButShuffleImagePath As String = ""
    Private ButMuteImagePath As String = ""
    'Private ArtImagePath As String = ""
    Private LblWhatWasPreviouslyLoaded As String = ""

    Private MyPosition As Integer = 0
    Private MyVolume As Integer = 0
    Private MyTrack As String = ""
    Private MyAlbum As String = ""
    Private MyArtist As String = ""
    Private MyRadiostation As String = ""
    Private MyNextTrack As String = ""
    Private MyNextAlbum As String = ""
    Private MyNextArtist As String = ""
    Private MyArt As String = ""
    Private MyMute As Boolean = False
    Private MyRepeat As Boolean = False
    Private MyShuffle As Boolean = False
    Private MyPlayerState As player_state_values = player_state_values.Stopped
    Private MyIPodName As String = ""
    Private MyPlayerisPairMaster As Boolean = False
    Private MyPlayerisPairSlave As Boolean = False

    Private MyLastSelecteNavBoxItems As String = ""
    Private MyLastSelectedNavBoxClass As String = ""

    Private ButPrev As clsJQuery.jqButton
    Private ButStop As clsJQuery.jqButton
    Private ButPlay As clsJQuery.jqButton
    Private ButNext As clsJQuery.jqButton
    Private ButMute As clsJQuery.jqButton
    Private ButLoudness As clsJQuery.jqButton
    Private ButRepeat As clsJQuery.jqButton
    Private ButShuffle As clsJQuery.jqButton
    Private ButGenres As clsJQuery.jqButton
    Private ButArtists As clsJQuery.jqButton
    Private ButAlbums As clsJQuery.jqButton
    Private ButTracks As clsJQuery.jqButton
    Private ButPlayLists As clsJQuery.jqButton
    Private ButRadioStations As clsJQuery.jqButton
    Private ButAudioBooks As clsJQuery.jqButton
    Private ButPodCasts As clsJQuery.jqButton
    Private ButFavorites As clsJQuery.jqButton
    Private ButGroup As clsJQuery.jqMultiSelect
    Private ButLineInput As clsJQuery.jqButton
    Private ButPair As clsJQuery.jqButton
    Private ButUnpair As clsJQuery.jqButton
    Private ButAddToPlaylist As clsJQuery.jqButton
    Private ButClearList As clsJQuery.jqButton
    Private ButEditor As clsJQuery.jqButton
    Private LblGenre As clsJQuery.jqButton
    Private LblArtist As clsJQuery.jqButton
    Private LblAlbum As clsJQuery.jqButton
    Private LblPlaylist As clsJQuery.jqButton
    Private LblRadioList As clsJQuery.jqButton
    Private LblAudiobooks As clsJQuery.jqButton
    Private LblPodcasts As clsJQuery.jqButton
    Private LblFavorites As clsJQuery.jqButton
    Private LblLineInput As clsJQuery.jqButton
    Private LblPairing As clsJQuery.jqButton
    Private NavigationBox As clsJQuery.jqListBoxEx
    Private PlaylistBox As clsJQuery.jqListBoxEx
    Private ArtImageBox As clsJQuery.jqButton
    Private GroupingOverlay As clsJQuery.jqOverlay
    Private ButTV As clsJQuery.jqButton
    Private MyClientID As String = ""

    Public Sub New(ByVal pagename As String, ClientID As String)
        MyBase.New(pagename)
        MyPageName = pagename
        If ClientID <> "" Then MyClientID = "?clientid=" & ClientID Else MyClientID = ""
        NavigationBox = New clsJQuery.jqListBoxEx("NavigationBox", MyPageName & MyClientID)
        PlaylistBox = New clsJQuery.jqListBoxEx("PlaylistBox", MyPageName & MyClientID)
        ArtImageBox = New clsJQuery.jqButton("ArtImageBox", "", MyPageName & MyClientID, True)
        ButPrev = New clsJQuery.jqButton("ButPrev", "Previous", MyPageName & MyClientID, True)
        ButStop = New clsJQuery.jqButton("ButStop", "Stop", MyPageName & MyClientID, True)
        ButPlay = New clsJQuery.jqButton("ButPlay", "Play", MyPageName & MyClientID, True)
        ButNext = New clsJQuery.jqButton("ButNext", "Next", MyPageName & MyClientID, True)
        ButMute = New clsJQuery.jqButton("ButMute", "Mute", MyPageName & MyClientID, True)
        ButLoudness = New clsJQuery.jqButton("ButLoudness", "Loudness", MyPageName & MyClientID, True)
        ButRepeat = New clsJQuery.jqButton("ButRepeat", "Repeat", MyPageName & MyClientID, True)
        ButShuffle = New clsJQuery.jqButton("ButShuffle", "Shuffle", MyPageName & MyClientID, True)
        ButGenres = New clsJQuery.jqButton("ButGenres", "Genres", MyPageName & MyClientID, True)
        ButArtists = New clsJQuery.jqButton("ButArtists", "Artists", MyPageName & MyClientID, True)
        ButAlbums = New clsJQuery.jqButton("ButAlbums", "Albums", MyPageName & MyClientID, True)
        ButTracks = New clsJQuery.jqButton("ButTracks", "Tracks", MyPageName & MyClientID, True)
        ButPlayLists = New clsJQuery.jqButton("ButPlayLists", "Playlists", MyPageName & MyClientID, True)
        ButRadioStations = New clsJQuery.jqButton("ButRadioStations", "Radiostations", MyPageName & MyClientID, True)
        ButAudioBooks = New clsJQuery.jqButton("ButAudioBooks", "Audiobooks", MyPageName & MyClientID, True)
        ButPodCasts = New clsJQuery.jqButton("ButPodCasts", "Podcasts", MyPageName & MyClientID, True)
        ButFavorites = New clsJQuery.jqButton("ButFavorites", "Favorites", MyPageName & MyClientID, True)
        ButGroup = New clsJQuery.jqMultiSelect("ButGroup", MyPageName & MyClientID, True)
        ButLineInput = New clsJQuery.jqButton("ButLineInput", "Line Input", MyPageName & MyClientID, True)
        ButPair = New clsJQuery.jqButton("ButPair", "Pair", MyPageName & MyClientID, True)
        ButUnpair = New clsJQuery.jqButton("ButUnpair", "Unpair", MyPageName & MyClientID, True)
        ButAddToPlaylist = New clsJQuery.jqButton("ButAddToPlaylist", "Add To Playlist", MyPageName & MyClientID, True)
        ButClearList = New clsJQuery.jqButton("ButClearList", "Clear Queue", MyPageName & MyClientID, True)
        ButEditor = New clsJQuery.jqButton("ButEditor", "ButEditor", MyPageName & MyClientID, True)
        LblGenre = New clsJQuery.jqButton("LblGenre", "", MyPageName & MyClientID, True)
        LblGenre.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Artists ...</h2>' });"
        LblArtist = New clsJQuery.jqButton("LblArtist", "", MyPageName & MyClientID, True)
        LblArtist.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Albums ...</h2>' });"
        LblAlbum = New clsJQuery.jqButton("LblAlbum", "", MyPageName & MyClientID, True)
        LblAlbum.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Tracks ...</h2>' });"
        LblPlaylist = New clsJQuery.jqButton("LblPlaylist", "", MyPageName & MyClientID, True)
        'LblPlaylist.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Playlist ...</h2>' });"
        LblRadioList = New clsJQuery.jqButton("LblRadioList", "", MyPageName & MyClientID, True)
        LblAudiobooks = New clsJQuery.jqButton("LblAudiobooks", "", MyPageName & MyClientID, True)
        LblPodcasts = New clsJQuery.jqButton("LblPodcasts", "", MyPageName & MyClientID, True)
        LblFavorites = New clsJQuery.jqButton("LblFavorites", "", MyPageName & MyClientID, True)
        LblLineInput = New clsJQuery.jqButton("LblLineInput", "", MyPageName & MyClientID, True)
        LblPairing = New clsJQuery.jqButton("LblPairing", "", MyPageName & MyClientID, True)
        GroupingOverlay = New clsJQuery.jqOverlay("GroupOverlay", MyPageName & MyClientID, False, "events_overlay")
        ButTV = New clsJQuery.jqButton("ButTV", "TV", MyPageName & MyClientID, True)
    End Sub

    Public WriteOnly Property RefToPlugIn
        Set(value As Object)
            PIReference = value
        End Set
    End Property

    Public WriteOnly Property ZoneUDN As String
        Set(value As String)
            MyZoneUDN = value
            Try
                MusicAPI = PIReference.GetAPIByUDN(MyZoneUDN)
            Catch ex As Exception
                Log("Error in GetPagePlugin getting the MusicAPI for ZoneUDN = " & MyZoneUDN & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Property
            End Try
            If MusicAPI Is Nothing Then
                Log("Error in GetPagePlugin, MusicAPI not found for ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_ERROR)
                Exit Property
            End If
            ZoneName = MusicAPI.GetZoneName
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for Zoneplayer = " & ZoneName & " set ZoneUDN = " & MyZoneUDN, LogType.LOG_TYPE_INFO)
        End Set
    End Property

    ' build and return the actual page
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, GenerateHeaderFooter As Boolean) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for PlayerControl called for Zoneplayer = " & ZoneName & " with pageName = " & pageName.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and queryString = " & queryString.ToString, LogType.LOG_TYPE_INFO)
        Dim stb As New StringBuilder

        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for PlayerControl called for ZoneUDN = " & MyZoneUDN & " and ZoneName = " & ZoneName & " and PageName = " & MyPageName, LogType.LOG_TYPE_INFO)

        Try
            ' handle any queries like mode=something
            Dim parts As Collections.Specialized.NameValueCollection = Nothing
            If (queryString <> "") Then
                parts = HttpUtility.ParseQueryString(queryString)
            End If

            Try
                ZoneName = MusicAPI.ZonePlayerName
                ZoneModel = MusicAPI.ZoneModel
                iPodPlayerName = MusicAPI.DockediPodPlayerName
            Catch ex As Exception
                Log("Error in GetPagePlugin getting ZoneName. Error  = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Return ""
            End Try
            Dim ClientID As String = ""
            Dim Control As String = ""
            Dim UpdateTime As String = "2000"
            Try
                If parts IsNot Nothing Then
                    If parts.HasKeys Then
                        For Each Key As String In parts.AllKeys
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin found Part = " & Key & " with Value =  " & parts(Key), LogType.LOG_TYPE_INFO)
                            If Key.ToUpper = "INSTANCE" Then

                            ElseIf Key.ToUpper = "CONTROL" Then
                                Control = parts(Key)
                            ElseIf Key.ToUpper = "CLIENTID" Then
                                ClientID = parts(Key)
                            ElseIf Key.ToUpper = "UPDATETIME" Then
                                UpdateTime = parts(Key)
                            Else
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Warning in GetPagePlugin. Found unknown Part = " & Key & " with Value =  " & parts(Key), LogType.LOG_TYPE_WARNING)
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                Log("Error in GetPagePlugin processing parts. Error  = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If ClientID <> "" Then MyClientID = "?clientid=" & ClientID Else MyClientID = ""

            ButTV.visible = False
            If MusicAPI.CheckPlayerCanPlayTV(ZoneModel) Then   ' changed on 7/12/2019 in v3.1.0.31
                ButTV.visible = True
            ElseIf Not MusicAPI.CheckPlayerIsPairable(ZoneModel) Then
                ButPair.visible = False
                ButUnpair.visible = False
            Else
                MyPlayerisPairMaster = MusicAPI.ZoneIsPairMaster
                MyPlayerisPairSlave = MusicAPI.ZoneIsPairSlave
                If MyPlayerisPairMaster Then
                    ButPair.visible = False
                    ButUnpair.visible = True
                    ButAudioBooks.visible = False
                    ButPodCasts.visible = False
                ElseIf MyPlayerisPairSlave Then
                    ButAudioBooks.visible = False
                    ButPodCasts.visible = False
                    ButPair.visible = False
                    ButUnpair.visible = True
                Else
                    ButPair.visible = True
                    ButUnpair.visible = False
                End If
            End If

            LblWhatsLoaded = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_loaded", ITEM_TYPE.artists.ToString)
            MyLastSelectedNavBoxClass = LblWhatsLoaded
            LblGenre.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_genre", "")
            LblArtist.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_artist", "")
            LblAlbum.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_album", "")
            LblPlaylist.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_playlist", "")
            LblRadioList.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_radiolist", "")
            LblAudiobooks.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_audiobook", "")
            LblPodcasts.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_podcasts", "")
            LblFavorites.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_favorites", "")
            LblLineInput.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_lineinput", "")
            LblPairing.label = GetStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_pairing", "")
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetPagePlugin for PlayerControl called for Zoneplayer = " & ZoneName & " has labels Genre = " & LblGenre.label & ", Artist=" & LblArtist.label & ", Album=" & LblAlbum.label & ", Playlist=" & LblPlaylist.label & ", RadioList=" & LblRadioList.label & ", Audiobooks=" & LblAudiobooks.label & ", Podcasts=" & LblPodcasts.label & ", Favorites=" & LblFavorites.label & ", Lineinput=" & LblLineInput.label & ", Pair=" & LblPairing.label, LogType.LOG_TYPE_INFO)
            LblWhatWasPreviouslyLoaded = LblWhatsLoaded

            If ZoneModel <> "WD100" Then
                LblAudiobooks.visible = False
                LblPodcasts.visible = False
                ButAudioBooks.visible = False
                ButPodCasts.visible = False
            Else
                ButAddToPlaylist.enabled = False
                ButClearList.enabled = False
            End If

            ar = Nothing
            arplaylist = Nothing
            If Control.ToUpper = "" Or Control.ToUpper = "NAVPANE" Then
                Select Case LblWhatsLoaded
                    Case ITEM_TYPE.genres.ToString
                        ar = MusicAPI.GetGenres()
                        LoadNavigationBox()
                    Case ITEM_TYPE.artists.ToString
                        ar = MusicAPI.GetArtists("", System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        LoadNavigationBox()
                    Case ITEM_TYPE.albums.ToString
                        ar = MusicAPI.GetAlbums(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        LoadNavigationBox()
                    Case ITEM_TYPE.tracks.ToString
                        If LblPlaylist.label <> "" Then
                            ' a playlist is loaded
                            ar = MusicAPI.GetPlaylistTracks(System.Web.HttpUtility.HtmlDecode(LblPlaylist.label))
                        Else
                            ar = MusicAPI.DBGetTracks(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        End If
                        LoadNavigationBox()
                    Case ITEM_TYPE.playlists.ToString
                        ar = MusicAPI.GetPlaylists()
                        LoadNavigationBox()
                    Case ITEM_TYPE.radioLists.ToString
                        ar = MusicAPI.LibGetRadioStationlists()
                        LoadNavigationBox()
                    Case ITEM_TYPE.Audiobooks.ToString
                        ar = MusicAPI.LibGetAudiobookslists(iPodPlayerName)
                        LoadNavigationBox()
                    Case ITEM_TYPE.Podcasts.ToString
                        ar = MusicAPI.LibGetPodcastlists(iPodPlayerName)
                        LoadNavigationBox()
                    Case ITEM_TYPE.Favorites.ToString
                        ar = MusicAPI.LibGetObjectslist("FV:2")
                        LoadNavigationBox()
                    Case ITEM_TYPE.LineInput.ToString
                        ar = MusicAPI.LibGetActiveLineInputlists()
                        LoadNavigationBox()
                    Case Else

                End Select


            End If
            If Control.ToUpper = "" Or Control.ToUpper = "QUEUE" Then
                arplaylist = MusicAPI.GetCurrentPlaylistTracks
                LoadPlayListBox()
            End If

            UpdateStatus()
            Me.reset()

            If Control <> "" Then
                ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
                'Me.AddBody(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
                'Me.AddBody(clsPageBuilder.DivEnd) ' ErrorMessage
                Select Case Control.ToUpper
                    Case "NAVPANE"
                        Me.RefreshIntervalMilliSeconds = UpdateTime
                        Me.AddBody(Me.AddAjaxHandlerPost("action=updatetime", MyPageName & MyClientID & "&control=navpane"))
                        Me.AddBody(clsPageBuilder.FormStart("NavPaneform", MyPageName & MyClientID & "&control=navpane", "post"))
                        Me.AddBody(clsPageBuilder.DivStart(MyPageName & MyClientID & "&control=navpane", ""))
                        Me.AddBody(GenerateNavPane())
                    Case "QUEUE"
                        Me.RefreshIntervalMilliSeconds = UpdateTime
                        Me.AddBody(Me.AddAjaxHandlerPost("action=updatetime", MyPageName & MyClientID & "&control=queue"))
                        Me.AddBody(clsPageBuilder.FormStart("Queueform", MyPageName & MyClientID & "&control=queue", "post"))
                        Me.AddBody(clsPageBuilder.DivStart(MyPageName & MyClientID & "&control=queue", ""))
                        Me.AddBody(GenerateQueue())
                    Case "NAVTREE"
                        Me.RefreshIntervalMilliSeconds = UpdateTime
                        Me.AddBody(Me.AddAjaxHandlerPost("action=updatetime", MyPageName & MyClientID & "&control=navtree"))
                        Me.AddBody(clsPageBuilder.FormStart("NavTreeform", MyPageName & MyClientID & "&control=navtree", "post"))
                        Me.AddBody(clsPageBuilder.DivStart(MyPageName & MyClientID & "&control=navtree", ""))
                        Me.AddBody(GenerateNavTree())
                    Case "NAVCONTROL"
                        Me.RefreshIntervalMilliSeconds = UpdateTime
                        Me.AddBody(Me.AddAjaxHandlerPost("action=updatetime", MyPageName & MyClientID & "&control=navcontrol"))
                        Me.AddBody(clsPageBuilder.FormStart("NavControlform", MyPageName & MyClientID & "&control=navcontrol", "post"))
                        Me.AddBody(clsPageBuilder.DivStart(MyPageName & MyClientID & "&control=navcontrol", ""))
                        Me.AddBody(GenerateNavControl())
                End Select

                Me.AddBody(clsPageBuilder.DivEnd) ' page end
                Me.AddBody(clsPageBuilder.FormEnd)

                Me.suppressDefaultFooter = True
                Return Me.BuildPage()
            End If

            If GenerateHeaderFooter Then
                Me.AddHeader(hs.GetPageHeader(MyPageName & MyClientID, "Player Control", "", "", False, True))
            End If

            stb.Append(clsPageBuilder.FormStart("PlayerConfigform", MyPageName & MyClientID, "post"))

            stb.Append(clsPageBuilder.DivStart(MyPageName & MyClientID, ""))
            Me.RefreshIntervalMilliSeconds = UpdateTime
            stb.Append(Me.AddAjaxHandlerPost("action=updatetime", MyPageName & MyClientID))
            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd) ' ErrorMessage

            stb.Append(clsPageBuilder.DivStart("PlayerPanel", "style='color:#0000FF'"))
            stb.Append("<table border='0' cellpadding='0' cellspacing='0' height='389px'><tr>")
            stb.Append("<td style='height: 389px; width: 860px; background-repeat:no-repeat;background-position:left;' background='" & ImagesPath & "Player-Green.png'>")

            stb.Append("<div style='position: relative'>") '; top: -150px'>")


            stb.Append("<div style='position: absolute; left: 300px; top: -185px'>")
            stb.Append("<img src='" & MusicAPI.PlayerIconURL & "' style='height:50px; width:50px'>")    ' added ' quotes around img src in v023
            stb.Append("</div>")
            stb.Append("<div style='position: absolute; left: 360px; top: -190px'>")
            stb.Append("<h1>" & ZoneName & "</h1>")
            stb.Append("</div>")

            stb.Append("<div id='ArtImageDiv' style='position: absolute; left: 55px; top: -130px;'>")
            ArtImageBox.style = "height:180px;width:180px"
            ArtImageBox.imagePathNormal = MyArt
            ArtImageBox.page = MyPageName & MyClientID
            stb.Append(ArtImageBox.Build)
            stb.Append("</div>")

            stb.Append("<div id='VolumeDiv' style='position: absolute; left: 25px; top: -130px'>")
            MyVolume = MusicAPI.PlayerVolume
            Dim VolumeSlider As New clsJQuery.jqSlider("VolumeSlider", 0, 100, MyVolume, clsJQuery.jqSlider.jqSliderOrientation.vertical, 180, MyPageName & MyClientID, True)
            VolumeSlider.toolTip = "Shows volume. Drag to set player volume"
            VolumeSlider.page = MyPageName & MyClientID
            stb.Append(VolumeSlider.build)
            stb.Append("</div>")

            stb.Append("<div id='MuteDiv' style='position: absolute; left: 15px; top: 55px'>")
            ButMute.style = "height:35px;width:35px"
            ButMute.imagePathNormal = ButMuteImagePath
            ButMute.toolTip = "Toggle between mute and unmute"
            ButMute.page = MyPageName & MyClientID
            stb.Append(ButMute.Build)
            stb.Append("</div>")

            Dim PlayerPosition As Integer = MusicAPI.SonosPlayerPosition
            Dim AlreadyPlayedSpan As TimeSpan = TimeSpan.FromSeconds(PlayerPosition)
            Dim AlreadyPlayedString As String = FormatMyTimeString(AlreadyPlayedSpan, LblDuration)
            Dim ToPlaySpan As TimeSpan = TimeSpan.FromSeconds(LblDuration - PlayerPosition)
            Dim ToPlayString As String = FormatMyTimeString(ToPlaySpan, LblDuration)
            If LblDuration = 0 Then
                ToPlayString = "0:00"
            End If
            stb.Append("<div id='PositionDiv' style='position: absolute; left: 25px; top: -165px; float: center'>")
            MyPosition = MusicAPI.SonosPlayerPosition
            Dim PositionSlider As New clsJQuery.jqSlider("PositionSlider", 0, LblDuration, MyPosition, clsJQuery.jqSlider.jqSliderOrientation.horizontal, 155, MyPageName & MyClientID, True)
            PositionSlider.toolTip = "Shows track position. Drag to change position"
            PositionSlider.page = MyPageName & MyClientID
            stb.Append(AlreadyPlayedString & "&nbsp;&nbsp;" & PositionSlider.build & "&nbsp;&nbsp;-" & ToPlayString)
            stb.Append("</div>")

            stb.Append("<div id='RepeatDiv' style='position: absolute; left: 235px; top: -90px'>")
            ButRepeat.imagePathNormal = ButRepeatImagePath
            ButRepeat.style = "height:40px;width:40px"
            ButRepeat.toolTip = "Toggle between Repeat and no-Repeat"
            ButRepeat.page = MyPageName & MyClientID
            stb.Append(ButRepeat.Build)
            stb.Append("</div>")

            stb.Append("<div id='ShuffleDiv' style='position: absolute; left: 235px; top: -40px'>")
            ButShuffle.imagePathNormal = ButShuffleImagePath
            ButShuffle.style = "height:40px;width:40px"
            ButShuffle.toolTip = "Toggle between Shuffled and Ordered (no-Shuffle)"
            ButShuffle.page = MyPageName & MyClientID
            stb.Append(ButShuffle.Build)
            stb.Append("</div>")

            stb.Append("<div id='LinkStateDiv' style='position: absolute; left: 290px; top: -115px'>")
            If LblLinkState <> "" Then stb.Append("Link State:  " & LblLinkState & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='TrackDiv' style='position: absolute; left: 290px; top: -100px'>")
            If LblTrackName <> "" Then stb.Append("Title:  " & LblTrackName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='ArtistDiv'  style='position: absolute; left: 290px; top: -85px'>")
            If LblArtistName <> "" Then stb.Append("Artist: " & LblArtistName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='AlbumDiv'  style='position: absolute; left: 290px; top: -70px'>")
            If LblAlbumName <> "" Then stb.Append("Album:  " & LblAlbumName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='RadioStationDiv'  style='position: absolute; left: 290px; top: -55px'>")
            If LblRadiostationName <> "" Then stb.Append("RadioStation:  " & LblRadiostationName & "<br>")
            If iPodPlayerName <> "" Then stb.Append("iPod PlayerName:  " & iPodPlayerName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='NextTrackDiv' style='position: absolute; left: 290px; top: -40px'>")
            If LblNextTrackName <> "" Then stb.Append("Next Title:  " & LblNextTrackName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='NextArtistDiv'  style='position: absolute; left: 290px; top: -25px'>")
            If LblNextArtistName <> "" Then stb.Append("Next Artist: " & LblNextArtistName & "<br>")
            stb.Append("</div>")
            stb.Append("<div id='NextAlbumDiv'  style='position: absolute; left: 290px; top: -10px'>")
            If LblNextAlbumName <> "" Then stb.Append("Next Album:  " & LblNextAlbumName & "<br>")
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 175px; top: 80px'>")
            ButPrev.imagePathNormal = ImagesPath & "player-prev.png"
            ButPrev.style = "height:65px;width:65px"
            ButPrev.toolTip = "Play Previous Track"
            ButPrev.page = MyPageName & MyClientID
            stb.Append(ButPrev.Build)
            stb.Append("</div>")

            stb.Append("<div id='PlayBtnDiv' style='position: absolute; left: 260px; top: 80px'>")
            MyPlayerState = MusicAPI.PlayerState()
            If MyPlayerState = Player_state_values.Playing Then
                ButPlay.imagePathNormal = ImagesPath & "player-pause.png"
            Else
                ButPlay.imagePathNormal = ImagesPath & "player-play.png"
            End If
            ButPlay.style = "height:95px;width:95px"
            ButPlay.toolTip = "Toggle between Play and Pause state"
            ButPlay.page = MyPageName & MyClientID
            stb.Append(ButPlay.Build)
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 370px; top: 110px'>")
            ButStop.imagePathNormal = ImagesPath & "player-stop.png"
            ButStop.style = "height:65px;width:65px"
            ButStop.toolTip = "Stop Player"
            ButStop.page = MyPageName & MyClientID
            stb.Append(ButStop.Build)
            stb.Append("</div>")

            stb.Append("<div style='position: absolute; left: 450px; top: 80px'>")
            ButNext.imagePathNormal = ImagesPath & "player-next.png"
            ButNext.style = "height:65px;width:65px"
            ButNext.toolTip = "Play Next Track"
            ButNext.page = MyPageName & MyClientID
            stb.Append(ButNext.Build)
            stb.Append("</div>")

            stb.Append(clsPageBuilder.DivStart("GroupPanel", "style='position: absolute; left: 650px; top: 75px; color:#0000FF; text-align: middle'"))
            BuildGroupingInfo(Control)
            GroupingOverlay.page = MyPageName & MyClientID
            stb.Append(GroupingOverlay.Build)
            stb.Append(clsPageBuilder.DivEnd) ' GroupPane

            stb.Append(clsPageBuilder.DivEnd) ' Position Relative


            stb.Append("</td></tr></table><br />")
            stb.Append(clsPageBuilder.DivEnd)   ' Player Panel


            stb.Append(clsPageBuilder.DivStart("ButtonPanel", "style='color:#0000FF; text-align: middle'"))
            stb.Append("<table><tr><td style='text-align: center'>")
            ButGenres.toolTip = "Select Genres in Navigationbox"
            ButGenres.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Genres...</h2>' });"
            ButGenres.page = MyPageName & MyClientID
            ButArtists.toolTip = "Select Artists in Navigationbox"
            ButArtists.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Artists...</h2>' });"
            ButArtists.page = MyPageName & MyClientID
            ButAlbums.toolTip = "Select Albums in Navigationbox"
            ButAlbums.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Albums...</h2>' });"
            ButAlbums.page = MyPageName & MyClientID
            ButTracks.toolTip = "Select Tracks in Navigationbox"
            ButTracks.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Tracks...</h2>' });"
            ButTracks.page = MyPageName & MyClientID
            ButPlayLists.toolTip = "Select Playlists in Navigationbox"
            ButPlayLists.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Playlists...</h2>' });"
            ButPlayLists.page = MyPageName & MyClientID
            ButRadioStations.toolTip = "Select Radiostations in Navigationbox"
            ButRadioStations.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Radiostations...</h2>' });"
            ButRadioStations.page = MyPageName & MyClientID
            ButAudioBooks.toolTip = "Select Audiobooks in Navigationbox"
            ButAudioBooks.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Audiobooks...</h2>' });"
            ButAudioBooks.page = MyPageName & MyClientID
            ButPodCasts.toolTip = "Select Podcasts in Navigationbox"
            ButPodCasts.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Podcasts...</h2>' });"
            ButPodCasts.page = MyPageName & MyClientID
            ButFavorites.toolTip = "Select Favorites in Navigationbox"
            ButFavorites.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Favorites...</h2>' });"
            ButFavorites.page = MyPageName & MyClientID
            ButLineInput.toolTip = "Select which Audio Line Input to connect to"
            ButLineInput.page = MyPageName & MyClientID
            ButPair.toolTip = "Select which player to pair with"
            ButPair.page = MyPageName & MyClientID
            ButUnpair.toolTip = "Unpair this player"
            ButUnpair.page = MyPageName & MyClientID
            ButTV.toolTip = "Select TV as Input"
            ButTV.page = MyPageName & MyClientID
            stb.Append(clsPageBuilder.DivStart("ButtonDiv", "style='color:#0000FF; text-align: middle'"))
            If ZoneModel = "WD100" Then
                stb.Append(ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButAudioBooks.Build & ButPodCasts.Build)
            Else
                stb.Append(ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButRadioStations.Build & ButLineInput.Build & ButTV.Build & ButPair.Build & ButUnpair.Build)
            End If
            stb.Append(clsPageBuilder.DivEnd) ' ButtonDiv
            stb.Append("</td></tr></table>")
            stb.Append(clsPageBuilder.DivEnd) ' ButtonPanel


            stb.Append(clsPageBuilder.DivStart("LabelDiv", "style='color:#0000FF; text-align: middle'"))
            'stb.Append("<table><tr><td style='text-align: center'>")
            LblAlbum.page = MyPageName & MyClientID
            LblArtist.page = MyPageName & MyClientID
            LblAudiobooks.page = MyPageName & MyClientID
            LblFavorites.page = MyPageName & MyClientID
            LblGenre.page = MyPageName & MyClientID
            LblLineInput.page = MyPageName & MyClientID
            LblPairing.page = MyPageName & MyClientID
            LblPlaylist.page = MyPageName & MyClientID
            LblPodcasts.page = MyPageName & MyClientID
            LblRadioList.page = MyPageName & MyClientID
            stb.Append(LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
            'stb.Append("</td></tr></table>")
            stb.Append(clsPageBuilder.DivEnd) ' LabelDiv

            stb.Append(clsPageBuilder.DivStart("NavigationPanel", "style='color:#0000FF; text-align: middle'"))
            stb.Append("<table ><tr><td nowrap='nowrap'>")
            'NavigationBox.style = "height:284px;width:370px;overflow:auto"
            NavigationBox.style = "height:284px;width:370px"
            NavigationBox.height = 284
            NavigationBox.width = 370
            NavigationBox.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
            'NavigationBox.UseBothClickEvents = True
            'NavigationBox.WaitTime = 200
            'PlaylistBox.style = "height:284px;width:370px;overflow:auto"
            PlaylistBox.style = "height:284px;width:370px"
            PlaylistBox.height = 284
            PlaylistBox.width = 370
            'PlaylistBox.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
            stb.Append(clsPageBuilder.DivStart("NavigationDiv", ""))
            NavigationBox.toolTip = "Navigationbox: Click to navigate into an item. Once at the bottom, use the 'Add to Playlist' button"
            NavigationBox.page = MyPageName & MyClientID
            stb.Append(NavigationBox.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td><td nowrap='nowrap'>")
            stb.Append(clsPageBuilder.DivStart("PlaylistDiv", ""))
            PlaylistBox.toolTip = "Playlistbox: Click to play item"
            PlaylistBox.page = MyPageName & MyClientID
            stb.Append(PlaylistBox.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td></tr><tr><td  align='left'>")
            stb.Append(clsPageBuilder.DivStart("ButAddToPlaylistDiv", ""))
            ButAddToPlaylist.toolTip = "The item selected in the Navigationbox will be added to the playlist when clicked"
            ButAddToPlaylist.page = MyPageName & MyClientID
            stb.Append(ButAddToPlaylist.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td><td align='left'>")
            stb.Append(clsPageBuilder.DivStart("ButClearListDiv", ""))
            ButClearList.toolTip = "Clears the entire Playlist when clicked"
            ButClearList.page = MyPageName & MyClientID
            stb.Append(ButClearList.Build)
            stb.Append(clsPageBuilder.DivEnd)
            stb.Append("</td></tr></table>")
            stb.Append(clsPageBuilder.DivEnd) ' NavigationPanel

            stb.Append(clsPageBuilder.DivEnd) ' page end
            stb.Append(clsPageBuilder.FormEnd)

            Me.AddBody(stb.ToString)

            If GenerateHeaderFooter Then
                ' add the body html to the page
                Me.AddFooter(hs.GetPageFooter)
                Me.suppressDefaultFooter = True
                ' return the full page
            End If

            Return Me.BuildPage()

        Catch ex As Exception
            Log("Error in GetPagePlugin for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return stb.ToString

    End Function

    Private Function GenerateNavPane() As String
        Dim stb As New StringBuilder
        'stb.Append("<iframe scrolling = ""no"" style=""border: 0px none; height: 1000px; margin-top: -52px; width: 1000px;"">")

        stb.Append(clsPageBuilder.DivStart("NavigationPanel", "style='color:#0000FF; text-align: middle'"))
        stb.Append("<table ><tr><td nowrap='nowrap'>")
        NavigationBox.style = "height:inherit;width:inherit;overflow:auto"
        'NavigationBox.style = "height:284px;width:370px"
        NavigationBox.height = 650
        NavigationBox.width = 600
        'NavigationBox.style = ""
        'NavigationBox.height = 0
        'NavigationBox.width = 0
        'NavigationBox.style = "overflow: auto; max-width: 1000px;"
        NavigationBox.functionToCallOnClick = "" '"$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading ...</h2>' });"
        'NavigationBox.UseBothClickEvents = True
        'NavigationBox.WaitTime = 200
        stb.Append(clsPageBuilder.DivStart("NavigationDiv", ""))
        NavigationBox.toolTip = "Navigationbox: Click to navigate into an item. Once at the bottom, use the 'Add to Playlist' button"
        NavigationBox.page = MyPageName & MyClientID & "&control=navpane"
        stb.Append(NavigationBox.Build)
        stb.Append(clsPageBuilder.DivEnd)
        stb.Append("</td></tr></table>")
        stb.Append(clsPageBuilder.DivEnd) ' NavigationPanel
        'stb.Append("</iframe>")

        Return stb.ToString

    End Function

    Private Function GenerateNavControl() As String
        Dim stb As New StringBuilder
        'stb.Append(clsPageBuilder.DivStart("ButtonPanel", "style='color:#0000FF; text-align: middle'"))
        'stb.Append("<table><tr><td style='text-align: center'>")
        ButGenres.toolTip = "Select Genres in Navigationbox"
        ButGenres.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Genres...</h2>' });"
        ButGenres.page = MyPageName & MyClientID & "&control=navcontrol"
        ButArtists.toolTip = "Select Artists in Navigationbox"
        ButArtists.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Artists...</h2>' });"
        ButArtists.page = MyPageName & MyClientID & "&control=navcontrol"
        ButAlbums.toolTip = "Select Albums in Navigationbox"
        ButAlbums.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Albums...</h2>' });"
        ButAlbums.page = MyPageName & MyClientID & "&control=navcontrol"
        ButTracks.toolTip = "Select Tracks in Navigationbox"
        ButTracks.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Tracks...</h2>' });"
        ButTracks.page = MyPageName & MyClientID & "&control=navcontrol"
        ButPlayLists.toolTip = "Select Playlists in Navigationbox"
        ButPlayLists.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Playlists...</h2>' });"
        ButPlayLists.page = MyPageName & MyClientID & "&control=navcontrol"
        ButRadioStations.toolTip = "Select Radiostations in Navigationbox"
        ButRadioStations.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Radiostations...</h2>' });"
        ButRadioStations.page = MyPageName & MyClientID & "&control=navcontrol"
        ButAudioBooks.toolTip = "Select Audiobooks in Navigationbox"
        ButAudioBooks.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Audiobooks...</h2>' });"
        ButAudioBooks.page = MyPageName & MyClientID & "&control=navcontrol"
        ButPodCasts.toolTip = "Select Podcasts in Navigationbox"
        ButPodCasts.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Podcasts...</h2>' });"
        ButPodCasts.page = MyPageName & MyClientID & "&control=navcontrol"
        ButFavorites.toolTip = "Select Favorites in Navigationbox"
        ButFavorites.functionToCallOnClick = "$.blockUI({ message: '<h2><img src=""/images/HomeSeer/ui/spinner.gif"" /> Loading Favorites...</h2>' });"
        ButFavorites.page = MyPageName & MyClientID & "&control=navcontrol"
        ButLineInput.toolTip = "Select which Audio Line Input to connect to"
        ButLineInput.page = MyPageName & MyClientID & "&control=navcontrol"
        ButPair.toolTip = "Select which player to pair with"
        ButPair.page = MyPageName & MyClientID & "&control=navcontrol"
        ButUnpair.toolTip = "Unpair this player"
        ButUnpair.page = MyPageName & MyClientID & "&control=navcontrol"
        ButTV.toolTip = "Select TV as Input"
        ButTV.page = MyPageName & MyClientID & "&control=navcontrol"
        ButClearList.toolTip = "Clears the entire Playlist when clicked"
        ButClearList.page = MyPageName & MyClientID & "&control=navcontrol"
        'stb.Append(ButClearList.Build)
        stb.Append(clsPageBuilder.DivStart("ButtonDiv", "style='color:#0000FF; text-align: middle'"))
        If ZoneModel = "WD100" Then
            stb.Append(ButClearList.Build & ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButAudioBooks.Build & ButPodCasts.Build)
        Else
            stb.Append(ButClearList.Build & ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButRadioStations.Build & ButLineInput.Build & ButTV.Build & ButPair.Build & ButUnpair.Build)
        End If
        stb.Append(clsPageBuilder.DivEnd) ' ButtonDiv
        'stb.Append("</td></tr></table>")
        'stb.Append(clsPageBuilder.DivEnd) ' ButtonPanel
        Return stb.ToString
    End Function

    Private Function GenerateNavTree() As String
        Dim stb As New StringBuilder
        stb.Append(clsPageBuilder.DivStart("LabelDiv", "style='color:#0000FF; text-align: middle'"))
        LblAlbum.page = MyPageName & MyClientID & "&control=navtree"
        LblArtist.page = MyPageName & MyClientID & "&control=navtree"
        LblAudiobooks.page = MyPageName & MyClientID & "&control=navtree"
        LblFavorites.page = MyPageName & MyClientID & "&control=navtree"
        LblGenre.page = MyPageName & MyClientID & "&control=navtree"
        LblLineInput.page = MyPageName & MyClientID & "&control=navtree"
        LblPairing.page = MyPageName & MyClientID & "&control=navtree"
        LblPlaylist.page = MyPageName & MyClientID & "&control=navtree"
        LblPodcasts.page = MyPageName & MyClientID & "&control=navtree"
        LblRadioList.page = MyPageName & MyClientID & "&control=navtree"
        stb.Append(LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
        stb.Append(clsPageBuilder.DivEnd) ' LabelDiv
        Return stb.ToString
    End Function

    Private Function GenerateQueue() As String
        Dim stb As New StringBuilder
        'PlaylistBox.style = "height:284px;width:370px"
        'PlaylistBox.height = 284
        'PlaylistBox.width = 370
        PlaylistBox.style = "overflow: auto; max-width: 1000px;"
        PlaylistBox.height = 0
        PlaylistBox.width = 0
        stb.Append(clsPageBuilder.DivStart("PlaylistDiv", ""))
        PlaylistBox.toolTip = "Playlistbox: Click to play item"
        PlaylistBox.page = MyPageName & MyClientID & "&control=queue"
        stb.Append(PlaylistBox.Build)
        stb.Append(clsPageBuilder.DivEnd)
        Return stb.ToString
    End Function

    Private Sub LoadNavigationBox(Optional GenerateDiv As Boolean = False)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadNavigationBox called for Player = " & ZoneName & " and GenerateDiv= " & GenerateDiv.ToString, LogType.LOG_TYPE_INFO)
        Try
            Dim i As Integer
            NavigationBox.items.Clear()
            If ar IsNot Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadNavigationBox called for Player = " & ZoneName & " and has " & (UBound(ar) + 1).ToString & " entries in the NavigationBox", LogType.LOG_TYPE_INFO)
                NavigationBox.selectedItemIndex = -1
                For i = 0 To UBound(ar)
                    NavigationBox.AddItem(ar(i).ToString, HttpUtility.UrlEncode(ar(i).ToString), False)
                    'Log("LoadListBox added '" & ar(i).ToString & "' to Navigationbox", LogType.LOG_TYPE_INFO)
                Next
            End If
            If GenerateDiv Then Me.divToUpdate.Add("NavigationDiv", NavigationBox.Build)
            If GenerateDiv Then Me.divToUpdate.Add("LabelDiv", LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
        Catch ex As Exception
            Log("Error in LoadNavigationBox for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub LoadPlayListBox(Optional GenerateDiv As Boolean = False)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadPlayListBox called for Player = " & ZoneName & " and GenerateDiv= " & GenerateDiv.ToString, LogType.LOG_TYPE_INFO)
        PlaylistBox.items.Clear()
        Try
            Dim i As Integer
            If arplaylist IsNot Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadPlayListBox called for Player = " & ZoneName & " and has " & (UBound(arplaylist) + 1).ToString & " entries in the PlaylistBox", LogType.LOG_TYPE_INFO)
                PlaylistBox.items.Clear()
                For i = 0 To UBound(arplaylist)
                    PlaylistBox.AddItem(arplaylist(i), HttpUtility.UrlEncode(arplaylist(i)), arplaylist(i) = MusicAPI.CurrentTrack)
                Next
                MusicAPI.MyQueueHasChanged = False
            End If
        Catch ex As Exception
            Log("Error in LoadPlayListBox for PlayerControl for Player = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub UpdateStatus()
        Try
            Dim st As String
            st = MusicAPI.CurrentTrack
            MyTrack = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblTrackName = st

            st = MusicAPI.CurrentArtist
            MyArtist = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblArtistName = st

            st = MusicAPI.CurrentAlbum
            MyAlbum = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblAlbumName = st

            st = MusicAPI.RadiostationName
            MyRadiostation = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblRadiostationName = st

            st = MusicAPI.NextTrack
            MyNextTrack = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblNextTrackName = st

            st = MusicAPI.NextArtist
            MyNextArtist = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblNextArtistName = st

            st = MusicAPI.NextAlbum
            MyNextAlbum = st
            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblNextAlbumName = st

            LblDuration = Val(MusicAPI.CurrentTrackDuration)

            st = MusicAPI.CurrentTrackDescription

            If st.Length > 86 Then st = st.Substring(0, 84) & "..."
            LblDescr = st

            MyArt = MusicAPI.CurrentAlbumArtPath.ToString
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateStatus for PlayerControl for player = " & MusicAPI.ZonePlayerName & " has shuffle state = " & MusicAPI.ShuffleStatus.ToLower, LogType.LOG_TYPE_INFO)
            MyShuffle = False
            Select Case MusicAPI.ShuffleStatus.ToLower
                Case "shuffled"
                    ButShuffleImagePath = ImagesPath & "Shuffle.png"
                    MyShuffle = True
                Case "ordered"
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                Case "sorted"
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                Case Else
                    ButShuffleImagePath = ImagesPath & "NoShuffle.png"
            End Select
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpdateStatus for PlayerControl for player = " & MusicAPI.ZonePlayerName & " has repeat state = " & MusicAPI.SonosRepeat.ToString, LogType.LOG_TYPE_INFO)
            MyRepeat = False
            Select Case MusicAPI.SonosRepeat
                Case Repeat_modes.repeat_all
                    ButRepeatImagePath = ImagesPath & "Repeat.png"
                    MyRepeat = True
                Case Repeat_modes.repeat_off
                    ButRepeatImagePath = ImagesPath & "NoRepeat.png"
                Case Repeat_modes.repeat_one
                    ButRepeatImagePath = ImagesPath & "Repeat.png"
                    MyRepeat = True
            End Select
            If MusicAPI.GetMuteState("Master") = True Then
                MyMute = True
                ButMuteImagePath = ImagesPath & "Muted.png"
            Else
                MyMute = False
                ButMuteImagePath = ImagesPath & "UnMuted.png"
            End If
            LblLinkState = ""
            Dim SlaveUDN As String = ""
            If MusicAPI.ZoneIsPairMaster Then
                LblLinkState = "Paired with " & PIReference.GetZoneNamebyUDN(MusicAPI.GetZonePairSlaveUDN)
                SlaveUDN = MusicAPI.GetZonePairSlaveUDN
            ElseIf MusicAPI.ZoneIsPairSlave Then
                LblLinkState = "Paired to " & PIReference.GetZoneNamebyUDN(MusicAPI.GetZonePairMasterUDN)
            ElseIf MusicAPI.PlayBarMaster And MusicAPI.ZoneSource = "TV" Then
                LblLinkState = "Connected to TV"
            ElseIf MusicAPI.PlayBarSlave Then
                LblLinkState = "Connected to Playbar"
            End If
            If MusicAPI.ZoneIsLinked And Not MusicAPI.ZoneIsPairSlave Then
                If LblLinkState <> "" Then LblLinkState &= " "
                LblLinkState &= "Linked to " & PIReference.GetZoneNamebyUDN(MusicAPI.LinkedZoneSource)
            End If
            Dim TargetZones() = MusicAPI.GetZoneDestination()
            Dim LinkZoneList As String = ""
            If TargetZones IsNot Nothing Then
                For Each Target As String In TargetZones
                    If Target <> "" And Target <> SlaveUDN Then
                        If LinkZoneList = "" Then LinkZoneList = "Linked with "
                        LinkZoneList &= PIReference.GetZoneNamebyUDN(Target) & " "
                    End If
                Next
            End If
            If LinkZoneList <> "" Then
                If LblLinkState <> "" Then LblLinkState &= " "
                LblLinkState &= LinkZoneList
            End If
        Catch ex As Exception
            Log("Error in UpdateStatus for PlayerControl with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Enum ITEM_TYPE
        genres = 1
        artists = 2
        albums = 3
        tracks = 4
        playlists = 5
        radioLists = 6  ' added for SonosController
        Audiobooks = 7  ' added for SonosController
        Podcasts = 8    ' added for SonosController
        LineInput = 9   ' added for SonosController
        Pairing = 10    ' added for SonosController
        Favorites = 11
    End Enum

    Private Sub ClearSelections()
        LblGenre.label = ""
        LblArtist.label = ""
        LblAlbum.label = ""
        LblPlaylist.label = ""
        LblRadioList.label = ""
        LblAudiobooks.label = ""
        LblPodcasts.label = ""
        LblFavorites.label = ""
        LblLineInput.label = ""
        LblPairing.label = ""
        LblWhatWasPreviouslyLoaded = ""
    End Sub

    Private Sub BuildGroupingInfo(Control As String)
        GroupingOverlay.toolTip = "This will open the selection box to group players"
        GroupingOverlay.label = "Group Players"
        GroupingOverlay.overlayHTML = clsPageBuilder.FormStart("Groupform", MyPageName & MyClientID, "post")
        GroupingOverlay.overlayHTML &= "<div>Add/Remove players in group<br><br>"
        Dim tbut1 As New clsJQuery.jqButton("ovSubmit", "Submit", MyPageName & MyClientID & "&control=" & Control, True)
        GroupingOverlay.overlayHTML &= tbut1.Build
        Dim tbut2 As New clsJQuery.jqButton("ovCancel", "Cancel", MyPageName & MyClientID & "&control=" & Control, True)
        GroupingOverlay.overlayHTML &= tbut2.Build & "<br /><br />"

        Dim PlayerUDNArray() As String
        Dim PlayerCheckBox As clsJQuery.jqCheckBox
        Try
            PlayerUDNArray = PIReference.GetAllActiveZones()
            'Dim PlayerDestList As String = MusicAPI.GetTargetZoneLinkedList
            If PlayerUDNArray IsNot Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("LoadButGroup called for Player = " & ZoneName & " and has " & (UBound(PlayerUDNArray) + 1).ToString & " entries", LogType.LOG_TYPE_INFO)
                For i = 0 To UBound(PlayerUDNArray)
                    Dim Player As HSPI = PIReference.GetAPIByUDN(PlayerUDNArray(i))
                    If Player.ZoneModel.ToUpper <> "WD100" And Player.ZoneModel.ToUpper <> "SUB" And Not Player.ZoneIsASlave Then
                        PlayerCheckBox = New clsJQuery.jqCheckBox("Label-" & Player.GetUDN, Player.GetZoneName, MyPageName & MyClientID & "&control=" & Control, False, False)
                        If (Player.ZoneIsLinked And Player.LinkedZoneSource = MyZoneUDN) Then
                            ' this is a player linked to this zone
                            PlayerCheckBox.checked = True
                        ElseIf Player.GetUDN = MyZoneUDN Then
                            PlayerCheckBox.checked = True
                        ElseIf Player.ZoneIsLinked Then
                            If PIReference.GetAPIByUDN(Player.LinkedZoneSource).GetTargetZoneLinkedList.IndexOf(MyZoneUDN) <> -1 Then
                                PlayerCheckBox.checked = True
                            Else
                                PlayerCheckBox.checked = False
                            End If
                        ElseIf Player.GetTargetZoneLinkedList.IndexOf(MyZoneUDN) <> -1 Then
                            PlayerCheckBox.checked = True ' This is the master of my zone
                        Else
                            PlayerCheckBox.checked = False
                        End If
                        GroupingOverlay.overlayHTML &= PlayerCheckBox.Build & "<br />"
                    End If
                Next
            End If
            MusicAPI.MyZoneGroupStateHasChanged = False
        Catch ex As Exception
            Log("Error in GetPagePlugin for PlayerControl for Player = " & ZoneName & " trying to check linked zones with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        GroupingOverlay.overlayHTML &= "</div>"
        GroupingOverlay.overlayHTML &= clsPageBuilder.FormEnd ' Groupform

    End Sub


    Public Function PlayerDevicePost(ZoneUDN As String, data As String, user As String, userRights As Integer) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost called for Player = " & ZoneName & " with ZoneUDN = " & ZoneUDN.ToString & " and data = " & data.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
        'PlayerDevicePost = Enums.ConfigDevicePostReturn.CallbackOnce
        PlayerDevicePost = ""
        'Me.reset()
        ' handle form items:
        Dim parts As Collections.Specialized.NameValueCollection
        parts = System.Web.HttpUtility.ParseQueryString(data)
        ' handle items like:
        ' if parts("id")="mybutton" then
        Dim ClientID As String = ""
        Dim Control As String = ""

        If parts IsNot Nothing Then
            Dim DoubleClickFlag As Boolean = False
            Try
                If parts.Item("click") = "double" Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost for PlayerControl for Player = " & ZoneName & " found double click key", LogType.LOG_TYPE_WARNING)
                    DoubleClickFlag = True
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayerDevicePost for PlayerControl for Player = " & ZoneName & " searching for Key = click with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                ClientID = parts.Item("clientid")
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayerDevicePost for PlayerControl for Player = " & ZoneName & " searching for Key = clientid with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                Control = parts.Item("control")
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in PlayerDevicePost for PlayerControl for Player = " & ZoneName & " searching for Key = control with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If ClientID <> "" Then MyClientID = "?clientid=" & ClientID Else MyClientID = ""
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(Part, "_")

                        Dim ObjectValue As String = System.Web.HttpUtility.HtmlDecode(parts(Part))
                        Select Case ObjectNameParts(0).ToUpper.ToString
                            Case "INSTANCE"
                            Case "CLIENTID"
                            Case "CONTROL"
                            Case "REF"
                            Case "PLUGIN"
                                If ObjectValue <> sIFACE_NAME Then
                                    Log("Error in PlayerDevicePost called with ZoneUDN = " & ZoneUDN.ToString & " and user = " & user & " and userRights = " & userRights.ToString & ". Wrong Plugin Name", LogType.LOG_TYPE_ERROR)
                                    Exit For
                                End If
                            Case "ACTION"
                            Case "ID"
                            Case "BUTPLAY"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TogglePlay()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Toggle Play command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTPREV"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackPrev()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Previous command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTSTOP"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.StopPlay()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Stop command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTPAUSE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.SonosPause()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Pause command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTNEXT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackNext()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Next command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTMUTE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleMuteState("Master")
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Mute command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTREPEAT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleRepeat()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Repeat command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "BUTSHUFFLE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleShuffle()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Shuffle command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                    'Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
                                End If
                            Case "VOLUMESLIDER"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Volume command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.PlayerVolume = Val(ObjectValue)
                            Case "POSITIONSLIDER"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Position command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.SonosPlayerPosition = Val(ObjectValue)
                            Case "BUTLOUDNESS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Loudness command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.ToggleLoudnessState("Master")
                            Case "BUTGENRES"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Genres command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButGenres_Click()
                            Case "BUTARTISTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Artists command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButArtists_Click()
                            Case "BUTALBUMS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Albums command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButAlbums_Click()
                            Case "BUTTRACKS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Tracks command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButTracks_Click()
                            Case "BUTPLAYLISTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Playlists command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPlayLists_Click()
                            Case "BUTRADIOSTATIONS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued RadioStations command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButRadioStations_Click()
                            Case "BUTAUDIOBOOKS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued AudioBooks command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButAudiobooks_Click()
                            Case "BUTPODCASTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Podcasts command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPodcasts_Click()
                            Case "BUTFAVORITES"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Favorites command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButFavorites_Click()
                            Case "BUTLINEINPUT"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued LineInput command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButLineInput_Click()
                            Case "BUTPAIR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Pair command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPair_Click()
                            Case "BUTUNPAIR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued UnPair command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButUnPair_Click()
                            Case "LBLGENRE"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued LblGenre command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                LblGenre_Click()
                            Case "LBLARTIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued LblArtist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                LblArtist_Click()
                            Case "BUTADDTOPLAYLIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued AddToPlaylist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                AddTrack_Click()
                            Case "BUTCLEARLIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued ClearList command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButClearList_Click()
                            Case "BUTEDITOR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Editor command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case "PLAYLISTBOX"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued Playlistbox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                PlaylistBox_SelectedIndexChanged(ObjectValue.ToString)
                                'Return Enums.ConfigDevicePostReturn.CallbackOnce
                            Case "NAVIGATIONBOX"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued NavigationBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                NavigationBox_SelectedIndexChanged(ObjectValue.ToString, DoubleClickFlag, ClientID, Control)
                                'Return Enums.ConfigDevicePostReturn.CallbackOnce
                            Case "BUTTV"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued TV command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButTV_Click()
                            Case "CLICK"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost issued click command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case Else
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_WARNING)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_WARNING)
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in PlayerDevicePost processing for Player = " & ZoneName & " with ZoneUDN = " & ZoneUDN.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerDevicePost for Player = " & ZoneName & " found parts to be empty", LogType.LOG_TYPE_INFO)
        End If
        If LblWhatWasPreviouslyLoaded <> LblWhatsLoaded Then
            SaveNavSettingstoINIFile(ClientID)
            LblWhatWasPreviouslyLoaded = LblWhatsLoaded
        End If
        Return Enums.ConfigDevicePostReturn.CallbackOnce

    End Function


    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostBackProc for PlayerControl called  for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
        'Me.reset()
        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(System.Web.HttpUtility.HtmlDecode(data))
        Dim UpdateFlag As Boolean = False
        Dim ClientID As String = ""
        Dim Control As String = ""
        If parts IsNot Nothing Then
            'Log("PostBackProc for PlayerControl called  for Player = " & ZoneName & " with part = '" & parts.GetKey(0).ToUpper.ToString & "'", LogType.LOG_TYPE_INFO)
            Try
                If parts.Item("action").ToUpper = "UPDATETIME" Then
                    UpdateFlag = True
                Else
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PostBackProc for PlayerControl called for Player = " & ZoneName & " with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
            End Try
            Dim DoubleClickFlag As Boolean = False
            Try
                If parts.Item("click") = "double" Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found double click key", LogType.LOG_TYPE_WARNING)
                    DoubleClickFlag = True
                End If
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in postBackProc for PlayerControl for Player = " & ZoneName & " searching for key = click with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                ClientID = parts.Item("clientid")
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in postBackProc for PlayerControl for Player = " & ZoneName & " searching for Key = clientid with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Try
                Control = parts.Item("control")
            Catch ex As Exception
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in postBackProc for PlayerControl for Player = " & ZoneName & " searching for Key = control with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            If ClientID <> "" Then MyClientID = "?clientid=" & ClientID Else MyClientID = ""
            If Control <> "" Then
                NavigationBox.page = MyPageName & MyClientID & "&control=" & Control
                PlaylistBox.page = MyPageName & MyClientID & "&control=" & Control
            End If
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(HttpUtility.UrlDecode(Part), "_")
                        If Not UpdateFlag Then
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_INFO)
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PlayerControl for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_INFO)
                        End If
                        Dim ObjectValue As String = HttpUtility.UrlDecode(parts(Part))
                        'Dim ObjectValue As String = parts(Part)

                        Select Case ObjectNameParts(0).ToString.ToUpper
                            Case "INSTANCE"
                            Case "CLIENTID"
                            Case "CONTROL"
                            Case "REF"
                            Case "PLUGIN"
                                If ObjectValue <> sIFACE_NAME Then
                                    Log("Error in postBackProc for Zoneplayer = " & ZoneName & " and page = " & page.ToString & ", data = " & data & ", user = " & user & ", userRights = " & userRights.ToString, LogType.LOG_TYPE_ERROR)
                                    Exit For
                                End If
                            Case "ACTION"
                                If ObjectValue = "updatetime" Then
                                    CheckForChanges(Control, ClientID)
                                End If
                            Case "ID"
                            Case "BUTPLAY"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TogglePlay()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Toggle Play command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTPREV"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackPrev()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Previous command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTSTOP"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.StopPlay()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Stop command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTPAUSE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.SonosPause()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Pause command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTNEXT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.TrackNext()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Next command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTMUTE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleMuteState("Master")
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Mute command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTREPEAT"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleRepeat()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Repeat command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "BUTSHUFFLE"
                                If ObjectValue.ToUpper = "SUBMIT" Then
                                    MusicAPI.ToggleShuffle()
                                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Shuffle command for Zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                                End If
                            Case "VOLUMESLIDER"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Volume command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.PlayerVolume = Val(ObjectValue)
                            Case "POSITIONSLIDER"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Position command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.SonosPlayerPosition = Val(ObjectValue)
                            Case "BUTLOUDNESS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Loudness command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                MusicAPI.ToggleLoudnessState("Master")
                            Case "BUTGENRES"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Genres command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButGenres_Click()
                            Case "BUTARTISTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Artists command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButArtists_Click()
                            Case "BUTALBUMS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Albums command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButAlbums_Click()
                            Case "BUTTRACKS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Tracks command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButTracks_Click()
                            Case "BUTPLAYLISTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Playlists command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPlayLists_Click()
                            Case "BUTRADIOSTATIONS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued RadioStations command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButRadioStations_Click()
                            Case "BUTAUDIOBOOKS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued AudioBooks command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButAudiobooks_Click()
                            Case "BUTPODCASTS"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Podcasts command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPodcasts_Click()
                            Case "BUTFAVORITES"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Favorites command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButFavorites_Click()
                            Case "BUTLINEINPUT"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued LineInput command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButLineInput_Click()
                            Case "BUTPAIR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Pair command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButPair_Click()
                            Case "BUTUNPAIR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued UnPair command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButUnPair_Click()
                            Case "LBLGENRE"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued LblGenre command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                LblGenre_Click()
                            Case "LBLARTIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued LblArtist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                LblArtist_Click()
                            Case "BUTADDTOPLAYLIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued AddToPlaylist command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                AddTrack_Click()
                            Case "BUTCLEARLIST"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued ClearList command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButClearList_Click()
                            Case "BUTEDITOR"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Editor command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case "PLAYLISTBOX"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued Playlistbox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                PlaylistBox_SelectedIndexChanged(ObjectValue.ToString)
                            Case "NAVIGATIONBOX"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued NavigationBox command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                NavigationBox_SelectedIndexChanged(ObjectValue.ToString, DoubleClickFlag, ClientID, Control)
                            Case "OVSUBMIT"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued ovSubmit command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ProcessGroupingChange(data)
                                If ObjectValue.ToString.ToUpper = "SUBMIT" Then
                                    GroupingOverlay.visible = False
                                    GroupingOverlay.page = MyPageName & MyClientID
                                    Me.divToUpdate.Add("GroupPanel", GroupingOverlay.Build)
                                End If
                                Exit For
                            Case "OVCANCEL"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued ovCancel command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                If ObjectValue.ToString.ToUpper = "SUBMIT" Then
                                    GroupingOverlay.visible = False
                                    GroupingOverlay.page = MyPageName & MyClientID
                                    Me.divToUpdate.Add("GroupPanel", GroupingOverlay.Build)
                                End If
                            Case "LABEL-RINCON"
                                ' ignore
                            Case "BUTTV"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued TV command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                                ButTV_Click()
                            Case "CLICK"
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc issued click command for Zoneplayer = " & ZoneName & " with Value = " & ObjectValue.ToString, LogType.LOG_TYPE_INFO)
                            Case Else
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for Player = " & ZoneName & " found Key = " & ObjectNameParts(0).ToString, LogType.LOG_TYPE_WARNING)
                                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for Player = " & ZoneName & " found Value = " & parts(Part).ToString, LogType.LOG_TYPE_WARNING)
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in postBackProc for PlayerControl processing with page = " & page.ToString & " and data = " & data.ToString & " and user = " & user.ToString & " and userRights = " & userRights.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("postBackProc for PlayerControl found parts to be empty", LogType.LOG_TYPE_INFO)
        End If
        If LblWhatWasPreviouslyLoaded <> LblWhatsLoaded Then
            SaveNavSettingstoINIFile(ClientID)
            LblWhatWasPreviouslyLoaded = LblWhatsLoaded
        End If
        Return MyBase.postBackProc(page, data, user, userRights)
    End Function

    Private Sub CheckForChanges(Control As String, ClientID As String)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlayerControl.CheckForChanges for Zoneplayer = " & ZoneName & " called and PlayerState = " & MusicAPI.PlayerState.ToString, LogType.LOG_TYPE_INFO)
        Try
            If Control = "" Then
                Dim trackHasChanged As Boolean = False
                Try
                    If LblDuration <> Val(MusicAPI.CurrentTrackDuration) Then
                        LblDuration = Val(MusicAPI.CurrentTrackDuration)
                        trackHasChanged = True
                    End If
                Catch ex As Exception
                    LblDuration = 0
                End Try
                If MyPlayerState <> MusicAPI.PlayerState Then
                    MyPlayerState = MusicAPI.PlayerState
                    If MyPlayerState = Player_state_values.Playing Then
                        ButPlay.imagePathNormal = ImagesPath & "player-pause.png"
                    Else
                        ButPlay.imagePathNormal = ImagesPath & "player-play.png"
                    End If
                    ButPlay.style = "height:95px;width:95px"
                    ButPlay.page = MyPageName & MyClientID
                    Me.divToUpdate.Add("PlayBtnDiv", ButPlay.Build)
                End If
                If (MusicAPI.PlayerState = Player_state_values.Playing) Or trackHasChanged Then
                    Dim PlayerPosition As Integer = MusicAPI.SonosPlayerPosition
                    Dim AlreadyPlayedSpan As TimeSpan = TimeSpan.FromSeconds(PlayerPosition)
                    Dim AlreadyPlayedString As String = FormatMyTimeString(AlreadyPlayedSpan, LblDuration)
                    Dim ToPlaySpan As TimeSpan = TimeSpan.FromSeconds(LblDuration - PlayerPosition)
                    Dim ToPlayString As String = FormatMyTimeString(ToPlaySpan, LblDuration)
                    If LblDuration = 0 Then
                        ToPlayString = "0:00"
                    End If
                    Dim PositionSlider As New clsJQuery.jqSlider("PositionSlider", 0, LblDuration, PlayerPosition, clsJQuery.jqSlider.jqSliderOrientation.horizontal, 155, MyPageName & MyClientID, True)
                    PositionSlider.toolTip = "Shows track position. Drag to change position"
                    Me.divToUpdate.Add("PositionDiv", AlreadyPlayedString & "&nbsp;&nbsp;" & PositionSlider.build & "&nbsp;&nbsp;-" & ToPlayString)
                End If
                If MusicAPI.PlayerVolume <> MyVolume Then
                    MyVolume = MusicAPI.PlayerVolume
                    Dim VolumeSlider As New clsJQuery.jqSlider("VolumeSlider", 0, 100, MyVolume, clsJQuery.jqSlider.jqSliderOrientation.vertical, 180, MyPageName & MyClientID, True)
                    VolumeSlider.toolTip = "Shows volume. Drag to set player volume"
                    Me.divToUpdate.Add("VolumeDiv", VolumeSlider.build)
                End If
                If MusicAPI.CurrentTrack <> MyTrack Then
                    MyTrack = MusicAPI.CurrentTrack
                    Dim st As String
                    st = MyTrack
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblTrackName = st
                    If LblTrackName <> "" Then Me.divToUpdate.Add("TrackDiv", "Title:  " & LblTrackName) Else Me.divToUpdate.Add("TrackDiv", "")
                End If
                If MusicAPI.CurrentAlbum <> MyAlbum Then
                    MyAlbum = MusicAPI.CurrentAlbum
                    Dim st As String
                    st = MyAlbum
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblAlbumName = st
                    If LblAlbumName <> "" Then Me.divToUpdate.Add("AlbumDiv", "Album:  " & LblAlbumName) Else Me.divToUpdate.Add("AlbumDiv", "")
                End If
                If MusicAPI.CurrentArtist <> MyArtist Then
                    MyArtist = MusicAPI.CurrentArtist
                    Dim st As String
                    st = MyArtist
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblArtistName = st
                    If LblArtistName <> "" Then Me.divToUpdate.Add("ArtistDiv", "Artist: " & LblArtistName) Else Me.divToUpdate.Add("ArtistDiv", "")
                End If
                If MusicAPI.CurrentAlbumArtPath.ToString <> MyArt Then
                    MyArt = MusicAPI.CurrentAlbumArtPath.ToString
                    ArtImageBox.imagePathNormal = MyArt
                    ArtImageBox.style = "height:180px;width:180px"
                    ArtImageBox.page = MyPageName & MyClientID
                    Me.divToUpdate.Add("ArtImageDiv", ArtImageBox.Build)
                End If
                If MusicAPI.RadiostationName <> MyRadiostation Then
                    MyRadiostation = MusicAPI.RadiostationName
                    Dim st As String
                    st = MyRadiostation
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblRadiostationName = st
                    If LblRadiostationName <> "" Then Me.divToUpdate.Add("RadioStationDiv", "Radiostation:  " & LblRadiostationName) Else Me.divToUpdate.Add("RadioStationDiv", "")
                End If
                If MusicAPI.NextTrack <> MyNextTrack Then
                    MyNextTrack = MusicAPI.NextTrack
                    Dim st As String
                    st = MyNextTrack
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblNextTrackName = st
                    If LblNextTrackName <> "" Then Me.divToUpdate.Add("NextTrackDiv", "Next Title:  " & LblNextTrackName) Else Me.divToUpdate.Add("NextTrackDiv", "")
                End If
                If MusicAPI.NextAlbum <> MyNextAlbum Then
                    MyNextAlbum = MusicAPI.NextAlbum
                    Dim st As String
                    st = MyNextAlbum
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblNextAlbumName = st
                    If LblNextAlbumName <> "" Then Me.divToUpdate.Add("NextAlbumDiv", "Next Album:  " & LblNextAlbumName) Else Me.divToUpdate.Add("NextAlbumDiv", "")
                End If
                If MusicAPI.NextArtist <> MyNextArtist Then
                    MyNextArtist = MusicAPI.NextArtist
                    Dim st As String
                    st = MyNextArtist
                    If st.Length > 86 Then st = st.Substring(0, 84) & "..."
                    LblNextArtistName = st
                    If LblNextArtistName <> "" Then Me.divToUpdate.Add("NextArtistDiv", "Next Artist: " & LblNextArtistName) Else Me.divToUpdate.Add("NextArtistDiv", "")
                End If
                If MyMute <> MusicAPI.PlayerMute Then
                    MyMute = MusicAPI.PlayerMute
                    If MyMute Then
                        ButMute.imagePathNormal = ImagesPath & "Muted.png"
                    Else
                        ButMute.imagePathNormal = ImagesPath & "UnMuted.png"
                    End If
                    ButMute.style = "height:35px;width:35px"
                    ButMute.page = MyPageName & MyClientID
                    Me.divToUpdate.Add("MuteDiv", ButMute.Build)
                End If
                Dim ShuffleChanged As Boolean = False
                Select Case MusicAPI.ShuffleStatus.ToLower
                    Case "shuffled"
                        ButShuffleImagePath = ImagesPath & "Shuffle.png"
                        If Not MyShuffle Then
                            ShuffleChanged = True
                            MyShuffle = True
                        End If
                    Case Else
                        ButShuffleImagePath = ImagesPath & "NoShuffle.png"
                        If MyShuffle Then
                            ShuffleChanged = True
                            MyShuffle = False
                        End If
                End Select
                ButShuffle.style = "height:40px;width:40px"
                ButShuffle.imagePathNormal = ButShuffleImagePath
                ButShuffle.page = MyPageName & MyClientID
                If ShuffleChanged Then Me.divToUpdate.Add("ShuffleDiv", ButShuffle.Build)
                Dim RepeatChanged As Boolean = False
                Select Case MusicAPI.SonosRepeat
                    Case Repeat_modes.repeat_all, Repeat_modes.repeat_one
                        ButRepeatImagePath = ImagesPath & "Repeat.png"
                        If Not MyRepeat Then
                            RepeatChanged = True
                            MyRepeat = True
                        End If
                    Case Repeat_modes.repeat_off
                        ButRepeatImagePath = ImagesPath & "NoRepeat.png"
                        If MyRepeat Then
                            RepeatChanged = True
                            MyRepeat = False
                        End If
                End Select
                ButRepeat.style = "height:40px;width:40px"
                ButRepeat.imagePathNormal = ButRepeatImagePath
                ButRepeat.page = MyPageName & MyClientID
                If RepeatChanged Then Me.divToUpdate.Add("RepeatDiv", ButRepeat.Build)
                If iPodPlayerName <> MusicAPI.DockediPodPlayerName Then
                    iPodPlayerName = MusicAPI.DockediPodPlayerName
                    If iPodPlayerName <> "" Then Me.divToUpdate.Add("RadioStationDiv", "iPod Player Name:  " & iPodPlayerName) Else Me.divToUpdate.Add("RadioStationDiv", "")
                End If
                Dim TempLinkState As String = ""
                Dim SlaveUDN As String = ""

                If MusicAPI.ZoneIsPairMaster Then
                    TempLinkState = "Paired with " & PIReference.GetZoneNamebyUDN(MusicAPI.GetZonePairSlaveUDN)
                    SlaveUDN = MusicAPI.GetZonePairSlaveUDN
                ElseIf MusicAPI.ZoneIsPairSlave Then
                    TempLinkState = "Paired to " & PIReference.GetZoneNamebyUDN(MusicAPI.GetZonePairMasterUDN)
                ElseIf MusicAPI.PlayBarMaster And MusicAPI.ZoneSource = "TV" Then
                    TempLinkState = "Connected to TV"
                ElseIf MusicAPI.PlayBarSlave Then
                    TempLinkState = "Connected to Playbar"
                End If
                If MusicAPI.ZoneIsLinked And Not MusicAPI.ZoneIsPairSlave Then
                    If TempLinkState <> "" Then TempLinkState &= " "
                    TempLinkState &= "Linked to " & PIReference.GetZoneNamebyUDN(MusicAPI.LinkedZoneSource)
                End If
                Dim TargetZones() = MusicAPI.GetZoneDestination()
                Dim LinkZoneList As String = ""
                If TargetZones IsNot Nothing Then
                    For Each Target As String In TargetZones
                        If Target <> "" And Target <> SlaveUDN Then
                            If LinkZoneList = "" Then LinkZoneList = "Linked with "
                            LinkZoneList &= PIReference.GetZoneNamebyUDN(Target) & " "
                        End If
                    Next
                End If
                If LinkZoneList <> "" Then
                    If TempLinkState <> "" Then TempLinkState &= " "
                    TempLinkState &= LinkZoneList
                End If
                If TempLinkState <> LblLinkState Then
                    LblLinkState = TempLinkState
                    If LblLinkState <> "" Then
                        Me.divToUpdate.Add("LinkStateDiv", "Link State:  " & LblLinkState)
                    Else
                        Me.divToUpdate.Add("LinkStateDiv", "")
                    End If
                End If
                If MusicAPI.MyQueueHasChanged Then
                    MusicAPI.MyQueueHasChanged = False
                    arplaylist = MusicAPI.GetCurrentPlaylistTracks
                    LoadPlayListBox()
                    PlaylistBox.page = MyPageName & MyClientID
                    WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_QueueChange", True)
                    Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
                End If
            ElseIf Control.ToUpper = "QUEUE" Then
                If MusicAPI.MyQueueHasChanged Or GetBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_QueueChange", False) Then
                    MusicAPI.MyQueueHasChanged = False
                    WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_QueueChange", False)
                    arplaylist = MusicAPI.GetCurrentPlaylistTracks
                    LoadPlayListBox()
                    PlaylistBox.page = MyPageName & MyClientID & "&control=" & Control
                    Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
                End If
            ElseIf Control.ToUpper = "NAVPANE" Then
                If GetBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavChange", False) Then
                    WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavChange", False)
                    NavigationBox.page = MyPageName & MyClientID & "&control=" & Control
                    Me.divToUpdate.Add("NavigationDiv", NavigationBox.Build)
                End If
                ' 
            ElseIf Control.ToUpper = "NAVTREE" Then
                If GetBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavTreeChange", False) Then
                    WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavTreeChange", False)
                    LblGenre.page = MyPageName & MyClientID & "&control=" & Control
                    LblArtist.page = MyPageName & MyClientID & "&control=" & Control
                    LblAlbum.page = MyPageName & MyClientID & "&control=" & Control
                    LblPlaylist.page = MyPageName & MyClientID & "&control=" & Control
                    Me.divToUpdate.Add("LabelDiv", LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
                End If
            End If

            If MusicAPI.CheckPlayerIsPairable(ZoneModel) Then
                If (MyPlayerisPairMaster <> MusicAPI.ZoneIsPairMaster) Or (MyPlayerisPairSlave <> MusicAPI.ZoneIsPairSlave) Then
                    MyPlayerisPairMaster = MusicAPI.ZoneIsPairMaster
                    MyPlayerisPairSlave = MusicAPI.ZoneIsPairSlave
                    ButPair.enabled = True
                    ButUnpair.enabled = True
                    ButPair.visible = True
                    ButUnpair.visible = True

                    ButGenres.page = MyPageName & MyClientID & "&control=" & Control
                    ButArtists.page = MyPageName & MyClientID & "&control=" & Control
                    ButAlbums.page = MyPageName & MyClientID & "&control=" & Control
                    ButTracks.page = MyPageName & MyClientID & "&control=" & Control
                    ButPlayLists.page = MyPageName & MyClientID & "&control=" & Control
                    ButRadioStations.page = MyPageName & MyClientID & "&control=" & Control
                    ButAudioBooks.page = MyPageName & MyClientID & "&control=" & Control
                    ButPodCasts.page = MyPageName & MyClientID & "&control=" & Control
                    ButFavorites.page = MyPageName & MyClientID & "&control=" & Control
                    ButLineInput.page = MyPageName & MyClientID & "&control=" & Control
                    ButPair.page = MyPageName & MyClientID & "&control=" & Control
                    ButUnpair.page = MyPageName & MyClientID & "&control=" & Control
                    ButTV.page = MyPageName & MyClientID & "&control=" & Control

                    If MyPlayerisPairMaster Or MyPlayerisPairSlave Then
                        Me.divToUpdate.Add("ButtonDiv", ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButRadioStations.Build & ButLineInput.Build & ButUnpair.Build)
                    Else
                        Me.divToUpdate.Add("ButtonDiv", ButGenres.Build & ButArtists.Build & ButAlbums.Build & ButTracks.Build & ButPlayLists.Build & ButFavorites.Build & ButRadioStations.Build & ButLineInput.Build & ButPair.Build)
                    End If
                End If
            End If
            If MusicAPI.MyZoneGroupStateHasChanged Then
                BuildGroupingInfo(Control)
                GroupingOverlay.page = MyPageName & MyClientID & "&control=" & Control
                Me.divToUpdate.Add("GroupPanel", GroupingOverlay.Build)
            End If
        Catch ex As Exception
            Log("Error in CheckForChanges with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function FormatMyTimeString(AlreadyPlayedSpan As TimeSpan, LblDuration As Integer) As String
        If LblDuration >= (24 * 60 * 10) Then
            ' more than one hour. so report in hour format
            FormatMyTimeString = Format(AlreadyPlayedSpan.Hours, "00") & ":" & Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        ElseIf LblDuration >= (24 * 60) Then
            FormatMyTimeString = Format(AlreadyPlayedSpan.Hours, "0") & ":" & Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        ElseIf LblDuration >= (10 * 60) Then
            FormatMyTimeString = Format(AlreadyPlayedSpan.Minutes, "00") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        Else
            FormatMyTimeString = Format(AlreadyPlayedSpan.Minutes, "0") & ":" & Format(AlreadyPlayedSpan.Seconds, "00")
        End If
    End Function

    Protected Sub ButGenres_Click()
        Try
            ar = MusicAPI.GetGenres()
            LblWhatsLoaded = ITEM_TYPE.genres.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
            LblAudiobooks.label = ""
            LblPodcasts.label = ""
            LblLineInput.label = ""
            LblPairing.label = ""
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub ButArtists_Click()
        Try
            ar = MusicAPI.GetArtists("", "")
            LblWhatsLoaded = ITEM_TYPE.artists.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception

        End Try


    End Sub

    Protected Sub ButAlbums_Click()
        Try
            ar = MusicAPI.GetAlbums("", "")
            LblWhatsLoaded = ITEM_TYPE.albums.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub ButTracks_Click()
        Try
            ar = MusicAPI.DBGetTracks("", "", "")
            Array.Sort(ar)
            LblWhatsLoaded = ITEM_TYPE.tracks.ToString
            ClearSelections()
            LoadNavigationBox(True)
            If ZoneModel = "WD100" Then
                ButAddToPlaylist.enabled = False
            Else
                ButAddToPlaylist.enabled = True
            End If
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub ButPlayLists_Click()
        Try
            ar = MusicAPI.GetPlaylists("", False)
            LblWhatsLoaded = ITEM_TYPE.playlists.ToString
            ClearSelections()
            LoadNavigationBox(True)
            If ZoneModel = "WD100" Then
                ButAddToPlaylist.enabled = False
            Else
                ButAddToPlaylist.enabled = True
            End If
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButRadioStations_Click()
        Try
            ar = MusicAPI.LibGetRadioStationlists()
            LblWhatsLoaded = ITEM_TYPE.radioLists.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButAudiobooks_Click()
        Try
            ar = MusicAPI.LibGetAudiobookslists(iPodPlayerName)
            LblWhatsLoaded = ITEM_TYPE.Audiobooks.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButPodcasts_Click()
        Try
            ar = MusicAPI.LibGetPodcastlists(iPodPlayerName)
            LblWhatsLoaded = ITEM_TYPE.Podcasts.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButFavorites_Click()
        Try
            ar = MusicAPI.LibGetObjectslist("FV:2")
            LblWhatsLoaded = ITEM_TYPE.Favorites.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButLineInput_Click()
        Try
            ar = MusicAPI.LibGetActiveLineInputlists()
            LblWhatsLoaded = ITEM_TYPE.LineInput.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButPair_Click()
        Try
            ar = MusicAPI.LibGetStereoPlayerlist(ZoneModel)
            LblWhatsLoaded = ITEM_TYPE.Pairing.ToString
            ClearSelections()
            LoadNavigationBox(True)
            ButAddToPlaylist.enabled = False
            Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButUnPair_Click()
        Try
            MusicAPI.SeparateStereoPair()
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ButPlay_Click()
        Try
            If MusicAPI.PlayerState = Player_state_values.Paused Then
                'music was paused so resume playing
                MusicAPI.PlayIfPaused()
            ElseIf MusicAPI.PlayerState = Player_state_values.Playing Then
                'music was playing so pause it
                MusicAPI.SonosPause()
            ElseIf LblWhatsLoaded = ITEM_TYPE.radioLists.ToString Then
                LblWhatsLoaded = ITEM_TYPE.radioLists.ToString
                ar = MusicAPI.LibGetRadioStationlists(System.Web.HttpUtility.HtmlDecode(LblRadioList.label))
                LoadNavigationBox()
            ElseIf LblWhatsLoaded = ITEM_TYPE.Audiobooks.ToString Then
                LblWhatsLoaded = ITEM_TYPE.Audiobooks.ToString
                ar = MusicAPI.LibGetAudiobookslists(System.Web.HttpUtility.HtmlDecode(LblAudiobooks.label))
                LoadNavigationBox()
            ElseIf LblWhatsLoaded = ITEM_TYPE.Podcasts.ToString Then
                LblWhatsLoaded = ITEM_TYPE.Podcasts.ToString
                ar = MusicAPI.LibGetPodcastlists(System.Web.HttpUtility.HtmlDecode(LblPodcasts.label))
                LoadNavigationBox()
            ElseIf LblWhatsLoaded = ITEM_TYPE.LineInput.ToString Then
                LblWhatsLoaded = ITEM_TYPE.LineInput.ToString
                ar = MusicAPI.LibGetActiveLineInputlists(System.Web.HttpUtility.HtmlDecode(LblLineInput.label))
                LoadNavigationBox()
            ElseIf LblWhatsLoaded = ITEM_TYPE.Pairing.ToString Then
                LblWhatsLoaded = ITEM_TYPE.Pairing.ToString
                ar = MusicAPI.LibGetStereoPlayerlist(ZoneModel)
                LoadNavigationBox()
            Else
                'nothing was playing, but they hit play.
                'if the current playlist has tracks in it, play that list
                Dim tracks() As String = MusicAPI.GetCurrentPlaylistTracks
                If UBound(tracks) = 0 And tracks(0) = "" Then
                    'MusicAPI.PlayMusic(LblArtist.label, LblAlbum.label, LblPlaylist.label, LblGenre.label, "", "", "", ListBox1.SelectedValue, "", LblAudiobooks.label, LblPodcasts.label) ' changed for SonosController
                    arplaylist = MusicAPI.GetCurrentPlaylistTracks
                    ar = Nothing
                    LoadPlayListBox()
                Else
                    ' build a playlist with selected tracks and play
                    MusicAPI.SonosPlay()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub ButTV_Click()
        Try
            MusicAPI.PlayTV()
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub LblGenre_Click()
        If LblGenre.label = "" Then Exit Sub
        ar = MusicAPI.GetArtists("", System.Web.HttpUtility.HtmlDecode(LblGenre.label))
        LblWhatsLoaded = ITEM_TYPE.artists.ToString
        LblArtist.label = ""
        LblAlbum.label = ""
        LoadNavigationBox()
        ButAddToPlaylist.enabled = False
        AddAjaxDivForNavBox()
    End Sub

    Protected Sub LblArtist_Click()
        If LblArtist.label = "" Then Exit Sub
        ar = MusicAPI.GetAlbums(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label))
        LblWhatsLoaded = ITEM_TYPE.albums.ToString
        LblAlbum.label = ""
        LoadNavigationBox()
        'ListBox1.AutoPostBack = True
        ButAddToPlaylist.enabled = False
        AddAjaxDivForNavBox()
    End Sub

    Protected Sub Double_Click(WhatsLoaded As String, Value As String, Control As String, ClientID As String)
        ' add selected track to existing playlist
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Double_Click called for Player = " & ZoneName & " and MyLastSelecteNavBoxItems = " & Value & " and MyLastSelectedNavBoxClass = " & WhatsLoaded.ToString, LogType.LOG_TYPE_INFO)
        Try
            If WhatsLoaded = "" And Value = "" Then Exit Sub
            Select Case WhatsLoaded
                Case ITEM_TYPE.genres.ToString

                Case ITEM_TYPE.artists.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(Value), "", "")
                Case ITEM_TYPE.albums.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(Value), "")
                Case ITEM_TYPE.tracks.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), System.Web.HttpUtility.HtmlDecode(Value))
                Case ITEM_TYPE.playlists.ToString
                    MusicAPI.PlayMusic("", "", System.Web.HttpUtility.HtmlDecode(Value), "", "", "", "", "", "", "", False, QueueActions.qaDontPlay)
                Case ITEM_TYPE.radioLists.ToString
                Case ITEM_TYPE.Audiobooks.ToString
                Case ITEM_TYPE.Podcasts.ToString
            End Select
            'MyLastSelecteNavBoxItems = ""
            'MyLastSelectedNavBoxClass = ""
            arplaylist = MusicAPI.GetCurrentPlaylistTracks
            LoadPlayListBox()
            PlaylistBox.page = MyPageName & MyClientID & "&control=" & Control
            WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_QueueChange", True)
            Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
        Catch ex As Exception
            Log("Error in Double_Click with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Protected Sub AddTrack_Click()
        ' add selected track to existing playlist
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddTrack_Click called for Player = " & ZoneName & " and MyLastSelecteNavBoxItems = " & MyLastSelecteNavBoxItems.ToString & " and MyLastSelectedNavBoxClass = " & MyLastSelectedNavBoxClass.ToString, LogType.LOG_TYPE_INFO)
        Try
            If MyLastSelecteNavBoxItems = "" And MyLastSelectedNavBoxClass = "" Then Exit Sub
            Select Case MyLastSelectedNavBoxClass
                Case ITEM_TYPE.genres.ToString

                Case ITEM_TYPE.artists.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(LblArtist.label), "", "")
                Case ITEM_TYPE.albums.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), "")
                Case ITEM_TYPE.tracks.ToString
                    MusicAPI.AddTrackToCurrentPlaylist(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), MyLastSelecteNavBoxItems)
                Case ITEM_TYPE.playlists.ToString
                    MusicAPI.PlayMusic("", "", MyLastSelecteNavBoxItems, "", "", "", "", "", "", "", False, QueueActions.qaDontPlay)
                Case ITEM_TYPE.radioLists.ToString
                Case ITEM_TYPE.Audiobooks.ToString
                Case ITEM_TYPE.Podcasts.ToString
            End Select
            'MyLastSelecteNavBoxItems = ""
            'MyLastSelectedNavBoxClass = ""
            arplaylist = MusicAPI.GetCurrentPlaylistTracks
            LoadPlayListBox()
            Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
        Catch ex As Exception
            Log("Error in AddTrack_Click with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Protected Sub ButClearList_Click()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ButClearList_Click called for Player = " & ZoneName, LogType.LOG_TYPE_INFO)
        MusicAPI.ClearCurrentPlayList()
        arplaylist = MusicAPI.GetCurrentPlaylistTracks
        LoadPlayListBox()
        Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
    End Sub

    Protected Sub PlaylistBox_SelectedIndexChanged(value As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("PlaylistBox_SelectedIndexChanged called  for Player = " & ZoneName & " with Value = " & value, LogType.LOG_TYPE_INFO)
        If value = "" Then Exit Sub
        Try
            If PlaylistBox.items.Count > 0 Then
                Dim ListBoxIndex As Integer = 0
                For Each Element In PlaylistBox.items
                    If HttpUtility.UrlDecode(Element.Value) = value Then
                        MusicAPI.SonosSkipToTrack(ListBoxIndex)
                        Exit For
                    End If
                    ListBoxIndex += 1
                Next
            End If
            ' make sure list shows current playlist
            arplaylist = MusicAPI.GetCurrentPlaylistTracks
        Catch ex As Exception
            Log("Error in ListBox1_SelectedIndexChanged with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        LoadPlayListBox()
        Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
    End Sub

    Private Sub AddAjaxDivForNavBox()
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddAjaxDivForNavBox called for Player = " & ZoneName, LogType.LOG_TYPE_INFO)
        Me.divToUpdate.Add("NavigationDiv", NavigationBox.Build)
        Me.divToUpdate.Add("LabelDiv", LblGenre.Build & ">" & LblArtist.Build & ">" & LblAlbum.Build & ">" & LblPlaylist.Build)
    End Sub

    Protected Sub NavigationBox_SelectedIndexChanged(Value As String, DoubleClick As Boolean, ClientID As String, Control As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("NavigationBox_SelectedIndexChanged called for Player = " & ZoneName & " with Value = " & Value & ", Double Click = " & DoubleClick.ToString & " and WhatsLoaded = " & LblWhatsLoaded.ToString, LogType.LOG_TYPE_INFO)
        If Value = "" Then Exit Sub
        MyLastSelecteNavBoxItems = Value
        MyLastSelectedNavBoxClass = LblWhatsLoaded
        If DoubleClick Then
            Double_Click(LblWhatsLoaded, Value, Control, ClientID)
            Exit Sub
        End If
        Try
            ar = Nothing
            arplaylist = Nothing
            Try
                Select Case LblWhatsLoaded
                    Case ITEM_TYPE.genres.ToString
                        ' genres loaded, go to artists
                        LblGenre.label = EncodeTags(Value)
                        LblWhatsLoaded = ITEM_TYPE.artists.ToString
                        ar = MusicAPI.GetArtists("", System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        LoadNavigationBox()
                        AddAjaxDivForNavBox()
                    Case ITEM_TYPE.artists.ToString
                        ' artists are loaded, go to albums
                        LblArtist.label = EncodeTags(Value)
                        LblWhatsLoaded = ITEM_TYPE.albums.ToString
                        ar = MusicAPI.GetAlbums(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        LoadNavigationBox()
                        AddAjaxDivForNavBox()
                    Case ITEM_TYPE.albums.ToString
                        ' albums are loaded, list tracks
                        LblAlbum.label = EncodeTags(Value)
                        LblWhatsLoaded = ITEM_TYPE.tracks.ToString
                        ar = MusicAPI.DBGetTracks(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label))
                        LoadNavigationBox()

                        If ZoneModel <> "WD100" Then
                            ' the Wireless Dock doesn't support queues. Play direct
                            ButAddToPlaylist.enabled = True
                        Else
                            ButAddToPlaylist.enabled = False
                        End If
                        AddAjaxDivForNavBox()
                        Me.divToUpdate.Add("ButAddToPlaylistDiv", ButAddToPlaylist.Build)
                    Case ITEM_TYPE.tracks.ToString
                        If ZoneModel = "WD100" Then
                            ' the Wireless Dock doesn't support queues. Play direct
                            MusicAPI.PlayMusic(System.Web.HttpUtility.HtmlDecode(LblArtist.label), System.Web.HttpUtility.HtmlDecode(LblAlbum.label), System.Web.HttpUtility.HtmlDecode(LblPlaylist.label), System.Web.HttpUtility.HtmlDecode(LblGenre.label), "", "", Value, "", "", "", False, QueueActions.qaPlayNow)
                        End If
                    Case ITEM_TYPE.playlists.ToString
                        LblPlaylist.label = EncodeTags(Value)
                        MusicAPI.PlayMusic("", "", Value, "", "", "", "", "", "", "", False, QueueActions.qaDontPlay)
                        arplaylist = MusicAPI.GetCurrentPlaylistTracks
                        LoadPlayListBox()
                        Me.divToUpdate.Add("PlaylistDiv", PlaylistBox.Build)
                    Case ITEM_TYPE.radioLists.ToString
                        LblRadioList.label = EncodeTags(Value)
                        MusicAPI.PlayMusic("", "", Value, "", "", "", "", "", "", "", False, QueueActions.qaPlayNow)
                    Case ITEM_TYPE.Audiobooks.ToString
                        LblAudiobooks.label = EncodeTags(Value)
                        MusicAPI.PlayMusic("", "", "", "", "", "", "", "", Value, "", False, QueueActions.qaPlayNow)
                    Case ITEM_TYPE.Podcasts.ToString
                        LblPodcasts.label = EncodeTags(Value)
                        MusicAPI.PlayMusic("", "", "", "", "", "", "", "", "", Value, False, QueueActions.qaPlayNow)
                    Case ITEM_TYPE.Favorites.ToString
                        LblFavorites.label = EncodeTags(Value)
                        MusicAPI.PlayFavorite(Value, True, QueueActions.qaPlayNow)
                    Case ITEM_TYPE.LineInput.ToString
                        LblLineInput.label = EncodeTags(Value)
                        MusicAPI.PlayLineInput(Value)
                    Case ITEM_TYPE.Pairing.ToString
                        LblPairing.label = EncodeTags(Value)
                        MusicAPI.CreateStereoPair(PIReference.GetUDNbyZoneName(Value))
                    Case Else
                End Select
                SaveNavSettingstoINIFile(ClientID)
            Catch ex As Exception
                Log("Error in NavigationBox_SelectedIndexChanged for Player = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SaveNavSettingstoINIFile(ClientID As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SaveNavSettingstoINIFile called for Zoneplayer = " & ZoneName & " with ClientID = " & ClientID.ToString, LogType.LOG_TYPE_INFO)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_loaded", LblWhatsLoaded)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_genre", LblGenre.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_artist", LblArtist.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_album", LblAlbum.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_playlist", LblPlaylist.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_radiolist", LblRadioList.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_audiobook", LblAudiobooks.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_podcast", LblPodcasts.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_favorites", LblFavorites.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_lineinput", LblLineInput.label)
        WriteStringIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_pairing", LblPairing.label)
        WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavChange", True)
        WriteBooleanIniFile("SonosMusicPage", MyZoneUDN & ClientID & "_NavTreeChange", True)
    End Sub

    Protected Sub ProcessGroupingChange(GroupingData As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange called for Zoneplayer = " & ZoneName & " with GroupingData = " & GroupingData.ToString, LogType.LOG_TYPE_INFO)
        Dim PlayerUDNArray() As String
        Dim PlayerGroupingArray() As GroupArrayElement = Nothing
        Try
            PlayerUDNArray = PIReference.GetAllActiveZones()
            If PlayerUDNArray IsNot Nothing Then
                If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange called for Player = " & ZoneName & " and has " & (UBound(PlayerUDNArray) + 1).ToString & " entries", LogType.LOG_TYPE_INFO)
                For i = 0 To UBound(PlayerUDNArray)
                    Dim Player As HSPI = PIReference.GetAPIByUDN(PlayerUDNArray(i))
                    If Player.ZoneModel.ToUpper <> "WD100" And Player.ZoneModel.ToUpper <> "SUB" And Not Player.ZoneIsPairSlave Then
                        If (Player.ZoneIsLinked And Player.LinkedZoneSource = MyZoneUDN) Then
                            ' this is a player linked to this zone
                            Dim GroupingElement As New GroupArrayElement
                            GroupingElement.UDN = Player.GetUDN
                            GroupingElement.Master = False
                            If PlayerGroupingArray Is Nothing Then
                                ReDim PlayerGroupingArray(0)
                                PlayerGroupingArray(0) = GroupingElement
                            Else
                                ReDim Preserve PlayerGroupingArray(PlayerGroupingArray.Count)
                                PlayerGroupingArray(PlayerGroupingArray.Count - 1) = GroupingElement
                            End If
                        ElseIf Player.GetUDN = MyZoneUDN Then
                            Dim GroupingElement As New GroupArrayElement
                            GroupingElement.UDN = MyZoneUDN
                            If Player.ZoneIsLinked Then
                                GroupingElement.Master = False
                            ElseIf Player.GetTargetZoneLinkedList <> "" Then
                                GroupingElement.Master = True
                            End If
                            If PlayerGroupingArray Is Nothing Then
                                ReDim PlayerGroupingArray(0)
                                PlayerGroupingArray(0) = GroupingElement
                            Else
                                ReDim Preserve PlayerGroupingArray(PlayerGroupingArray.Count)
                                PlayerGroupingArray(PlayerGroupingArray.Count - 1) = GroupingElement
                            End If
                        ElseIf Player.ZoneIsLinked Then
                            If PIReference.GetAPIByUDN(Player.LinkedZoneSource).GetTargetZoneLinkedList.IndexOf(MyZoneUDN) <> -1 Then
                                ' this is a player linked to this zone
                                Dim GroupingElement As New GroupArrayElement
                                GroupingElement.UDN = Player.GetUDN
                                GroupingElement.Master = False
                                If PlayerGroupingArray Is Nothing Then
                                    ReDim PlayerGroupingArray(0)
                                    PlayerGroupingArray(0) = GroupingElement
                                Else
                                    ReDim Preserve PlayerGroupingArray(PlayerGroupingArray.Count)
                                    PlayerGroupingArray(PlayerGroupingArray.Count - 1) = GroupingElement
                                End If
                            End If
                        ElseIf Player.GetTargetZoneLinkedList.IndexOf(MyZoneUDN) <> -1 Then
                            ' this must be the master
                            Dim GroupingElement As New GroupArrayElement
                            GroupingElement.UDN = Player.GetUDN
                            GroupingElement.Master = True
                            If PlayerGroupingArray Is Nothing Then
                                ReDim PlayerGroupingArray(0)
                                PlayerGroupingArray(0) = GroupingElement
                            Else
                                ReDim Preserve PlayerGroupingArray(PlayerGroupingArray.Count)
                                PlayerGroupingArray(PlayerGroupingArray.Count - 1) = GroupingElement
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            Log("Error in ProcessGroupingChange for PlayerControl for Player = " & ZoneName & " trying to check linked zones with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange called for Player = " & ZoneName & " Found " & PlayerGroupingArray.Count.ToString & " linked players", LogType.LOG_TYPE_INFO)
        Dim parts As Collections.Specialized.NameValueCollection
        parts = HttpUtility.ParseQueryString(GroupingData)
        If parts IsNot Nothing Then
            Try
                Dim Part As String
                For Each Part In parts.AllKeys
                    If Part IsNot Nothing Then
                        Dim ObjectNameParts As String()
                        ObjectNameParts = Split(Part, "-")
                        Dim ObjectValue As String = System.Web.HttpUtility.HtmlDecode(parts(Part))
                        Select Case ObjectNameParts(0).ToString.ToUpper
                            Case "OVSUBMIT"
                                ' ignore
                            Case "LABEL"
                                Dim Found As Boolean = False
                                If PlayerGroupingArray.Count <> 0 Then
                                    For Each Groupelement In PlayerGroupingArray
                                        If Groupelement.UDN = ObjectNameParts(1).ToString Then
                                            ' match
                                            Groupelement.Confirmed = True
                                            Found = True
                                            Exit For
                                        End If
                                    Next
                                End If
                                If Not Found Then
                                    ' add it
                                    Dim GroupingElement As New GroupArrayElement
                                    GroupingElement.UDN = ObjectNameParts(1).ToString
                                    GroupingElement.Added = True
                                    GroupingElement.Confirmed = True
                                    If PlayerGroupingArray Is Nothing Then
                                        ReDim PlayerGroupingArray(0)
                                        PlayerGroupingArray(0) = GroupingElement
                                    Else
                                        ReDim Preserve PlayerGroupingArray(PlayerGroupingArray.Count)
                                        PlayerGroupingArray(PlayerGroupingArray.Count - 1) = GroupingElement
                                    End If
                                End If
                        End Select
                    End If
                Next
            Catch ex As Exception
                Log("Error in ProcessGroupingChange for Zone = " & ZoneName & " and GroupingData = " & GroupingData.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try
        End If
        If PlayerGroupingArray.Count <= 1 Then Exit Sub ' we're done, there is nothing to do
        Dim OldMasterUDN As String = ""
        Dim OldMasterStillMaster As Boolean = False
        Dim NewMasterUDN As String = ""

        Try
            ' check for player to be removed
            For Each Groupelement In PlayerGroupingArray
                If Not Groupelement.Confirmed And Not Groupelement.Master Then ' only remove non master players
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange for Player = " & ZoneName & " is removing a player = " & Groupelement.UDN.ToString, LogType.LOG_TYPE_INFO)
                    PIReference.GetAPIByUDN(Groupelement.UDN).Unlink()
                End If
            Next
        Catch ex As Exception
            Log("Error in ProcessGroupingChange for Zone = " & ZoneName & " removing players and GroupingData = " & GroupingData.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            ' Find the master
            For Each Groupelement In PlayerGroupingArray
                If Groupelement.Master Then
                    OldMasterUDN = Groupelement.UDN
                    OldMasterStillMaster = Groupelement.Confirmed
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange for Player = " & ZoneName & " found Master = " & OldMasterUDN.ToString & " and Active State = " & OldMasterStillMaster.ToString, LogType.LOG_TYPE_INFO)
                    Exit For
                End If
            Next
        Catch ex As Exception
            Log("Error in ProcessGroupingChange for Zone = " & ZoneName & " checking for Master and GroupingData = " & GroupingData.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If OldMasterUDN = "" Then
            ' there is no master yet, this is the start of a grouping. 
            ' pick the next new master
            For Each Groupelement In PlayerGroupingArray
                If Not Groupelement.Master And Groupelement.Confirmed Then
                    OldMasterUDN = Groupelement.UDN
                    OldMasterStillMaster = True
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange for Player = " & ZoneName & " found a new Master = " & NewMasterUDN.ToString, LogType.LOG_TYPE_INFO)
                    Exit For
                End If
            Next
        End If
        Try
            ' check for players to be added
            For Each Groupelement In PlayerGroupingArray
                If Groupelement.Added And Not Groupelement.Master Then
                    If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange for Player = " & ZoneName & " is adding a player = " & Groupelement.UDN.ToString, LogType.LOG_TYPE_INFO)
                    PIReference.GetAPIByUDN(Groupelement.UDN).Link("uuid:" & OldMasterUDN)
                End If
            Next
        Catch ex As Exception
            Log("Error in ProcessGroupingChange for Zone = " & ZoneName & " adding players and GroupingData = " & GroupingData.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not OldMasterStillMaster Then
                ' pick the next new master
                For Each Groupelement In PlayerGroupingArray
                    If Not Groupelement.Master And Groupelement.Confirmed Then
                        NewMasterUDN = Groupelement.UDN
                        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ProcessGroupingChange for Player = " & ZoneName & " found a new Master = " & NewMasterUDN.ToString, LogType.LOG_TYPE_INFO)
                        Exit For
                    End If
                Next
            End If
            If Not OldMasterStillMaster And NewMasterUDN <> "" Then
                ' go change the master
                Dim OldMaster As HSPI = PIReference.GetAPIByUDN(OldMasterUDN)
                OldMaster.DelegateGroupCoordinationTo(NewMasterUDN, False)
            End If
        Catch ex As Exception
            Log("Error in ProcessGroupingChange for Zone = " & ZoneName & " assigning new Master and GroupingData = " & GroupingData.ToString & " and Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



End Class

<Serializable()> Public Class GroupArrayElement
    Public UDN As String
    Public Master As Boolean = False
    Public Added As Boolean = False
    Public Confirmed As Boolean = False
End Class
