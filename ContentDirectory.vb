Imports System.Xml
Imports System.Drawing


Partial Public Class HSPI

    Private MyMaxDepthCounter As Integer = 0

    Public Function GetContainerFromServer(ObjectID As String) As DBRecord()
        If g_bDebug Then Log("GetContainerFromServer called for device - " & ZoneName & " with ObjectID=" & ObjectID, LogType.LOG_TYPE_INFO)
        GetContainerFromServer = Nothing
        If ContentDirectory Is Nothing Then
            If g_bDebug Then Log("GetContainerFromServer called for device - " & ZoneName & " but no handle to ContentDirectory", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If

        Dim SList As DBRecord() = Nothing
        Dim LoopIndex As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0
        'Dim ObjectFilter As String = "dc:title,upnp:album,upnp:artist,upnp:genre,upnp:albumArtURI,res"
        Dim ObjectFilter As String = "dc:title"

        Dim InArg(5)
        Dim OutArg(3)
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = ObjectFilter             ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = MaxNbrOfUPNPObjects      ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String


        StartIndex = 0
        Dim RecordIndex As Integer = 0

        Do

            InArg(3) = StartIndex               ' Index         UI4

            Try
                Call ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in GetContainerFromServer for device = " & ZoneName & " with ObjectID = " & ObjectID.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)

            If g_bDebug Then Log("GetContainerFromServer found " & TotalMatches.ToString & " entries and " & TotalMatches.ToString & " total matched for device = " & ZoneName & " and ObjectID = " & ObjectID.ToString, LogType.LOG_TYPE_INFO)

            If NumberReturned = 0 Then
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End If

            Dim xmlData As XmlDocument = New XmlDocument
            Try
                xmlData.LoadXml(OutArg(0).ToString)
                'hs.Writelog(IFACE_NAME, "XML=" & Value.ToString) ' used for testing
            Catch ex As Exception
                Log("Error in GetContainerFromServer at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & ZoneName & " and ObjectID = " & ObjectID.ToString & "  loading XML. Error = " & ex.Message & ". XML = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
                GetContainerFromServer = SList
                SList = Nothing
                Exit Function
            End Try
            Try
                'Get a list of all the child elements
                Dim nodelist As XmlNodeList = xmlData.DocumentElement.ChildNodes
                If SuperDebug Then Log("GetContainerFromServer Nbr of items in XML Data = " & nodelist.Count, LogType.LOG_TYPE_INFO) ' this starts with <Event>
                If SuperDebug Then Log("GetContainerFromServer Document root node: " & xmlData.DocumentElement.Name, LogType.LOG_TYPE_INFO)
                'Parse through all nodes
                For Each outerNode As XmlNode In nodelist
                    If UCase(outerNode.Name) = "CONTAINER" Then
                        Dim NewRecord = New DBRecord
                        Dim ObjectClass As String() = Nothing
                        NewRecord.Id = ""
                        NewRecord.ParentID = ""
                        NewRecord.Title = ""
                        NewRecord.AlbumName = ""
                        NewRecord.ArtistName = ""
                        NewRecord.Genre = ""
                        NewRecord.IconURL = ""
                        Try
                            NewRecord.Id = outerNode.Attributes("id").Value
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ParentID = outerNode.Attributes("parentID").Value
                        Catch ex As Exception
                        End Try

                        Try
                            NewRecord.Title = outerNode.Item("dc:title").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ItemClass = "CONTAINER"
                        Catch ex As Exception
                        End Try

                        If SList Is Nothing Then
                            ReDim SList(0)
                            SList(0) = NewRecord
                        Else
                            ReDim Preserve SList(RecordIndex)
                            SList(RecordIndex) = NewRecord
                        End If
                        If SuperDebug Then Log("Record #" & RecordIndex.ToString & " with value " & NewRecord.Title, LogType.LOG_TYPE_INFO)
                        RecordIndex = RecordIndex + 1
                    ElseIf UCase(outerNode.Name) = "ITEM" Then
                        Dim NewRecord = New DBRecord
                        Dim ObjectClass As String() = Nothing
                        NewRecord.Id = ""
                        NewRecord.ParentID = ""
                        NewRecord.Title = ""
                        NewRecord.AlbumName = ""
                        NewRecord.ArtistName = ""
                        NewRecord.Genre = ""
                        NewRecord.IconURL = ""
                        Dim ClassInfo As String = ""
                        Try
                            NewRecord.Id = outerNode.Attributes("id").Value
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ParentID = outerNode.Attributes("parentID").Value
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.Title = outerNode.Item("dc:title").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.AlbumName = outerNode.Item("upnp:album").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.ArtistName = outerNode.Item("upnp:artist").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            NewRecord.Genre = outerNode.Item("upnp:genre").InnerText
                        Catch ex As Exception
                        End Try
                        Try
                            ClassInfo = outerNode.Item("upnp:class").InnerText
                            NewRecord.ClassType = ProcessClassInfo(ClassInfo)
                        Catch ex As Exception
                        End Try
                        If NewRecord.ClassType = UPnPClassType.ctPictures Then
                            Try
                                NewRecord.IconURL = outerNode.Item("res").InnerText
                            Catch ex As Exception
                            End Try
                        Else
                            Try
                                NewRecord.IconURL = outerNode.Item("upnp:albumArtURI").InnerText
                            Catch ex As Exception
                            End Try
                        End If
                        If NewRecord.IconURL = "yyyy" Then
                            Dim IConImage As Bitmap
                            Try
                                IConImage = New Bitmap(NewRecord.IconURL)
                            Catch ex As Exception
                                'just use the no art image
                                IConImage = New Bitmap(hs.GetAppPath.ToString & NoArtPath)
                                Log("Error in GetContainerFromServer for device - " & ZoneName & " getting imageart with error = " & ex.Message & " Path = " & MyIconURL.ToString, LogType.LOG_TYPE_ERROR)
                            End Try
                            'Dim IConImage As Image
                            'IConImage. = 150
                            Try
                                Dim propItem As System.Drawing.Imaging.PropertyItem
                                For Each propItem In IConImage.PropertyItems
                                    Log("Picture info in GetContainerFromServer for device - " & ZoneName & " Id = " & propItem.Id.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & ZoneName & " Type = " & propItem.Type.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & ZoneName & " Value = " & propItem.Value.ToString, LogType.LOG_TYPE_INFO)
                                    Log("Picture info in GetContainerFromServer for device - " & ZoneName & " Length = " & propItem.Len.ToString, LogType.LOG_TYPE_INFO)
                                Next
                            Catch ex As Exception
                                Log("Error in GetContainerFromServer for device - " & ZoneName & " getting imageinfo with error = " & ex.Message & " Path = " & MyIconURL.ToString, LogType.LOG_TYPE_ERROR)
                            End Try

                            IConImage.Dispose()


                            'GetPicture(MyIconURL)
                        End If
                        Try
                            NewRecord.ItemClass = "ITEM"
                        Catch ex As Exception
                        End Try
                        If SList Is Nothing Then
                            ReDim SList(0)
                            SList(0) = NewRecord
                        Else
                            ReDim Preserve SList(RecordIndex)
                            SList(RecordIndex) = NewRecord
                        End If
                        If SuperDebug Then Log("Record #" & RecordIndex.ToString & " with value " & NewRecord.Title, LogType.LOG_TYPE_INFO)
                        RecordIndex = RecordIndex + 1
                    Else
                        If g_bDebug Then Log("Error in GetContainerFromServer for UPnPDevice = " & ZoneName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes found node = " & outerNode.Name.ToString, LogType.LOG_TYPE_ERROR)
                    End If
                Next
            Catch ex As Exception
                Log("Error in GetContainerFromServer at level = " & MyMaxDepthCounter.ToString & " for UPnPDevice = " & ZoneName & " and ObjectID = " & ObjectID.ToString & "  processing Childnodes with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            'hs.WaitEvents()
        Loop

        GetContainerFromServer = SList
        SList = Nothing

        If g_bDebug Then Log("GetContainerFromServer for device - " & ZoneName & " returned " & RecordIndex.ToString & " Server records ", LogType.LOG_TYPE_INFO)

    End Function


    Private Function ProcessClassInfo(ClassString As String) As UPnPClassType
        ProcessClassInfo = UPnPClassType.ctUnknown
        If ClassString = "" Then Exit Function
        Dim ClassItems As String()
        ClassItems = ClassString.Split(".")
        If UBound(ClassItems) > 1 Then
            Select Case UCase(ClassItems(2))
                Case "IMAGEITEM"
                    ProcessClassInfo = UPnPClassType.ctPictures
                Case "AUDIOITEM"
                    ProcessClassInfo = UPnPClassType.ctMusic
                Case "VIDEOITEM"
                    ProcessClassInfo = UPnPClassType.ctVideo
                Case "PLAYLISTITEM"
                    ProcessClassInfo = UPnPClassType.ctMusic
                    'Case "TEXTITEM"
                    'Case "BOOKMARKITEM"
                    'Case "EPGITEM"
                Case Else
                    If g_bDebug Then Log("Warning ProcessClassInfo called with unknown UPnPClass = " & ClassString, LogType.LOG_TYPE_WARNING)


            End Select
        End If

    End Function

End Class
