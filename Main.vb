Imports Scheduler
Imports HomeSeerAPI
Imports HSCF.Communication.Scs.Communication.EndPoints.Tcp
Imports HSCF.Communication.ScsServices.Client
Imports HSCF.Communication.ScsServices.Service


Module Main
    Public WithEvents client As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
    Dim WithEvents clientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)
    Public MainInstance As String = ""
    Public ServerIPAddress As String = ""

    Private host As HomeSeerAPI.IHSApplication

    Private gAppAPI As HSPI
    Private sIp As String = "127.0.0.1"

    Friend colTrigs_Sync As System.Collections.SortedList
    Friend colTrigs As System.Collections.SortedList
    Friend colActs_Sync As System.Collections.SortedList
    Friend colActs As System.Collections.SortedList

    Public AllInstances As New SortedList

    Public Class InstanceHolder
        Public hspi As HSPI
        Public client As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
        Public clientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)
        Public host As HomeSeerAPI.IHSApplication
    End Class

    Public Function AddInstance(InstanceName As String) As HSPI
        If g_bDebug Then Log("AddInstance called with InstanceName = " & InstanceName, LogType.LOG_TYPE_INFO)
        AddInstance = Nothing
        If AllInstances.Contains(InstanceName) Then
            If g_bDebug Then Log("AddInstance called with InstanceName = " & InstanceName & " but instance already exists", LogType.LOG_TYPE_WARNING)
            Exit Function
        End If
        Dim PlugAPI As HSPI = New HSPI
        PlugAPI.instance = InstanceName
        PlugAPI.isRoot = False

        Dim lhost As HomeSeerAPI.IHSApplication
        Dim lclient As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
        Dim lclientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)

        lclient = ScsServiceClientBuilder.CreateClient(Of IHSApplication)(New ScsTcpEndPoint(sIp, 10400), PlugAPI)
        lclientCallback = ScsServiceClientBuilder.CreateClient(Of IAppCallbackAPI)(New ScsTcpEndPoint(sIp, 10400), PlugAPI)

        Try
            lclient.Connect()
            lclientCallback.Connect()
            lhost = lclient.ServiceProxy
        Catch ex As Exception
            Log("Error in AddInstance for InstanceName = " & InstanceName & ". Cannot start instance with Error = " & ":" & ex.Message, LogType.LOG_TYPE_ERROR)
            PlugAPI.DestroyPlayer(True)
            PlugAPI = Nothing
            Exit Function
        End Try

        Try
            lhost.Connect(sIFACE_NAME, InstanceName)

            ' everything is ok, save instance
            Dim ih As New InstanceHolder
            ih.client = lclient
            ih.clientCallback = lclientCallback
            ih.host = lhost
            ih.hspi = PlugAPI
            AllInstances.Add(InstanceName, ih)
        Catch ex As Exception
            PlugAPI.DestroyPlayer(True)
            PlugAPI = Nothing
            Log("Error in AddInstance connecting or disconnecting InstanceName = " & InstanceName & " with Error = " & ":" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Return PlugAPI
    End Function


    Public Function RemoveInstance(InstanceName As String) As String
        If g_bDebug Then Log("RemoveInstance called with InstanceName = " & InstanceName, LogType.LOG_TYPE_INFO)
        If Not AllInstances.Contains(InstanceName) Then
            Return "Instance does not exist"
        End If
        Dim DE As DictionaryEntry
        Try
            For Each DE In AllInstances
                Dim key As String = DE.Key
                If key.ToLower = InstanceName.ToLower Then
                    Dim ih As InstanceHolder = DE.Value
                    ih.hspi.ShutdownIO()
                    ih.client.Disconnect()
                    ih.client = Nothing
                    ih.clientCallback.Disconnect()
                    ih.clientCallback = Nothing
                    AllInstances.Remove(key)
                    Exit For
                End If
            Next
        Catch ex As Exception
            If g_bDebug Then Log("Error in RemoveInstance for InstanceName = " & InstanceName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return "Error removing instance: " & ex.Message
        End Try
        Return ""
    End Function

    Sub Main()

        Dim argv As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        argv = My.Application.CommandLineArgs

        Dim sCmd As String
        For Each sCmd In argv
            Dim ch(0) As String
            ch(0) = "="
            Dim parts() As String = sCmd.Split(ch, StringSplitOptions.None)
            Select Case parts(0).ToLower
                Case "server"
                    sIp = parts(1)
                    ServerIPAddress = sIp
                Case "instance"
                    Try
                        MainInstance = parts(1)
                    Catch ex As Exception
                        Maininstance = ""
                    End Try
            End Select
        Next

        gAppAPI = New hspi

        Console.WriteLine("Connecting to server at " & sIp & " with Instance = '" & Maininstance & "' ...")
        client = ScsServiceClientBuilder.CreateClient(Of IHSApplication)(New ScsTcpEndPoint(sIp, 10400), gAppAPI)
        clientCallback = ScsServiceClientBuilder.CreateClient(Of IAppCallbackAPI)(New ScsTcpEndPoint(sIp, 10400), gAppAPI)

        ' My code

        If Maininstance <> "" Then
            tIFACE_NAME = tIFACE_NAME & "_" & Maininstance.ToString
            MusicDBPath = "/html/" & tIFACE_NAME & "/MusicDB/SonosDB.sdb"
            RadioStationsDBPath = "/html/" & tIFACE_NAME & "/MusicDB/SonosRadioStationsDB.sdb"
            DockedPlayersDBPath = "/html/" & tIFACE_NAME & "/MusicDB/"
            AnnouncementPath = "/" & tIFACE_NAME & "/Announcements/"
            ArtWorkPath = "/images/" & tIFACE_NAME & "/Artwork/"
            URLArtWorkPath = "/images/" & tIFACE_NAME & "/Artwork/"  '"/" & tIFACE_NAME & "/Artwork/"  
            FileArtWorkPath = tIFACE_NAME & "\Artwork\"
            'DebugLogFileName = "/" & tIFACE_NAME & "/Logs/SonosDebug.txt"
            MyINIFile = tIFACE_NAME & ".ini"    ' Configuration File
        End If


        Dim Attempts As Integer = 1

TryAgain:
        Try
            client.Connect()
            clientCallback.Connect()

            host = client.ServiceProxy

            Dim APIVersion As Double = host.APIVersion  ' will cause an error if not really connected

            callback = clientCallback.ServiceProxy
            APIVersion = callback.APIVersion  ' will cause an error if not really connected
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Cannot connect attempt " & Attempts.ToString & ": " & ex.Message, LogType.LOG_TYPE_ERROR)
            If ex.Message.ToLower.Contains("timeout occurred.") Then
                Attempts += 1
                If Attempts < 6 Then GoTo TryAgain
            End If

            If client IsNot Nothing Then
                client.Dispose()
                client = Nothing
            End If
            If clientCallback IsNot Nothing Then
                clientCallback.Dispose()
                clientCallback = Nothing
            End If
            wait(4)
            Return
        End Try

        Try
            ' connect to HS so it can register a callback to us
            host.Connect(sIFACE_NAME, Maininstance)

            ' create the user object that is the real plugin, accessed from the pluginAPI wrapper
            callback = callback
            hs = host

            Console.WriteLine("Connected, waiting to be initialized...")
            Do
                Threading.Thread.Sleep(30)
            Loop While client.CommunicationState = HSCF.Communication.Scs.Communication.CommunicationStates.Connected And Not bShutDown And Not MyShutDownRequest
            If Not bShutDown Then
                Try
                    Dim DE As DictionaryEntry
                    For Each DE In AllInstances
                        'RemoveInstance(DE.Key)
                        Dim ih As InstanceHolder = DE.Value
                        Try
                            ih.client.Disconnect()
                            ih.client = Nothing
                            ih.clientCallback.Disconnect()
                            ih.clientCallback = Nothing
                        Catch ex As Exception
                            Log("Error in main shutting down Instance = " & DE.Key & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    Next
                Catch ex As Exception
                    Log("Error in main shutting down Instances with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                gAppAPI.ShutdownIO()
                Console.WriteLine("Connection lost, exiting")
            Else
                Console.WriteLine("Shutting down plugin")
            End If
            ' disconnect from server for good here
            hs = Nothing
            client.Disconnect()
            clientCallback.Disconnect()
            wait(2)
            End
        Catch ex As Exception
            If Not ImRunningOnLinux Then Console.WriteLine("Cannot connect(2): " & ex.Message, LogType.LOG_TYPE_ERROR)
            wait(2)
            End
            Return
        End Try

    End Sub


    Private Sub client_Disconnected(ByVal sender As Object, ByVal e As System.EventArgs) Handles client.Disconnected
        Console.WriteLine("Disconnected from server - client")
    End Sub


    Private Sub wait(ByVal secs As Integer)
        If g_bDebug Then Log("Wait in Main called with Wait = " & secs.ToString, LogType.LOG_TYPE_INFO)
        Threading.Thread.Sleep(secs * 1000)
    End Sub

    
End Module
