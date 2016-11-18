<Serializable()> _
Public Class IniFile
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString _
      Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, _
      ByVal nSize As Integer, ByVal lpFileName As String) _
      As Integer
    Private Declare Ansi Function WritePrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpString As String, _
      ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt _
      Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal nDefault As Integer, _
      ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, _
      ByVal lpKeyName As Integer, ByVal lpString As Integer, _
      ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileSection _
     Lib "kernel32.dll" Alias "GetPrivateProfileSectionA" _
     (ByVal lpAppName As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Integer, ByVal lpFileName As String) _
    As Integer  '

    Dim strFilename As String

    ' Constructor, accepting a filename
    Public Sub New(ByVal Filename As String)
        strFilename = Filename
    End Sub


    Protected Overrides Sub Finalize()
        Try
            strFilename = Nothing
        Catch ex As Exception
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function GetString_(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(4096) ' changed from 256
        intCharCount = GetPrivateProfileString(Section, Key, _
           [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then
            GetString_ = Left(objResult.ToString, intCharCount)
        Else
            GetString_ = ""
        End If
    End Function

    Public Function GetInteger_(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key, _
           [Default], strFilename)
    End Function

    Public Function GetBoolean_(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key, _
           CInt([Default]), strFilename) = 1)
    End Function

    Public Sub WriteString_(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger_(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString_(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean_(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString_(Section, Key, CStr(Convert.ToInt32(Value)))
        Flush()
    End Sub

    Public Overloads Sub DeleteEntry_(ByVal Section As String, _
      ByVal Key As String) ' delete single line from section
        WritePrivateProfileString(Section, Key, Nothing, strFilename)
        Flush()
    End Sub


    Public Function GetSection_(ByVal Section As String) As Dictionary(Of String, String)
        ' Returns a whole section from your INI file
        Dim intCharCount As Integer
        Dim objResult As String = Space(4096)
        Try
            intCharCount = GetPrivateProfileSection(Section, objResult, objResult.Length, strFilename)
        Catch ex As Exception
        End Try
        Dim TempString As String = ""
        Dim StringIndex As Integer = 0
        Dim OutArray() As String = {""}
        Try
            Dim tempstrings
            tempstrings = Split(objResult, Chr(0))
            For Each TempString In tempstrings
                TempString = Trim(TempString)
                If TempString <> "" Then
                    ReDim Preserve OutArray(StringIndex)
                    OutArray(StringIndex) = TempString
                    StringIndex = StringIndex + 1
                End If
            Next
        Catch ex As Exception
            Log("Error in GetSection with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Log("GetSection called with section = " & Section & " and # Result = " & OutArray.Count.ToString, LogType.LOG_TYPE_INFO)
        objResult = Nothing
        Dim KeyValues As New Dictionary(Of String, String)()
        For i As Integer = 0 To OutArray.Length - 1
            Dim SeparatorPosition As Integer = InStr(OutArray(i), "=")
            Dim Key, Value As String
            If SeparatorPosition > 0 Then
                Key = Left(OutArray(i), SeparatorPosition - 1)
                Value = Mid(OutArray(i), SeparatorPosition + 1)
            Else
                Key = OutArray(i)
                Value = Nothing
            End If
            KeyValues.Add(Key, Value)
        Next i
        Log("GetSection called with section = " & Section & " and # Keys = " & KeyValues.Count.ToString, LogType.LOG_TYPE_INFO)
        GetSection_ = KeyValues
    End Function

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub

End Class
