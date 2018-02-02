Public Class Registry

    Implements IDisposable

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region


    Public HiveName As String
    Public KeyName As String
    Public ValueName As String
    Public ValueDefault As String
    Private iRepeat As Integer
    Private mstrHive As String
    Private mstrKey As String
    Private HKEY_CROOT As String = "HKEY_CLASSES_ROOT"
    Private HKEY_CUSER As String = "HKEY_CURRENT_USER"
    Private HKEY_LM As String = "HKEY_LOCAL_MACHINE"
    Private HKEY_USERS As String = "HKEY_USERS"
    Private HKEY_CCONF As String = "HKEY_CURRENT_CONFIG"

    Public Function GetKey( _
        Optional ByVal strHiveName As String = "", _
        Optional ByVal strKeyName As String = "", _
        Optional ByVal strValueName As String = "") As String
        'HKEY_CURRENT_USER\Software\<CompanyName>\<AppName>
        Dim s$ = ""
        Try
            If strHiveName <> "" Then HiveName = strHiveName
            If strKeyName <> "" Then KeyName = strKeyName
            If strValueName <> "" Then ValueName = strValueName
            s = HiveName & "\" & KeyName
            If IsKey(HiveName, KeyName, ValueName) Then
                Return ReadKey(HiveName, KeyName, ValueName, "")
            Else
                Return ""
            End If

        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function ReadKey( _
        Optional ByVal strHiveName As String = "", _
        Optional ByVal strKeyName As String = "", _
        Optional ByVal strValueName As String = "", _
        Optional ByVal strDefault As String = "") As String
        Dim s$
        If strHiveName <> "" Then HiveName = strHiveName
        If strKeyName <> "" Then KeyName = strKeyName
        If strValueName <> "" Then ValueName = strValueName
        Try
            s = HiveName & "\" & KeyName
            s = My.Computer.Registry.GetValue(s, ValueName, strDefault)
            Return s

        Catch ex As Exception
            Return strDefault
        End Try
    End Function

    Public Function SetKey( _
        ByVal strValue As String, _
        Optional ByVal strHiveName As String = "", _
        Optional ByVal strKeyName As String = "", _
        Optional ByVal strValueName As String = "") As Boolean
        If strHiveName <> "" Then HiveName = strHiveName
        If strKeyName <> "" Then KeyName = strKeyName
        If strValueName <> "" Then ValueName = strValueName
        Dim s$ = ""
        Try
            s = HiveName & "\" & KeyName
            If Not IsKey(HiveName, KeyName, ValueName, "") Then
                My.Computer.Registry.GetValue(s, ValueName, "")
            End If
            My.Computer.Registry.SetValue(s, ValueName, strValue)
            Return True

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & "Admin rights are required.", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Registry.Set.Key")
            Return False
        End Try
    End Function

    Public Function IsKey( _
        Optional ByVal strHiveName As String = "", _
        Optional ByVal strKeyName As String = "", _
        Optional ByVal strValueName As String = "", _
        Optional ByVal strDefault As String = "") As Boolean
        If strHiveName <> "" Then HiveName = strHiveName
        If strKeyName <> "" Then KeyName = strKeyName
        If strValueName <> "" Then ValueName = strValueName
        Dim s$ = ""
        s = HiveName & "\" & KeyName
        Try
            s = My.Computer.Registry.GetValue(s, ValueName, strDefault)
            Return (s <> "")

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub New( _
        ByVal strHive As String, ByVal strKeyName As String, _
        ByVal strValueName As String, ByVal strDefault As String)
        HiveName = strHive
        KeyName = strKeyName
        ValueName = strValueName
        ValueDefault = strDefault
    End Sub
End Class
