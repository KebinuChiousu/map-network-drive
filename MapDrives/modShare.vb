Imports System.Security

Public Module modShare

    Structure NetworkShare
        Dim LocalDrive As String
        Dim NetworkPath As String
        Dim Username As String
        Dim Password As SecureString
    End Structure

    Public Share As New List(Of String)
    Public ShareInfo As New Dictionary(Of String, NetworkShare)

    Function AddShare( _
                       ByRef LocalDrive As String, _
                       ByRef NetworkPath As String, _
                       ByRef Username As String, _
                       ByRef Password As SecureString _
                     ) As String

        Dim key As String = LocalDrive & " - " & NetworkPath
        Dim ns As New NetworkShare

        ns.LocalDrive = LocalDrive
        ns.NetworkPath = NetworkPath
        ns.Username = Username
        ns.Password = Password

        Share.Add(key)
        ShareInfo.Add(key, ns)

        Return key

    End Function

    Sub RemoveShare(ByRef key As String)

        Share.Remove(key)
        ShareInfo.Remove(key)

    End Sub
   

End Module
