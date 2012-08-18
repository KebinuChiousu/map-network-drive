Imports System
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Encoding
Imports System.Runtime.InteropServices
Imports System.Reflection.Assembly
Imports System.Reflection

Class clsSecureString

    Public entropy As Byte()

    Sub New()
        entropy = GetEntropy()
    End Sub

    Private Function GetEntropy() As Byte()

        Dim assy As System.Reflection.Assembly

        Dim temp As String

        assy = GetExecutingAssembly()

        temp = assy.GetCustomAttributes(GetType(GuidAttribute), False)(0).value

        Return Encoding.Unicode.GetBytes(temp)

    End Function

    Public Function EncryptString( _
                                    ByVal input As SecureString _
                                  ) As String

        Dim encryptedData As Byte()
        Dim temp As String


        encryptedData = ProtectedData.Protect( _
                        Unicode.GetBytes( _
                                          ToInsecureString(input)), _
                                          entropy, _
                                          DataProtectionScope.CurrentUser _
                                        )

        temp = System.Convert.ToBase64String(encryptedData)

        Return temp

    End Function

    Public Function DecryptString(ByVal encryptedData As String) As SecureString

        Try
            Dim decryptedData As Byte()
            Dim temp As Byte()
            Dim ss As SecureString

            temp = System.Convert.FromBase64String(encryptedData)

            decryptedData = ProtectedData. _
                            Unprotect( _
                                       temp, _
                                       entropy, _
                                       DataProtectionScope.CurrentUser _
                                     )

            ss = ToSecureString(Unicode.GetString(decryptedData))

            Return ss
        Catch ex As Exception
            el.WriteToErrorLog( _
                                ex.Message, _
                                ex.StackTrace, _
                                "Map Drives" _
                              )

            Return New SecureString()
        End Try

    End Function

    Public Function ToSecureString(ByVal input As String) As SecureString

        Dim secure As New SecureString()

        For Each c As Char In input
            secure.AppendChar(c)
        Next

        secure.MakeReadOnly()

        Return secure

    End Function

    Public Function ToInsecureString(ByVal input As SecureString) As String

        Dim returnValue As String = String.Empty
        Dim ptr As IntPtr = Marshal.SecureStringToBSTR(input)

        Try
            returnValue = Marshal.PtrToStringBSTR(ptr)
        Finally
            Marshal.ZeroFreeBSTR(ptr)
        End Try

        Return returnValue

    End Function


End Class
