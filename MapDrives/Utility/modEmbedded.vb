Imports System.Reflection
Imports System.Reflection.Assembly
Imports System.IO

Module modEmbedded

    Public Function EmbeddedObj(ByVal Name As String) As Stream

        Dim assy As Assembly

        Dim obj As Stream

        Dim str As String = ""

        assy = GetExecutingAssembly()
        Dim resources() As String

        resources = assy.GetManifestResourceNames()

        For Each resourceName As String In resources
            If InStr(resourceName, Name) <> 0 Then
                str = resourceName
                Exit For
            End If
        Next

        obj = assy.GetManifestResourceStream(str)

        Return obj

    End Function

    Sub GetEmbeddedFile(ByVal filename As String)

        Dim UMS As UnmanagedMemoryStream
        Dim outfile As Stream
        Const sz As Integer = 4096
        Dim buf As Byte()
        Dim nRead As Integer

        ReDim buf(sz)

        UMS = EmbeddedObj(filename)

        File.Delete(filename)
        outfile = File.Create(filename)

        While True
            nRead = UMS.Read(buf, 0, sz)
            If nRead < 1 Then
                Exit While
            End If
            outfile.Write(buf, 0, nRead)
        End While

        outfile.Close()


    End Sub

End Module
