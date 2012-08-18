Imports aejw
Imports aejw.cNetworkDrive
Imports System.IO
Imports System.Collections.Specialized
Imports System.Security
Imports ErrorsAndEvents

Module modMain

    Public ifr As IniFileReader
    Public el As New ErrorLogger

    Sub Main()

        Dim Path As String

        Path = My.Application.Info.DirectoryPath

        If Not File.Exists(Path & "\" & "XMLtoINI.xslt") Then
            GetEmbeddedFile("XMLtoINI.xslt")
        End If

        InitIniFile()

        Dim arrPar() As String = Command.Split(",")
        Dim zForm As Windows.Forms.Form

        If arrPar(0) = "-config" Then
            zForm = New frmConfigure
            zForm.ShowDialog()
        Else
            MapDrives()
        End If

    End Sub

#Region "config.ini Manipulation Routines"

    Sub InitIniFile()

        Dim fi As FileInfo
        Dim sc As StringCollection

        fi = New FileInfo(Application.StartupPath & "\\" & "config.ini")

        If fi.Exists Then
            ifr = New IniFileReader(Application.StartupPath & "\config.ini", True)
            LoadIniFile()
        Else
            ifr = New IniFileReader(Application.StartupPath & "\config.ini", True)
            sc = ifr.GetIniComments(Nothing)
            sc.Add("Network Share Mapping Configuration File")
            ifr.SetIniComments(Nothing, sc)
            ifr.OutputFilename = Application.StartupPath & "\config.ini"
            ifr.SaveAsIniFile()
        End If

    End Sub

    Sub LoadIniFile()

        Dim SS As New clsSecureString

        Dim scSection As StringCollection
        Dim scKey As StringCollection

        Dim Section As String
        Dim Key As String
        Dim Value As String

        Dim idxSection As Integer
        Dim idxKey As Integer

        Dim Drive As String = ""
        Dim Path As String = ""
        Dim Username As String = ""
        Dim Password As New SecureString

        scSection = ifr.AllSections

        For idxSection = 0 To scSection.Count - 1
            Section = scSection(idxSection)
            scKey = ifr.AllKeysInSection(Section)

            For idxKey = 0 To scKey.Count - 1
                Key = scKey(idxKey)

                Value = ifr.GetIniValue(Section, Key)

                Select Case Key
                    Case "Local Drive"
                        Drive = Value
                    Case "Network Path"
                        Path = Value
                    Case "Username"
                        Username = Value
                    Case "Password"
                        Value = Replace(Value, "|", "=")
                        Password = SS.DecryptString(Value)

                End Select

            Next

            AddShare( _
                      Drive, _
                      Path, _
                      Username, _
                      Password _
                    )


        Next

        SS = Nothing

    End Sub

#End Region

    Sub MapDrives()

        Dim idx As Integer
        Dim key As String
        Dim ns As NetworkShare

        For idx = 0 To Share.Count - 1

            key = Share(idx)
            ns = ShareInfo.Item(key)

            MapDrive( _
                      ns.LocalDrive, _
                      ns.NetworkPath, _
                      ns.Username, _
                      ns.Password _
                    )

        Next


    End Sub

    Sub MapDrive( _
                  ByRef Drive As String, _
                  ByRef Path As String, _
                  ByRef User As String, _
                  ByRef Password As SecureString _
                )


        Dim cls As New cNetworkDrive
        Dim SS As New clsSecureString

        Dim Pass As String

        Pass = SS.ToInsecureString(Password)

        SS = Nothing

        cls.LocalDrive = Drive
        cls.ShareName = Path
        cls.Force = True
        cls.Persistent = False
        cls.PromptForCredentials = False
        cls.SaveCredentials = False

        If User = "" Then
            Try
                cls.MapDrive()
            Catch ex As Exception
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Map Drive")
                Exit Sub
            End Try
        Else

            Try
                cls.MapDrive(User, Pass)
            Catch ex As Exception
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Map Drive")
                Exit Sub
            End Try

        End If

    End Sub

End Module
