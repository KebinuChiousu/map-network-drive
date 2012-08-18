Imports aejw
Imports System.Security
Imports System.IO
Imports System.Collections.Specialized

Public Class frmConfigure

    Private Sub frmConfigure_Load( _
                                   ByVal sender As Object, _
                                   ByVal e As System.EventArgs _
                                 ) Handles Me.Load

        Dim a As Array

        a = Share.ToArray

        cmbShares.Items.Clear()
        cmbShares.Items.AddRange(a)


    End Sub

    Private Sub btnExit_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles btnExit.Click

        Dim val As Object

        If btnApply.Enabled = True Then
            val = MsgBox("Save Changes?", vbYesNo)

            If val = vbYes Then
                UpdateIniFile()
            End If

        End If

        Application.Exit()

    End Sub

    Private Sub btnCheck_Click( _
                                ByVal sender As System.Object, _
                                ByVal e As System.EventArgs _
                              ) Handles btnCheck.Click

        CheckShare(True)

    End Sub

    Private Sub btnAdd_Click( _
                              ByVal sender As System.Object, _
                              ByVal e As System.EventArgs _
                            ) Handles btnAdd.Click

        Dim SS As New clsSecureString

        Dim Drive As String = txtDrive.Text
        Dim Path As String = txtPath.Text
        Dim User As String = txtUsername.Text
        Dim Pass As SecureString = SS.ToSecureString(txtPassword.Text)

        Dim key As String = Drive & " - " & Path

        If modShare.Share.Contains(key) Then
            MsgBox("Share already exists aborting")
            Exit Sub
        End If

        If CheckShare() Then
            key = AddShare( _
                            Drive, _
                            Path, _
                            User, _
                            Pass _
                          )
            cmbShares.Items.Add(key)
            cmbShares.Text = ""
            cmbShares.SelectedText = key
            btnApply.Enabled = True
        End If

        SS = Nothing

    End Sub

    Private Sub btnRemove_Click( _
                                 ByVal sender As System.Object, _
                                 ByVal e As System.EventArgs _
                               ) Handles btnRemove.Click

        Dim key As String = cmbShares.Text

        RemoveShare(key)
        cmbShares.Items.Remove(key)

        txtDrive.Text = ""
        txtPath.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""

        cmbShares.Text = ""

        btnApply.Enabled = True

    End Sub

    Private Sub btnUpdate_Click( _
                                 ByVal sender As System.Object, _
                                 ByVal e As System.EventArgs _
                               ) Handles btnUpdate.Click

        Dim SS As New clsSecureString

        Dim Drive As String = txtDrive.Text
        Dim Path As String = txtPath.Text
        Dim User As String = txtUsername.Text
        Dim Pass As SecureString = SS.ToSecureString(txtPassword.Text)

        Dim key As String = cmbShares.Text

        If CheckShare() Then
            RemoveShare(key)
            AddShare( _
                      Drive, _
                      Path, _
                      User, _
                      Pass _
                    )
            btnApply.Enabled = True
            MsgBox("Share Information Updated Successfully")
        End If

        SS = Nothing

    End Sub

    Private Sub cmbShares_SelectedIndexChanged( _
                                                ByVal sender As System.Object, _
                                                ByVal e As System.EventArgs _
                                              ) Handles cmbShares.SelectedIndexChanged

        txtDrive.Text = ""
        txtPath.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""

        Dim SS As New clsSecureString

        Dim key As String
        Dim ns As NetworkShare
        Dim cmb As ComboBox = CType(sender, ComboBox)

        key = cmb.Text
        ns = ShareInfo.Item(key)

        txtDrive.Text = ns.LocalDrive
        txtPath.Text = ns.NetworkPath
        txtUsername.Text = ns.Username
        txtPassword.Text = SS.ToInsecureString(ns.Password)

        SS = Nothing


    End Sub

    Private Sub btnApply_Click( _
                                ByVal sender As System.Object, _
                                ByVal e As System.EventArgs _
                              ) Handles btnApply.Click

        UpdateIniFile()
        btnApply.Enabled = False

    End Sub

    Function CheckShare(Optional ByVal Prompt As Boolean = False) As Boolean

        Dim cls As New cNetworkDrive

        Dim val As Object

        Dim Drive As String = txtDrive.Text
        Dim Path As String = txtPath.Text
        Dim User As String = txtUsername.Text
        Dim Pass As String = txtPassword.Text

        If Drive = "" Then
            MsgBox("Drive is Required (ex C:)", vbCritical)
            Return False
        End If

        If Path = "" Then
            MsgBox("Path is Required (ex \\computer\share", vbCritical)
            Return False
        End If

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
                MsgBox(ex.Message.ToString, vbCritical)
                Return False
            End Try
        Else
            If Pass = "" Then
                MsgBox("Password is required", vbCritical)
                Return False
            Else
                Try
                    cls.MapDrive(User, Pass)
                Catch ex As Exception
                    MsgBox(ex.Message.ToString, vbCritical)
                    Return False
                End Try
            End If
        End If

        If cNetworkDrive.IsNetworkDrive(Drive) Then

            If Prompt Then
                val = MsgBox( _
                              "Drive Mapping Succeeded" & vbCrLf & _
                              "Disconnect Drive?", _
                              MsgBoxStyle.YesNo _
                            )
            Else
                val = vbYes
            End If

            If val = vbYes Then
                cls.UnMapDrive(Drive, True)
            End If
        End If

        Return True

    End Function

#Region "config.ini Manipulation Routines"

    Sub UpdateIniFile()

        Dim key As String
        Dim ns As NetworkShare
        Dim idx As Integer

        Dim SS As New clsSecureString
        Dim pass As String

        ifr = Nothing
        Kill(Application.StartupPath & "\config.ini")
        InitIniFile()

        For idx = 0 To Share.Count - 1
            key = Share(idx)
            ns = ShareInfo.Item(key)

            ifr.SetIniValue( _
                             key, _
                             "Local Drive", _
                             ns.LocalDrive _
                           )

            ifr.SetIniValue( _
                             key, _
                             "Network Path", _
                             ns.NetworkPath _
                           )

            ifr.SetIniValue( _
                             key, _
                             "Username", _
                             ns.Username _
                           )

            pass = SS.EncryptString(ns.Password)
            pass = Replace(pass, "=", "|")

            ifr.SetIniValue( _
                             key, _
                             "Password", _
                             pass _
                           )

        Next


        ifr.OutputFilename = Application.StartupPath & "\config.ini"
        ifr.SaveAsIniFile()

    End Sub

    
#End Region

End Class
