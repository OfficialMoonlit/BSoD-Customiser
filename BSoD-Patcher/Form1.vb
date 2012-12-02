Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Class frmMain
    Dim strPad As String

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        lblLength.Text = Len(TextBox2.Text)
    End Sub

    Private Sub txtBlock1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock1.TextChanged
        'strPad = Strings.Left(strPad.PadRight(94), 94)
        'txtBlock1.Text = strPad
        strPad = txtBlock1.Text
        lblBlock1.Text = "(" & Len(Strings.Trim(strPad)) & "/96)"
    End Sub

    Private Sub txtBlock2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock2.TextChanged
        strPad = txtBlock2.Text
        lblBlock2.Text = "(" & Len(Strings.Trim(strPad)) & "/139)"
    End Sub

    Private Sub txtBlock3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock3.TextChanged
        strPad = txtBlock3.Text
        lblBlock3.Text = "(" & Len(Strings.Trim(strPad)) & "/491)"
    End Sub

    Private Sub txtBlock5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock5.TextChanged
        strPad = txtBlock5.Text
        lblBlock5.Text = "(" & Len(Strings.Trim(strPad)) & "/23)"
    End Sub

    Private Sub txtBlock7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock7.TextChanged
        strPad = txtBlock7.Text
        lblBlock7.Text = "(" & Len(Strings.Trim(strPad)) & "/34)"
    End Sub

    Private Sub txtBlock9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBlock9.TextChanged
        strPad = txtBlock9.Text
        lblBlock9.Text = "(" & Len(Strings.Trim(strPad)) & "/126)"
    End Sub

    Private Sub cmdPatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPatch.Click
        Dim x As New MsgBoxResult

        If txtSource.Text = txtDest.Text Then
            MsgBox("Source and destination file names are identical." & vbCrLf & "Please choose another destination file name or location.", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
            Exit Sub
        ElseIf txtSource.Text = "" Then
            MsgBox("Invalid source file, please choose another file name or location.", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
            Exit Sub
        ElseIf txtDest.Text = "" Then
            MsgBox("Invalid destination file, please choose another file name or location.", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
            Exit Sub
        End If

        Try
            Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
            Dim f As FileStream = New FileStream(txtSource.Text, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            f = New FileStream(txtSource.Text, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            md5.ComputeHash(f)
            Dim hash As Byte() = md5.Hash
            Dim buff As StringBuilder = New StringBuilder
            Dim hashByte As Byte
            For Each hashByte In hash
                buff.Append(String.Format("{0:X2}", hashByte))
            Next

            If buff.ToString() = "CE218BC7088681FAA06633E218596CA7" Then
                x = MsgBox("File hash: " & vbCrLf & buff.ToString() & vbCrLf & vbCrLf & "Windows XP Service Pack 2 kernel. Patch now?", vbOKCancel + MsgBoxStyle.Information, "BSoD-Patcher")
                If x = MsgBoxResult.Cancel Then
                    MsgBox("Patch aborted!", vbOKOnly + MsgBoxStyle.Information, "BSoD-Patcher")
                    Exit Sub
                End If
                Try
                    File.Copy(txtSource.Text, txtDest.Text, True)
                    Dim encText As New System.Text.UTF8Encoding()
                    Dim btText() As Byte
                    Dim bw As New IO.BinaryWriter(IO.File.Open(txtDest.Text, IO.FileMode.Open, IO.FileAccess.ReadWrite))
                    bw.BaseStream.Seek(&H1F8B1C, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock1.Text.PadRight(96), 96))
                    bw.Write(btText)
                    bw.BaseStream.Seek(&H1F8BC4, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock2.Text.PadRight(139), 139))
                    bw.Write(btText)
                    bw.BaseStream.Seek(&H1F8C58, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock3.Text.PadRight(491), 491))
                    bw.Write(btText)
                    bw.BaseStream.Seek(&H1F8E4C, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock5.Text.PadRight(23), 23))
                    bw.Write(btText)
                    bw.BaseStream.Seek(&H1F8F00, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock7.Text.PadRight(34), 34))
                    bw.Write(btText)
                    bw.BaseStream.Seek(&H1F8F50, IO.SeekOrigin.Begin)
                    btText = encText.GetBytes(Strings.Left(txtBlock9.Text.PadRight(126), 126))
                    bw.Write(btText)
                    If chkStop.CheckState = CheckState.Unchecked Then
                        bw.BaseStream.Seek(&H5C92E, IO.SeekOrigin.Begin)
                        btText = encText.GetBytes("                                          ")
                        bw.Write(btText)
                    End If
                    bw.Close()
                    strPad = vbNullString
                    MsgBox("Patched!", vbOKOnly + MsgBoxStyle.Information, "BSoD-Patcher")
                Catch ex As Exception
                    MsgBox("Patch failed!", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
                End Try
            Else
                MsgBox("File hash: " & vbCrLf & buff.ToString() & vbCrLf & vbCrLf & "Unrecognised kernel! This file cannot be patched.", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
                'MsgBox("File hash: " & vbCrLf & buff.ToString() & vbCrLf & vbCrLf & "Unrecognised kernel! It is not recommended that you patch this file." & vbCrLf & "You may patch at your own risk. Patch now?", MsgBoxStyle.OkCancel + MsgBoxStyle.Exclamation, "BSoD-Patcher")
            End If

        Catch ex As Exception
            MsgBox("File failed hash check!", vbOKOnly + MsgBoxStyle.Exclamation, "BSoD-Patcher")
        End Try

    End Sub





    Private Sub cmdSourceBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceBrowse.Click
        Dim fileOpen As New OpenFileDialog
        fileOpen.InitialDirectory = Application.ExecutablePath
        fileOpen.Filter = "Executable Files (*.exe)|*.exe"
        fileOpen.AddExtension = True
        fileOpen.DefaultExt = ".exe"
        If fileOpen.ShowDialog() = Windows.Forms.DialogResult.OK Then txtSource.Text = fileOpen.FileName
    End Sub

    Private Sub cmdPatchedBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPatchedBrowse.Click
        Dim fileSave As New SaveFileDialog
        fileSave.InitialDirectory = Application.ExecutablePath
        fileSave.Filter = "Executable Files (*.exe)|*.exe"
        fileSave.AddExtension = True
        fileSave.DefaultExt = ".exe"
        If fileSave.ShowDialog() = Windows.Forms.DialogResult.OK Then txtDest.Text = fileSave.FileName
    End Sub

 
    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        MsgBox("BSoD-Patcher by Moonlit, v1.2.2 / March 2011" & vbCrLf & vbCrLf & "Patches a Windows kernel to create custom BSoD messages." & vbCrLf & "(Windows XP Service Pack 2 only.)" & vbCrLf & vbCrLf & "Neither the creator or distributors of this application are responsible for its (mis)use, and no guarantee is made of its fitness for any purpose." & vbCrLf & "Use at your own risk, and always make backups.", MsgBoxStyle.Question + MsgBoxStyle.OkOnly, "About...")
    End Sub
End Class

