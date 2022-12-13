Imports System.IO

Public Class form2


    Private Sub form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer = 0
        Dim KeyVals() As String = {strSDVol, strJPGDest, strSDVidDest, strSDVidExt, strVideoSrc, strVideoExt, strVideoDest}
        Dim textbox As New List(Of TextBox)
        textbox.Add(TextBox1)
        textbox.Add(TextBox2)
        textbox.Add(TextBox3)
        textbox.Add(TextBox4)
        textbox.Add(TextBox5)
        textbox.Add(TextBox6)
        textbox.Add(TextBox7)

        For Each Value As String In KeyVals
            textbox(i).Text = Value
            i += 1
        Next

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = SystemColors.MenuHighlight
        Label2.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        'Pick SD Card volumelabel (source drive)
        'Folder Browser
        Dim prompt As String = "Please set the Volume Label for this drive."
        Dim title As String = "Volume Label"
        Dim defaultResponse As String = String.Empty
        Dim answer As Object
        Dim volLabel As System.IO.DriveInfo
        Dim folderDlg As New FolderBrowserDialog
        folderDlg.Description = "Choose your SD Card Drive."
        folderDlg.ShowNewFolderButton = False
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            If folderDlg.SelectedPath.Length = 3 Then
                Dim root As Environment.SpecialFolder = folderDlg.RootFolder
                volLabel = My.Computer.FileSystem.GetDriveInfo(folderDlg.SelectedPath)
                If volLabel.VolumeLabel = "" Then
                    Me.SetTopLevel(False)
                    answer = InputBox(prompt, title, defaultResponse)
                    volLabel.VolumeLabel = answer
                End If
                Me.SetTopLevel(True)
                TextBox1.Text = volLabel.VolumeLabel
                'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_SD_Volume", volLabel.VolumeLabel)
                'MsgBox("Value has been updated.")
            Else
                MsgBox(folderDlg.SelectedPath & " is not a valid selection for this field." & vbCrLf & "Please choose your camera card drive letter.")
            End If
        End If

    End Sub

    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.WhiteSmoke
        Label2.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        'Pick JPG destination folder
        'Folder Browser

        'Dim volLabel As System.IO.DriveInfo
        Dim folderDlg As New FolderBrowserDialog
        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
            TextBox2.Text = folderDlg.SelectedPath
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_jpg_Destination", folderDlg.SelectedPath)
            'MsgBox("Value has been updated.")
        End If
    End Sub

    Private Sub Label3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseEnter
        Label3.BackColor = SystemColors.MenuHighlight
        Label3.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave
        Label3.BackColor = Color.WhiteSmoke
        Label3.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        'Pick SD Card Video Destination
        'Folder Browser

        'Dim volLabel As System.IO.DriveInfo
        Dim folderDlg As New FolderBrowserDialog

        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
            TextBox3.Text = folderDlg.SelectedPath
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_vid_Destination", folderDlg.SelectedPath)
            'MsgBox("Value has been updated.")
        End If
    End Sub

    Private Sub Label4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label4.MouseEnter
        Label4.BackColor = SystemColors.MenuHighlight
        Label4.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label4.MouseLeave
        Label4.BackColor = Color.WhiteSmoke
        Label4.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        'Pick SD Card Video File Extension.  (.MOV), (.MTS), (.MP4), ETC
        'File Browser

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Choose a Video File From Your Camera, Or That Came From The Camera."
        fd.InitialDirectory = "C:\"
        fd.Filter = "Video Files (*.mov, *.avi, *.mp4, *.m4v, *.mts, *.mkv, *.mxf)|*.mov;*.avi;*.mp4;*.m4v;*.mts;*.mkv;*.mxf|All Files (*.*)|*.*"
        fd.FilterIndex = 1
        fd.Multiselect = False
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Dim strExt As String = strFileName.Substring(strFileName.Length - 4)
            TextBox4.Text = ("*" & strExt)
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_Vid_Type", strExt)
            'MsgBox("Value has been updated.")
        End If

    End Sub

    Private Sub Label5_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label5.MouseEnter
        Label5.BackColor = SystemColors.MenuHighlight
        Label5.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label5.MouseLeave
        Label5.BackColor = Color.WhiteSmoke
        Label5.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        'Pick Video Camera Source Folder (Folder on Camera where files are)
        'Folder Browser

        Dim folderDlg As New FolderBrowserDialog
        folderDlg.Description = "Choose the folder on the video camera with the video files."
        folderDlg.ShowNewFolderButton = False
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
            Dim strPath As String = folderDlg.SelectedPath
            strPath = strPath.Substring(3)

            TextBox5.Text = strPath
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_vid_Source", strPath)
            'MsgBox("Value has been updated.")
        
        End If
    End Sub

    Private Sub Label6_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label6.MouseEnter
        Label6.BackColor = SystemColors.MenuHighlight
        Label6.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label6.MouseLeave
        Label6.BackColor = Color.WhiteSmoke
        Label6.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        'Pick Video camera File Extension.  (.MOV), (.MTS), (.MP4), ETC
        'File Browser

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Choose a Video File From Your Camera, Or That Came From The Camera."
        fd.InitialDirectory = "C:\"
        fd.Filter = "Video Files (*.mov, *.avi, *.mp4, *.m4v, *.mts, *.mkv, *.mxf)|*.mov;*.avi;*.mp4;*.m4v;*.mts;*.mkv;*.mxf|All Files (*.*)|*.*"
        fd.FilterIndex = 1
        fd.Multiselect = False
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Dim strExt As String = strFileName.Substring(strFileName.Length - 4)
            TextBox6.Text = ("*" & strExt)
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_Type", strExt)
            'MsgBox("Value has been updated.")
        End If

    End Sub

    Private Sub Label7_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label7.MouseEnter
        Label7.BackColor = SystemColors.MenuHighlight
        Label7.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label7.MouseLeave
        Label7.BackColor = Color.WhiteSmoke
        Label7.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        'Pick Video Camera Files Destination
        'Folder Browser

        'Dim volLabel As System.IO.DriveInfo
        Dim folderDlg As New FolderBrowserDialog

        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder
            TextBox7.Text = folderDlg.SelectedPath
            'My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_Destination", folderDlg.SelectedPath)
            'MsgBox("Value has been updated.")
        End If
    End Sub

    Private Sub Label8_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.MouseEnter
        Label8.BackColor = SystemColors.MenuHighlight
        Label8.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.MouseLeave
        Label8.BackColor = Color.WhiteSmoke
        Label8.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        Dim i As Integer = 0
        'Dim KeyVals() As String = {strSDVol, strJPGDest, strSDVidDest, strSDVidExt, strVideoSrc, strVideoExt, strVideoDest}
        Dim Keys() As String = {"Camera_SD_Volume", "Camera_jpg_Destination", "Camera_vid_Destination", "Camera_Vid_Type", "Video_vid_Source", "Video_Type", "Video_Destination"}
        Dim textbox As New List(Of TextBox)
        textbox.Add(TextBox1)
        textbox.Add(TextBox2)
        textbox.Add(TextBox3)
        textbox.Add(TextBox4)
        textbox.Add(TextBox5)
        textbox.Add(TextBox6)
        textbox.Add(TextBox7)


        For Each Value As String In Keys
            'Update all the settings here.
            On Error Resume Next
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, textbox(i).Text)
            i += 1
        Next

        check_settings()
        MsgBox("All settings have been updated.")
        Me.Close()
        Form1.Focus()
    End Sub

    Private Sub Label10_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label10.MouseEnter
        Label10.BackColor = SystemColors.MenuHighlight
        Label10.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label10_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label10.MouseLeave
        Label10.BackColor = Color.WhiteSmoke
        Label10.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        Me.Close()
        Form1.Focus()
    End Sub

    Private Sub Label11_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label11.MouseEnter
        Label11.BackColor = SystemColors.MenuHighlight
        Label11.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label11_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label11.MouseLeave
        Label11.BackColor = Color.WhiteSmoke
        Label11.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class