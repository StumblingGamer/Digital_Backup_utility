Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.IO
Imports System
Imports System.Collections
Imports Microsoft.VisualBasic.FileIO


Public Class Form1
    Public filesinfolder As Integer
    '
    'The messages to look for.
    Private Const DBT_CONFIGCHANGED As Integer = &H18
    Private Const WM_DEVICECHANGE As Integer = &H219
    Private Const DBT_DEVICEARRIVAL As Integer = &H8000
    Private Const DBT_DEVICEREMOVECOMPLETE As Integer = &H8004
    Private Const DBT_DEVTYP_VOLUME As Integer = &H2  '
    '
    'Get the information about the detected volume.
    Private Structure DEV_BROADCAST_VOLUME

        Dim Dbcv_Size As Integer
        Dim Dbcv_Devicetype As Integer
        Dim Dbcv_Reserved As Integer
        Dim Dbcv_Unitmask As Integer
        Dim Dbcv_Flags As Short

    End Structure
    Function doesExisttest(ByVal strToTest As String, ByVal filedate As String, ByVal filesize As String)  '(filename from camera card, creationdate of said file, filesize of said file)
        'Input looks like this:
        'Item = Filesize:Filedate
        'Textbox1 = Filedate
        'Textbox2 = Filesize
        '
        'doesExisttest(filenametocheck)
        '
        'these are set to false, and will be set to true during the function if a condition is met.
        Dim strDate As Boolean = False
        Dim strSize As Boolean = False
        Dim filename As String = ""
        Dim retrun As String = ""
        Dim strFileDate As String = ""
        Dim strFileSize As String = ""

        'this checks each item in the list (files found at the destination) against a single file in question (the current file trying to get copied from the camera).
        ''For Each item As String In ListBox1.Items
        For Each item As String In My.Computer.FileSystem.GetFiles(strVideoDest, FileIO.SearchOption.SearchAllSubDirectories, strVideoExt)  'change this to all files
            Dim strFile As FileInfo = New FileInfo(item)
            MsgBox(item)
            strFileDate = strFile.CreationTime
            strFileSize = strFile.Length

            'if filedate is the same
            ''If Split(item, ":", 2)(1) = filedate Then
            If strFileDate = filedate Then
                MsgBox("filedates are the same")
                strDate = True

                'if filesize is the same
                ''If Split(item, ":", 2)(0) = filesize Then
                If strFileSize = filesize Then
                    MsgBox("file size is the same")
                    'These will only be true if both conditions are true
                    strSize = True
                    'strDate = True
                    'We found a existing file, exit the loop
                    Exit For
                ElseIf strFileSize <> filesize Then
                    'We didn't find an existing file, keep looping
                    strSize = False
                    'filename exists at destination, has same file date, but different file size (get filename of existing file to overwrite)
                    filename = item
                End If
                'if the dates matched, and it has checked the filesize, exit the loop
                Exit For
            Else
                'if the date doesn't match, keep looping
                strDate = False
            End If
        Next

        'Set the return value here based on the output of the logic.  (this happens once the loop is done, or exited out of)
        '
        'This line says: Date matches, filesize doesn't.  Get filename from destination file, return it as the name of file to be written
        If strDate = True And strSize = False Then
            'Label3.Text = "Exists, but filelength is too short: OVERWITE file at destination"
            Return filename & ":true"
            'Exit Sub
        End If

        If strDate = True And strSize = True Then
            'Label3.Text = "Exists, and is identical: Don't copy."
            Return ":false"
            'Exit Sub
        End If

        If strDate = False And strSize = False Then
            'Label3.Text = "Doesn't Exist, copy as normal."
            Return ":true"
        End If

    End Function

    Function doesExist(ByVal strToTest As String, ByVal filedate As String, ByVal filesize As String)  '(filename from camera card, creationdate of said file, filesize of said file)
        'Input looks like this:
        'Item = Filesize:Filedate
        'Textbox1 = Filedate
        'Textbox2 = Filesize
        '
        'doesExist(filenametocheck)
        '
        'these are set to false, and will be set to true during the function if a condition is met.
        Dim strDate As Boolean = False
        Dim strSize As Boolean = False
        Dim filename As String = Split(Path.GetFileName(strToTest), ".", 2)(0)
        Dim retrun As String = ""
        Dim strFileDate As String = ""
        Dim strFileSize As String = ""

        filedate = Replace(filedate, "/", "")
        filedate = Replace(filedate, ":", "")
        filedate = Replace(filedate, "AM", "")
        filedate = Replace(filedate, "PM", "")
        filedate = Replace(filedate, " ", "")
        MsgBox(filedate)


        MsgBox(Split(Path.GetFileName(strToTest), ".", 2)(0))


        'this checks each item in the list (files found at the destination) against a single file in question (the current file trying to get copied from the camera).
        ''For Each item As String In ListBox1.Items
        For Each item As String In My.Computer.FileSystem.GetFiles(strVideoDest, FileIO.SearchOption.SearchAllSubDirectories, strVideoExt)  'change this to all files
            Dim strFile As FileInfo = New FileInfo(item)
            'MsgBox(item)
            strFileDate = strFile.CreationTime
            strFileSize = strFile.Length

            'if filedate is the same
            ''If Split(item, ":", 2)(1) = filedate Then
            If strFileDate = filedate Then
                strDate = True

                'if filesize is the same
                ''If Split(item, ":", 2)(0) = filesize Then
                If strFileSize = filesize Then
                    'These will only be true if both conditions are true
                    strSize = True
                    'strDate = True
                    'We found a existing file, exit the loop
                    Exit For
                ElseIf strFileSize <> filesize Then
                    'We didn't find an existing file, keep looping
                    strSize = False
                    'filename exists at destination, has same file date, but different file size (get filename of existing file to overwrite)
                    'filename = item
                End If
                'if the dates matched, and it has checked the filesize, exit the loop
                Exit For
            Else
                'if the date doesn't match, keep looping
                strDate = False
            End If
        Next

        'Set the return value here based on the output of the logic.  (this happens once the loop is done, or exited out of)
        '
        'This line says: Date matches, filesize doesn't.  Get filename from destination file, return it as the name of file to be written
        If strDate = True And strSize = False Then
            MsgBox(filename + ":true")
            'Label3.Text = "Exists, but filelength is too short: OVERWITE file at destination"
            Return filename & ":true"
            'Exit Sub
        End If

        If strDate = True And strSize = True Then
            'Label3.Text = "Exists, and is identical: Don't copy."
            Return ":false"
            'Exit Sub
        End If

        If strDate = False And strSize = False Then
            'Label3.Text = "Doesn't Exist, copy as normal."
            Return ":true"
        End If

    End Function

    Function vidExsists(ByVal actualname As String, ByVal numberedname As String) As String
        Dim fiDest As FileInfo
        Dim fiSource As FileInfo
        Dim dtSrcCreated As DateTime = File.GetCreationTime(actualname) 'created date of file on video camera
        Dim output As String

        MsgBox("CALLED AGAIN")
        For Each objFoundFile As Object In My.Computer.FileSystem.GetFiles(strVideoDest, FileIO.SearchOption.SearchTopLevelOnly, strVideoExt) 'every video found at video destination folder
            Dim dtDestCreated As DateTime = File.GetCreationTime(objFoundFile) 'get created date for file found at destination
            fiDest = New FileInfo(objFoundFile) 'file at destination
            fiSource = New FileInfo(actualname) 'file from the source to copy

            'first check (is file date same)
            If dtSrcCreated = dtDestCreated And fiSource.Length = fiDest.Length Then
                MsgBox("Dates are the same, and filesize is the same." & fiDest.FullName & " | " & fiSource.FullName & "   DO NOT COPY")
                Return ":false"  'don't copy
            ElseIf dtSrcCreated = dtDestCreated And Not fiSource.Length = fiDest.Length Then
                MsgBox("Dates are the same, and filesize is NOT the same. - " & fiDest.Name & " -- " & fiSource.FullName & " OVERWRITE")
                Return fiDest.Name & ":true" 'copy - OVERWRITE
                'ElseIf Not dtSrcCreated = dtDestCreated And Not fiSource.Length = fiDest.Length Then
                'MsgBox("Nothing Matches, destination file should be: " & (numberedname + 1) & " -- " & objFoundFile.ToString & " / Copy this file: " & fiSource.FullName)
                'Return (numberedname + 1) & ":true"   'Copy Next numbered filename in line.
            End If

            'MsgBox("Nothing Matches, destination file should be: " & (numberedname + 1) & " -- " & objFoundFile.ToString & " / Copy this file: " & fiSource.FullName)
            ' Return (numberedname + 1) & ":true"   'Copy Next numbered filename in line.
        Next

        MsgBox("Nothing Matches, destination file should be: " & (numberedname + 1) & " -- Copy: " & actualname)
        Return (numberedname + 1) & ":true"   'Copy Next numbered filename in line.


        'call like so: vidExsists(sourcefile as string)
        'Dim fiDest As FileInfo
        'Dim fiSource As FileInfo
        'Dim strTemp As String = "zero"
        'Dim dtSrcCreated As DateTime = File.GetCreationTime(sourcefile) 'created date of file on video camera
        '
        'For Each objFoundFile As Object In My.Computer.FileSystem.GetFiles(strVideoDest, FileIO.SearchOption.SearchTopLevelOnly, strVideoExt) 'every video found on video
        '        Dim dtDestCreated As DateTime = File.GetCreationTime(objFoundFile) 'get created date for file found at destination
        '
        'If Not dtDestCreated = dtSrcCreated Then Return "zero"
        '
        'If dtDestCreated = dtSrcCreated Then 'if the created dates are the same then...
        '        'strTemp = "single"
        'fiDest = New FileInfo(objFoundFile) 'file at destination
        'fiSource = New FileInfo(sourcefile) 'file from the source to copy
        '
        'If CBool(Not ((fiDest.Attributes And FileAttributes.System) Or (fiDest.Attributes And FileAttributes.Hidden))) Then 'if it's able to read attributes then...
        '        If CBool(Not ((fiSource.Attributes And FileAttributes.System) Or (fiSource.Attributes And FileAttributes.Hidden))) Then
        '        If fiSource.Length = fiDest.Length Then 'if sourcefile file size is same as destination then...
        '        Return "double"
        ''MsgBox(fiSource.Length & "  Files are the same date, and size  " & fiDest.Length) 'file found on destination has same creation date and filesize (do not transfer)
        'Else 'if filesize doesn't match then...
        'Return "single"
        ''MsgBox(fiSource.Length & "  they have the same date, but different file size  " & fiDest.Length) 'files have same creation date but not filesize (transfer)
        'End If
        'End If
        'End If
        'End If
        '
        'Next

    End Function

    ' The delegate
    Delegate Sub SetLabelText_Delegate(ByVal [Label] As Label, ByVal [text] As String)

    ' The delegates subroutine.
    Private Sub SetLabelText_ThreadSafe(ByVal [Label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If [Label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[Label], [text]})
        Else
            [Label].Text = [text]
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SD, CAM, bool, drive As String
        Me.Height = 290

        Call check_settings()

        SD = isConnected(strSDVol, "vol")
        bool = SD.Substring(SD.Length - 4)
        drive = SD.Substring(0, 2)

        If String.Compare(bool, "True", True) = 0 Then
            Label4.Text = SD.Substring(0, 3)
            If Not Label1.Text.Contains("Connected") Then
                Label1.Text = Label1.Text & vbCrLf & "- SD Card is Connected -"
            End If
        End If

        CAM = isConnected("PRIVATE\AVCHD\BDMV\STREAM", "dir")
        bool = CAM.Substring(CAM.Length - 4)
        drive = CAM.Substring(0, 2)

        If String.Compare(bool, "True", True) = 0 Then
            Label5.Text = CAM.Substring(0, 2)
            If Not Label2.Text.Contains("Connected") Then
                Label2.Text = Label2.Text & vbCrLf & "- Video Camera is Connected -"
            End If
            'For Each foundFile As Object In My.Computer.FileSystem.GetFiles("I:\", FileIO.SearchOption.SearchAllSubDirectories, "*.mts")
            'ListBox1.Items.Add(foundFile)
            'Next
        End If

        Timer1.Enabled = True

    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork 'This is be used for copying Pictures from Camera only
        Dim x As Integer = 0
        Dim i As Integer = 0
        Dim q As Integer = 0
        Dim progress As Integer = 0
        Dim args As Object() = DirectCast(e.Argument, Object())
        Dim strDrive As String = CStr(args(0))                  'source directory
        Dim strExt As String = CStr(args(1))                    '"*.jpg"
        Dim strDest As String = CStr(args(2))                   'destination top path... without "\"
        Dim strDestMOV As String = CStr(args(3))                'destination if file is MOV video... without "\"
        Dim scrFileCount As String = CStr(args(4))              'source filecount (files to be copied)
        Dim month, year, first1, strDest1 As String

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Code from another project  -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING --     '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim diSourceDir As New DirectoryInfo(strDrive)

        For Each foundFile As String In My.Computer.FileSystem.GetFiles(strDrive, FileIO.SearchOption.SearchAllSubDirectories, strExt)  'change this to all files
            If BackgroundWorker1.CancellationPending = True Then
                ' Set Cancel to True
                e.Cancel = True
                Exit For
            End If

            Dim fi As FileInfo = New FileInfo(foundFile)

            If LCase(fi.Extension) = ".jpg" Then

                Dim first As DateTime = Split(fi.CreationTime, " ", 3)(0)
                first1 = first.ToString("MM/dd/yyyy")
                month = Split(first1, "/", 4)(0)
                year = Split(first1, "/", 4)(2)
                strDest1 = strDest & "\" & year & "\" & month & "-" & year                  'sets destination

                x += 1
                progress = ((x * 100) / scrFileCount)
                BackgroundWorker1.ReportProgress(progress)
                SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & x & " of " & scrFileCount & "] Files...")
                SetLabelText_ThreadSafe(Me.Label7, "Transfering: " + foundFile + Environment.NewLine + "To Destination: " + strDest1)

                If Directory.Exists(strDest1) = False Then Directory.CreateDirectory(strDest1) 'create directory if it doesn't already exist

                Dim newfile As String = strDest1 & "\" & fi.Name

                If Not File.Exists(newfile) Then                                              'if fileNAME doesn't exist at destination
                    fi.CopyTo(newfile, False)                                                 'copy file normal
                    i += 1
                ElseIf Not CompareFiles(newfile, fi.FullName) Then                            'filename is the same, byte compare = false (rename & copy)
                    Dim fullname As String = strDest1 & "\(1) " & fi.Name
                    If Not File.Exists(fullname) Then
                        fi.CopyTo(fullname, False)                                            'copy file
                        i += 1
                    Else
                        q += 1
                    End If

                Else                                                                          'exists and is byte identical (skip copy/next file)
                    q += 1

                End If
                BackgroundWorker1.ReportProgress(progress)

            ElseIf fi.Extension = ".MOV" Then   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x += 1
                progress = ((x * 100) / scrFileCount)
                BackgroundWorker1.ReportProgress(progress)
                SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & x & " of " & scrFileCount & "] Files...")
                SetLabelText_ThreadSafe(Me.Label7, "Transfering: " & foundFile & "...")

                If Directory.Exists(strDestMOV) = False Then Directory.CreateDirectory(strDestMOV) 'create directory if it doesn't already exist

                Dim newfile As String = strDestMOV & "\" & fi.Name

                If Not File.Exists(newfile) Then                                              'if fileNAME doesn't exist at destination
                    fi.CopyTo(newfile, False)                                                 'copy file normal
                    i += 1
                ElseIf Not CompareFiles(newfile, fi.FullName) Then                            'filename is the same, byte compare = false (rename & copy)
                    Dim fullname As String = strDestMOV & "\(1) " & fi.Name
                    If Not File.Exists(fullname) Then
                        fi.CopyTo(fullname, False)                                            'copy file
                        i += 1
                    Else
                        q += 1
                    End If

                Else                                                                          'exists and is byte identical (skip copy/next file)
                    q += 1

                End If

                BackgroundWorker1.ReportProgress(progress)
            End If

        Next

        BackgroundWorker1.ReportProgress(progress)
        SetLabelText_ThreadSafe(Me.Label8, i)
        SetLabelText_ThreadSafe(Me.Label9, q)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Code from another project  -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING -- WORKING --     '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Fake Copying Code  -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING --  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'For Each foundFile As Object In My.Computer.FileSystem.GetFiles(strDrive, FileIO.SearchOption.SearchAllSubDirectories, strExt)
        '        If BackgroundWorker1.CancellationPending = True Then
        '        ' Set Cancel to True
        'e.Cancel = True
        'Exit For
        'End If
        ''''''''''''''''''''''''
        'x += 1
        'SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & x & " of " & scrFileCount & "] Files...")
        'System.Threading.Thread.Sleep(50)
        'SetLabelText_ThreadSafe(Me.Label7, foundFile)
        'progress = ((x * 100) / scrFileCount)
        'BackgroundWorker1.ReportProgress(progress)
        '
        '''''''''''''''''''''''
        'Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Fake Copying Code  -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING -- TESTING --  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If e.ProgressPercentage <= 100 Then
            Me.ProgressBar1.Value = e.ProgressPercentage
            Me.ProgressBar1.Refresh()
        End If

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled Then
            Form5.Close()
            MsgBox("Operation has been cancelled.")

            Label1.Text = "Copy Pictures From SD Card" + Environment.NewLine + "- SD Card is Connected -"
            Label6.Text = "Operation Canceled."
            Label7.Text = "Operation Canceled."

            Timer1.Enabled = True
            Do While Me.Height > 295
                Me.Height = Me.Height - 2
                Me.Refresh()
            Loop

            Label2.Enabled = True
            Label3.Enabled = True
            Label10.Enabled = True
        Else
            Label1.Text = "Copy Pictures From SD Card" + Environment.NewLine + "- SD Card is Connected -"
            Label6.Text = "100% Transfered all files successfully."

            MsgBox(Label8.Text + " files have successfully transfered." + Environment.NewLine + Label9.Text + " files already existed at destination.", , "Operation complete")

            Timer1.Enabled = True
            Do While Me.Height > 295
                Me.Height = Me.Height - 2
                Me.Refresh()
            Loop

            Label2.Enabled = True
            Label3.Enabled = True
            Label10.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim strlogic As String = ""
        Dim test, bool As String
        Dim strDrive As String

        test = isConnected(strVideoSrc, "dir")
        bool = test.Substring(test.Length - 4)
        strDrive = test.Substring(0, 3)
        Dim strPath As String = strDrive & strVideoSrc
        MsgBox(strPath)

        If String.Compare(bool, "True", True) = 0 Then

            For Each objFoundFile As Object In My.Computer.FileSystem.GetFiles(strPath, FileIO.SearchOption.SearchAllSubDirectories, strVideoExt) 'every video found on video camera
                Dim strFile As FileInfo = New FileInfo(objFoundFile)
                strlogic = Nothing
                strlogic = doesExisttest(strFile.FullName, strFile.CreationTime, strFile.Length)
                MsgBox(strlogic & " -- " & strFile.Name & " -- " & strFile.Length & " -- " & strFile.CreationTime)
            Next
        End If

    End Sub

    Private Sub Label1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.Click
        If Not BackgroundWorker1.IsBusy Then
            Dim test, bool, drive As String
            Dim progress As Integer = 0

            test = isConnected(strSDVol, "vol")
            bool = test.Substring(test.Length - 4)

            drive = test.Substring(0, 2)

            If String.Compare(bool, "True", True) = 0 Then
                If Label1.Text.Contains("Connected") Then

                Else
                    Label1.Text = Label1.Text + Environment.NewLine + "- SD Card is Connected -"
                End If

                Timer1.Enabled = False
                Label6.Text = "0%"
                ProgressBar2.Visible = False
                ProgressBar1.Maximum = 100
                ProgressBar1.Value = 0

                Do While Me.Height < 391
                    Me.Height = Me.Height + 2
                    Me.Refresh()
                Loop

                filesinfolder = 0
                For Each foundFile As Object In My.Computer.FileSystem.GetFiles(drive, FileIO.SearchOption.SearchAllSubDirectories, "*.*")
                    Dim fi As FileInfo = New FileInfo(foundFile)

                    If LCase(fi.Extension) = ".jpg" Then
                        filesinfolder += 1
                    ElseIf LCase(fi.Extension) = ".mov" Then
                        filesinfolder += 1
                    End If
                Next

                If Not BackgroundWorker1.IsBusy Then
                    'temp 3/26/15
                    '::BACK::
                    ''MsgBox("JPG Destination: " + strJPGDest + Environment.NewLine + "MOV Destination: " + strSDVidDest)
                    BackgroundWorker1.RunWorkerAsync(New Object() {Label4.Text, "*.*", strJPGDest, strSDVidDest, filesinfolder})       'text4.text = "drive letter"
                    Label1.Text = "Copying Pictures From SD Card" + Environment.NewLine + "- Click to Cancel Operation -"
                    Label2.Enabled = False
                    Label3.Enabled = False
                    Label10.Enabled = False

                End If
            Else
                Label1.Text = Replace(Label1.Text, Environment.NewLine + "- SD Card is Connected -", "")
                MessageBox.Show("SD Card is not currently connected, please connect the SD card and try again.", "SD Card not detected!")
            End If
        Else
            Form5.Show()
            BackgroundWorker1.CancelAsync()
        End If

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Dim test, bool As String
        Dim exsource As Boolean = False
        Dim exdestination As Boolean = False
        Dim srcPath As String = ""
        Dim dstPath As String = ""
        Dim srcDrive, dstDrive As String

        If Not BackgroundWorker2.IsBusy Then
            test = isConnected(strVideoSrc, "dir")
            bool = test.Substring(test.Length - 4)
            srcDrive = test.Substring(0, 3)
            srcPath = Path.Combine(srcDrive, strVideoSrc)
            dstDrive = strVideoDest.Substring(0, 3)

            If String.Compare(bool, "True", True) = 0 Then
                If Label2.Text.Contains("Connected") Then
                    'nothing
                Else
                    Label2.Text = Label1.Text & vbCrLf & "- Video Camera is Connected -"
                End If

                Timer1.Enabled = False
                Label6.Text = "0%"

                ProgressBar2.Value = 0
                ProgressBar2.Visible = True
                ProgressBar1.Maximum = 100
                ProgressBar1.Value = 0

                Do While Me.Height < 410
                    Me.Height = Me.Height + 2
                    Me.Refresh()
                Loop

                'START RENAME/COPY
                ProgressBar1.Maximum = 100
                ProgressBar1.Value = 0

                BackgroundWorker2.WorkerSupportsCancellation = True
                BackgroundWorker2.WorkerReportsProgress = True
                BackgroundWorker2.RunWorkerAsync(New Object() {srcPath, srcDrive, strVideoDest, dstDrive, 0})
                Label2.Text = "Copying Videos From Camera" + Environment.NewLine + "- Click to Cancel Operation -"
                Label1.Enabled = False
                Label3.Enabled = False
                Label10.Enabled = False
            Else
                MsgBox("Video Camera has not been detected.")
                Label2.Text = Replace(Label2.Text, Environment.NewLine + "- Video Camera is Connected -", "")
            End If
        Else
            Form5.Show()
            BackgroundWorker2.CancelAsync()
        End If

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Application.Exit()

    End Sub

    Private Sub Label8_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.DoubleClick
        form2.Show()
    End Sub

    Private Sub Label1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.MouseEnter
        Label1.BackColor = SystemColors.MenuHighlight
        Label1.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.MouseLeave
        Label1.BackColor = Color.WhiteSmoke
        Label1.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = SystemColors.MenuHighlight
        Label2.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.WhiteSmoke
        Label2.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Label3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseEnter
        Label3.BackColor = SystemColors.MenuHighlight
        Label3.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave
        Label3.BackColor = Color.WhiteSmoke
        Label3.ForeColor = SystemColors.MenuHighlight
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim SD, CAM, bool, drive As String

        SD = isConnected(strSDVol, "vol")
        bool = SD.Substring(SD.Length - 4)
        drive = SD.Substring(0, 2)

        If String.Compare(bool, "True", True) = 0 Then
            Label4.Text = SD.Substring(0, 3)
            If Not Label1.Text.Contains("Connected") Then
                Label1.Text = Label1.Text & vbCrLf & "- SD Card is Connected -"
            End If
        Else
            If Label1.Text.Contains("Connected") Then
                Label1.Text = Replace(Label1.Text, vbCrLf & "- SD Card is Connected -", "")
            End If
        End If

        CAM = isConnected("PRIVATE\AVCHD\BDMV\STREAM", "dir")
        bool = CAM.Substring(CAM.Length - 4)
        drive = CAM.Substring(0, 2)

        If String.Compare(bool, "True", True) = 0 Then
            Label5.Text = CAM.Substring(0, 2)
            If Not Label2.Text.Contains("Connected") Then
                Label2.Text = Label2.Text & vbCrLf & "- Video Camera is Connected -"
            End If
        Else
            If Label2.Text.Contains("Connected") Then
                Label2.Text = Replace(Label2.Text, vbCrLf & "- Video Camera is Connected -", "")
            End If
        End If
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim args As Object() = DirectCast(e.Argument, Object())
        Dim srcDrive As String = CStr(args(1))
        Dim srcPath As String = CStr(args(0))
        Dim dstPath As String = CStr(args(2))
        Dim dstDrive As String = CStr(args(3))
        Dim checked As String = CBool(args(4))
        Dim copy As Boolean
        Dim filecount As Integer = 0
        Dim SourceSize As Long = GetFolderSize(srcPath)
        Dim progress As Long = 0
        Dim x As Long = 0
        Dim q As Integer = 0
        Dim mb As Integer = 0
        Dim a, b As Integer
        Dim mymodified As String = ""
        Dim mysize As System.IO.FileInfo

        Dim Destfiles As String() = Directory.GetFiles(dstPath, strVideoExt)                'put files at destination in a list
        Dim srcfiles As String() = Directory.GetFiles(srcPath, strVideoExt)                 'adds source files to an array

        Dim mycreation As String = ""

        For Each srcfile As Object In srcfiles
            If BackgroundWorker2.CancellationPending = True Then
                ' Set Cancel to True
                e.Cancel = True
                Exit For
            End If

            filecount += 1
            Dim cursrcfile As FileInfo = New FileInfo(srcfile)
            mymodified = cursrcfile.LastWriteTime
            If checked = True Then
                mycreation = cursrcfile.LastWriteTime
            Else
                mycreation = cursrcfile.CreationTime
            End If
            Dim srcfilename As String = filenameformat(mycreation, cursrcfile.Extension)

            If Destfiles.Count > 0 Then
                For Each destfile As Object In Destfiles
                    If BackgroundWorker2.CancellationPending = True Then
                        ' Set Cancel to True
                        e.Cancel = True
                        Exit For
                    End If

                    copy = True
                    Dim curdestfile As FileInfo = New FileInfo(destfile)
                    If curdestfile.Name = srcfilename And curdestfile.Length >= cursrcfile.Length Then
                        x += cursrcfile.Length
                        progress = ((x * 100) / SourceSize)
                        BackgroundWorker2.ReportProgress(progress)
                        SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...  [" & Format((cursrcfile.Length / 1024), "###,###,###") & " KB]")
                        SetLabelText_ThreadSafe(Me.Label7, "Transfering: " + Environment.NewLine + srcfile + Environment.NewLine + "To Destination: " + Environment.NewLine + Path.Combine(dstPath, srcfilename))
                        b += 1
                        copy = False
                        Exit For
                    End If
                Next
            Else
                copy = True
            End If

            If copy = True Then
                On Error Resume Next
                a += 1
                Dim source As String = Path.Combine(srcPath, cursrcfile.Name)
                Dim dest As String = Path.Combine(dstPath, srcfilename)

                If File.Exists(dest) Then
                    File.Delete(dest)
                End If

                Dim inputstream = New System.IO.FileStream(source, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite)
                Dim outputstream = New System.IO.FileStream(dest, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)

                Dim buffer = New Byte(64 * 4096) {}
                Dim bytesRead As Integer = 0

                Do
                    bytesRead = inputstream.Read(buffer, 0, buffer.Length)
                    If bytesRead > 0 Then
                        mysize = My.Computer.FileSystem.GetFileInfo(dest)
                        x += bytesRead
                        progress = ((x * 100) / SourceSize)
                        mb = (x / 1048576)
                        outputstream.Write(buffer, 0, bytesRead)
                        BackgroundWorker2.ReportProgress(progress)
                        SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...  [" & Format((mysize.Length / 1024), "###,###,###") & " - " & Format((cursrcfile.Length / 1024), "###,###,###") & " KB]")
                        SetLabelText_ThreadSafe(Me.Label7, "Transfering: " + Environment.NewLine + srcfile + Environment.NewLine + "To Destination: " + Environment.NewLine + Path.Combine(dstPath, srcfilename))
                        SetLabelText_ThreadSafe(Me.Label11, Format((mysize.Length / 1024), "###,###,###"))
                        SetLabelText_ThreadSafe(Me.Label12, Format((cursrcfile.Length / 1024), "###,###,###"))
                    End If

                    BackgroundWorker2.ReportProgress(progress)
                    SetLabelText_ThreadSafe(Me.Label6, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...  [" & Format((mysize.Length / 1024), "###,###,###") & " - " & Format((cursrcfile.Length / 1024), "###,###,###") & " KB]")
                    SetLabelText_ThreadSafe(Me.Label7, "Transfering: " + Environment.NewLine + srcfile + Environment.NewLine + "To Destination: " + Environment.NewLine + Path.Combine(dstPath, srcfilename))
                    SetLabelText_ThreadSafe(Me.Label11, Format((mysize.Length / 1024), "###,###,###"))
                    SetLabelText_ThreadSafe(Me.Label12, Format((cursrcfile.Length / 1024), "###,###,###"))
                Loop While bytesRead > 0

                outputstream.Flush()
                inputstream.Flush()
                outputstream.Close()
                inputstream.Close()

                File.SetCreationTime(dest, mycreation)
                File.SetLastWriteTime(dest, mymodified)

                If BackgroundWorker1.CancellationPending = True Then
                    ' Set Cancel to True
                    e.Cancel = True
                    Exit For
                End If
            End If
        Next
        SetLabelText_ThreadSafe(Me.Label8, a)
        SetLabelText_ThreadSafe(Me.Label9, b)
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Dim num1, num2 As Long

        If e.ProgressPercentage <= 100 Then
            Me.ProgressBar1.Value = e.ProgressPercentage
            Me.ProgressBar1.Refresh()
        End If

        Int32.TryParse(Replace(Label12.Text, ",", ""), num2)
        Int32.TryParse(Replace(Label11.Text, ",", ""), num1)

        Me.ProgressBar2.Maximum = num2
        Me.ProgressBar2.Value = num1

    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        If e.Cancelled Then
            Form5.Close()

            Label6.Text = "Operation Canceled."
            Label7.Text = "Operation Canceled."

            MsgBox("Operation has been cancelled.")

            Timer1.Enabled = True
            Do While Me.Height > 290
                Me.Height = Me.Height - 2
                Me.Refresh()
            Loop

            ProgressBar2.Visible = False
            Label1.Enabled = True
            Label3.Enabled = True

            If Label2.Text.Contains("Cancel") Then
                Label2.Text = "Copy Videos From Camera" + Environment.NewLine + "- Video Camera is Connected -"
                Label10.Enabled = True
            End If
            If Label10.Text.Contains("Cancel") Then
                Label10.Text = "Copy Videos From Folder"
                Label2.Enabled = True
            End If
        Else
            Label6.Text = "100% Transfered all files successfully."

            MsgBox(Label8.Text & " files have successfully transfered." & vbCrLf & Label9.Text & " files already existed at destination.", , "Operation complete")

            Timer1.Enabled = True
            Do While Me.Height > 290
                Me.Height = Me.Height - 2
                Me.Refresh()
            Loop

            Label1.Enabled = True
            Label3.Enabled = True

            If Label2.Text.Contains("Cancel") Then
                Label2.Text = "Copy Videos From Camera" + Environment.NewLine + "- Video Camera is Connected -"
                Label10.Enabled = True
            End If
            If Label10.Text.Contains("Cancel") Then
                Label10.Text = "Copy Videos From Folder"
                Label2.Enabled = True
            End If

        End If

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Dim exsource As Boolean = False
        Dim exdestination As Boolean = False
        Dim srcPath As String = ""
        Dim dstPath As String = ""
        Dim srcDrive, dstDrive As String

        If Not BackgroundWorker2.IsBusy Then
            'Pick Folder to copy Videos From.
            'Folder Browser
            Dim folderDlg As New FolderBrowserDialog
            folderDlg.Reset()
            folderDlg.Description = "Choose Source Folder"
            folderDlg.SelectedPath = "X:\Camera Backup\Family Videos\Bainbridge"

            folderDlg.ShowNewFolderButton = True
            If (folderDlg.ShowDialog() = DialogResult.OK) Then
                Dim root As Environment.SpecialFolder = folderDlg.RootFolder
                srcPath = folderDlg.SelectedPath
                srcDrive = Mid(srcPath, 1, 3)
            Else
                MsgBox("You must choose a folder.")
                Exit Sub
            End If

            dstDrive = strVideoDest.Substring(0, 3)

            Timer1.Enabled = False
            Label6.Text = "0%"

            ProgressBar2.Value = 0
            ProgressBar2.Visible = True
            ProgressBar1.Maximum = 100
            ProgressBar1.Value = 0

            Do While Me.Height < 410
                Me.Height = Me.Height + 2
                Me.Refresh()
            Loop

            BackgroundWorker2.WorkerSupportsCancellation = True
            BackgroundWorker2.WorkerReportsProgress = True
            BackgroundWorker2.RunWorkerAsync(New Object() {srcPath, srcDrive, strVideoDest, dstDrive, 0})
            Label10.Text = "Copying Videos From Folder" + Environment.NewLine + "- Click to Cancel Operation -"
            Label2.Enabled = False
            Label3.Enabled = False
            Label1.Enabled = False
        Else
            Form5.Show()
            BackgroundWorker2.CancelAsync()
        End If
    End Sub

    Private Sub Label10_MouseEnter(sender As Object, e As EventArgs) Handles Label10.MouseEnter
        Label10.BackColor = SystemColors.MenuHighlight
        Label10.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label10_MouseLeave(sender As Object, e As EventArgs) Handles Label10.MouseLeave
        Label10.BackColor = Color.WhiteSmoke
        Label10.ForeColor = SystemColors.MenuHighlight
    End Sub

End Class