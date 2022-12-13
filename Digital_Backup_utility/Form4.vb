Imports System
Imports System.IO
Imports System.Collections
Imports Microsoft.VisualBasic.FileIO
Imports System.Runtime.InteropServices
Imports System.Reflection


Public Class Form4

    ' The delegate
    Delegate Sub SetLabelText_Delegate(ByVal [Label] As Label, ByVal [text] As String)

    ' The delegates subroutine.
    Private Sub SetLabelText_ThreadSafe(ByVal [Label] As Label, ByVal [text] As String)
        ' InvokeRequired compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If [Label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[Label], [text]})
        Else
            [Label].Text = [text]
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim exsource As Boolean = False
        Dim exdestination As Boolean = False
        Dim srcPath As String = ""
        Dim dstPath As String = ""
        Dim srcDrive, dstDrive As String


        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            srcPath = TextBox1.Text
            srcDrive = srcPath.Substring(0, 3)                         'source drive letter
            dstPath = TextBox2.Text
            dstDrive = dstPath.Substring(0, 3)

            If Directory.Exists(srcPath) Then
                exsource = True
            Else
                MsgBox("Source directory doesn't exist.")
                Exit Sub
            End If

            If Directory.Exists(dstPath) Then
                exdestination = True
            Else
                MsgBox("Destination directory doesn't exist.")
                Exit Sub
            End If
        Else
            MsgBox("Both textboxes must be filled out.")
            Exit Sub
        End If

        'START RENAME/COPY

        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0

        If Not BackgroundWorker1.IsBusy Then
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.WorkerReportsProgress = True
            BackgroundWorker1.RunWorkerAsync(New Object() {srcPath, srcDrive, dstPath, dstDrive, CheckBox1.CheckState})

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim srcfiles As String() = Directory.GetFiles(TextBox2.Text, "*.mts")

        For Each srcfile As Object In srcfiles
            Dim cursrcfile As FileInfo = New FileInfo(srcfile)
            'MsgBox("Modified: " & cursrcfile.LastWriteTime & " - Created: " & cursrcfile.CreationTime)
            Dim mycreation As String = cursrcfile.CreationTime

            File.SetCreationTime(srcfile, cursrcfile.LastWriteTime)
            Dim srcfilename As String = filenameformat(cursrcfile.CreationTime, cursrcfile.Extension)
            MsgBox(srcfilename)
            My.Computer.FileSystem.RenameFile(srcfile, srcfilename)
        Next

    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        If CheckBox2.Checked = True Then
            CheckBox2.Checked = False
        End If
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        If CheckBox1.Checked = True Then
            CheckBox1.Checked = False
            CheckBox2.Checked = True
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim args As Object() = DirectCast(e.Argument, Object())
        Dim srcPath As String = CStr(args(0))
        Dim srcDrive As String = CStr(args(1))
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


        Dim Destfiles As String() = Directory.GetFiles(dstPath, "*.mts")                'put files at destination in a list
        Dim srcfiles As String() = Directory.GetFiles(srcPath, "*.mts")                 'adds source files to an array

        Dim mycreation As String = ""

        For Each srcfile As Object In srcfiles
            filecount += 1
            Dim cursrcfile As FileInfo = New FileInfo(srcfile)
            If checked = True Then
                mycreation = cursrcfile.LastWriteTime
            Else
                mycreation = cursrcfile.CreationTime
            End If
            Dim srcfilename As String = filenameformat(mycreation, cursrcfile.Extension)

            If Destfiles.Count > 0 Then
                For Each destfile As Object In Destfiles
                    copy = True
                    Dim curdestfile As FileInfo = New FileInfo(destfile)
                    If curdestfile.Name = srcfilename Then
                        x += curdestfile.Length
                        progress = ((x * 100) / SourceSize)
                        BackgroundWorker1.ReportProgress(progress)
                        SetLabelText_ThreadSafe(Me.Label3, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...")
                        copy = False
                        Exit For
                    End If
                Next
            Else
                copy = True
            End If


            If copy = True Then
                Dim source As String = Path.Combine(srcPath, cursrcfile.Name)
                Dim dest As String = Path.Combine(dstPath, srcfilename)

                Dim inputstream = New System.IO.FileStream(source, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite)
                Dim outputstream = New System.IO.FileStream(dest, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)

                Dim buffer = New Byte(64 * 4096) {}
                Dim bytesRead As Integer = 0

                Do
                    bytesRead = inputstream.Read(buffer, 0, buffer.Length)
                    If bytesRead > 0 Then
                        x += bytesRead
                        progress = ((x * 100) / SourceSize)
                        mb = (x / 1048576)
                        outputstream.Write(buffer, 0, bytesRead)
                        BackgroundWorker1.ReportProgress(progress)
                        'Me.ProgressBar1.Value = progress
                        'SetLabelText_ThreadSafe(Me.Label3, Format(mb, "###,###,###,##0 MB"))
                        'Label3.Text = progress & "%" & "  Transfering [" & counter & " of " & filecount & "] Files..."
                        SetLabelText_ThreadSafe(Me.Label3, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...")
                    End If

                    BackgroundWorker1.ReportProgress(progress)
                    SetLabelText_ThreadSafe(Me.Label3, progress & "%" & "  Transfering [" & filecount & " of " & srcfiles.Count & "] Files...")

                Loop While bytesRead > 0

                outputstream.Flush()
                inputstream.Flush()
                outputstream.Close()
                inputstream.Close()

                File.SetCreationTime(dest, mycreation)
            End If

        Next
        SetLabelText_ThreadSafe(Me.Label3, progress & "%" & "  Transfered [" & filecount & "] Files...")
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Me.ProgressBar1.Value = e.ProgressPercentage
        Me.ProgressBar1.Refresh()

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        MsgBox("done.")
    End Sub

End Class