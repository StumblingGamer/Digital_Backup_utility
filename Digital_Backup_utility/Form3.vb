Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Public Class Form3

    Public Sub CopyFile(ByVal sourceFileName As String, ByVal destinationFileName As String, ByVal showUI As UIOption, ByVal onUserCancel As UICancelOption)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sourceDirectory As String = "X:\Camera Backup\Family Videos\Bainbridge\New Camcorder"
        Dim destinationDirectory As String = "X:\temp"

        'If Not Directory.Exists(destinationDirectory) Then
        'Dim di As DirectoryInfo = Directory.CreateDirectory(destinationDirectory)
        ' End If

        If Directory.Exists(sourceDirectory) Then 'AndAlso Directory.Exists(destinationDirectory) Then
            Directory.Move(sourceDirectory, destinationDirectory)
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim source As String = "X:\temp"
        Dim destination As String = "X:\Camera Backup\Family Videos\Bainbridge\New Camcorder"



        BackgroundWorker1.RunWorkerAsync(New Object() {source, destination})  ' passes source and destination to backgroundworker


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Creates a 6 digit filename
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Dim filename As String = 0
        'Dim filename1 As String
        '
        'Do While filename < 150
        'Do While filename.Length < 6
        'filename = "0" & filename
        'Loop
        'filename1 = filename & ".mts"
        ''ListBox1.Items.Add(filename1)
        ''ListBox1.Refresh()
        'filename += 1
        'Loop

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Creates a 6 digit filename
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim fi As New System.IO.DirectoryInfo("X:\Camera Backup\Family Videos\Bainbridge\New Camcorder\")
        Dim files = fi.GetFiles("*.mts").ToList
        Dim last = (From file In files Select file Order By file.CreationTime Descending).FirstOrDefault
        MsgBox((Replace(Split(last.Name, ".", 2)(0), "0", "")) + 1)

        Dim filename As String = (Replace(Split(last.Name, ".", 2)(0), "0", "")) + 1
        Dim filename1 As String
        Dim strSource As String = "X:\temp\"
        Dim strDest As String = "X:\Camera Backup\Family Videos\Bainbridge\New Camcorder\"


        Dim orderedFiles = New System.IO.DirectoryInfo(strSource).GetFiles("*.mts").OrderBy(Function(x) x.LastWriteTime)
        For Each f As System.IO.FileInfo In orderedFiles
            Do While filename.Length < 6
                filename = "0" & filename
            Loop

            filename1 = filename & ".MTS"
            My.Computer.FileSystem.CopyFile(f.FullName, strDest & filename1)

            'ListBox2.Items.Add(String.Format("{0,-15} {1,12}", f.Name, f.LastWriteTime.ToString) & " -- Newname: " & strDest & filename1)
            'ListBox1.Items.Add(filename1)
            'ListBox2.Refresh()
            filename += 1
        Next
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim strDest As String = "X:\Camera Backup\Family Videos\Bainbridge\New Camcorder\"
        Dim fi2 As New System.IO.DirectoryInfo(strDest)
        Dim files = fi2.GetFiles("*.mts").ToList
        Dim last = (From file In files Select file Order By file.CreationTime Descending).FirstOrDefault

        Dim filename As String = (Replace(Split(last.Name, ".", 2)(0), "0", "")) + 1
        Dim filename1 As String
        Dim strSource As String = "X:\temp\"
        Dim fi As FileInfo
        Dim fi1 As FileInfo
        Dim li(3) As String
        Dim sourcebytes As Long = 0
        Dim totalbytes As Long = 0
        Dim progress As Long = 0
        Dim destbytes As Long = 0
        Dim transfered As Long = 0

        For Each foundFile1 As String In My.Computer.FileSystem.GetFiles(strSource, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mts")
            fi1 = New FileInfo(foundFile1)
            If CBool(Not ((fi1.Attributes And FileAttributes.System) Or (fi1.Attributes And FileAttributes.Hidden))) Then
                sourcebytes += fi1.Length
            End If
        Next

        Dim orderedFiles = New System.IO.DirectoryInfo(strSource).GetFiles("*.mts").OrderBy(Function(x) x.LastWriteTime)

        For Each f As System.IO.FileInfo In orderedFiles
            Do While filename.Length < 6
                filename = "0" & filename
            Loop

            destbytes = 0
            filename1 = filename & ".MTS"

            Dim i As Integer = ((transfered * 100) / sourcebytes)
            BackgroundWorker1.ReportProgress(i)

            My.Computer.FileSystem.CopyFile(f.FullName, strDest & filename1)  ', showUI:=FileIO.UIOption.AllDialogs
            transfered += f.Length
            ListBox2.Items.Add(f.FullName & " -- Copied to: " & strDest & filename1)
            For Each foundFile1 As String In My.Computer.FileSystem.GetFiles(strDest, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.mts")
                fi = New FileInfo(foundFile1)
                'If CBool(Not ((fi.Attributes And FileAttributes.System) Or (fi.Attributes And FileAttributes.Hidden))) Then
                'destbytes += fi.Length
                'End If    
                i = ((transfered * 100) / sourcebytes)
                BackgroundWorker1.ReportProgress(i)
            Next
            Label2.Text = Format(transfered, "###,###,###,##0.00 Bytes") & " of " & Format(sourcebytes, "###,###,###,##0.00 Bytes")
            destbytes = 0
            'ListBox2.Items.Add(String.Format("{0,-15} {1,12}", f.Name, f.LastWriteTime.ToString) & " -- Newname: " & strDest & filename1)
            'ListBox1.Items.Add(filename1)
            'ListBox2.Refresh()
            i = ((transfered * 100) / sourcebytes)
            BackgroundWorker1.ReportProgress(i)
            filename += 1
        Next
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
        Label1.Text = ProgressBar1.Value & "%"

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        MsgBox("Done")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim fi As New System.IO.DirectoryInfo("X:\Camera Backup\Family Videos\Bainbridge\New Camcorder\")
        Dim files = fi.GetFiles("*.mts").ToList
        Dim last = (From file In files Select file Order By file.CreationTime Descending).FirstOrDefault
        MsgBox((Replace(Split(last.Name, ".", 2)(0), "0", "")) + 1)

    End Sub
End Class