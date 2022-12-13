Imports System.IO
Imports System.Security.Cryptography
Module Module1
    Public strSDVol As String
    Public strJPGDest As String
    Public strSDVidDest As String
    Public strSDVidExt As String
    Public strVideoSrc As String
    Public strVideoExt As String
    Public strVideoDest As String

    Public Function filenameformat(ByVal thedate As String, ByVal ext As String)
        filenameformat = LCase(Replace(thedate, "/", "-"))
        filenameformat = Replace(filenameformat, ":", "")
        filenameformat = Replace(filenameformat, " am", "")
        filenameformat = Replace(filenameformat, " pm", "")
        filenameformat = filenameformat & ext
    End Function



    Function isConnected(ByVal strDir As String, ByVal strType As String) As String
        Dim allDrives() As IO.DriveInfo = IO.DriveInfo.GetDrives()
        Dim d As IO.DriveInfo

        isConnected = ""

        For Each d In allDrives
            If d.IsReady = False Or d.DriveType <> IO.DriveType.Removable Then
                isConnected = "False"
            ElseIf d.IsReady = True AndAlso d.DriveType = IO.DriveType.Removable Then

                Select Case strType
                    Case "dir"
                        If IO.Directory.Exists(IO.Path.Combine(d.Name, strDir)) Then
                            isConnected = d.Name & " True"
                            Exit Function
                        End If
                    Case "vol"
                        If String.Compare(d.VolumeLabel, strDir, True) = 0 Then
                            isConnected = d.Name & " True"
                            Exit Function
                        End If
                    Case Else
                        isConnected = "False"
                End Select
            End If

        Next


    End Function

    Public Function CompareFiles(ByVal FileFullPath1 As String, ByVal FileFullPath2 As String) As Boolean

        'returns true if two files passed to it are identical, false
        'otherwise

        'does byte comparison; works for both text and binary files

        'Throws exception on errors; you can change to just return 
        'false if you prefer


        Dim objMD5 As New MD5CryptoServiceProvider()
        Dim objEncoding As New System.Text.ASCIIEncoding()

        Dim aFile1() As Byte, aFile2() As Byte
        Dim strContents1, strContents2 As String
        Dim objReader As StreamReader
        Dim objFS As FileStream
        Dim bAns As Boolean
        If Not File.Exists(FileFullPath1) Then Throw New Exception(FileFullPath1 & " doesn't exist")
        If Not File.Exists(FileFullPath2) Then Throw New Exception(FileFullPath2 & " doesn't exist")

        Try

            objFS = New FileStream(FileFullPath1, FileMode.Open)
            objReader = New StreamReader(objFS)
            aFile1 = objEncoding.GetBytes(objReader.ReadToEnd)
            strContents1 = objEncoding.GetString(objMD5.ComputeHash(aFile1))
            objReader.Close()
            objFS.Close()


            objFS = New FileStream(FileFullPath2, FileMode.Open)
            objReader = New StreamReader(objFS)
            aFile2 = objEncoding.GetBytes(objReader.ReadToEnd)
            strContents2 = objEncoding.GetString(objMD5.ComputeHash(aFile2))

            bAns = strContents1 = strContents2
            objReader.Close()
            objFS.Close()
            aFile1 = Nothing
            aFile2 = Nothing

        Catch ex As Exception
            Throw ex

        End Try

        Return bAns
    End Function

    Sub check_settings()
        Dim i As Integer = 0
        Dim Keys() As String = {"Camera_SD_Volume", "Camera_jpg_Destination", "Camera_vid_Destination", "Camera_Vid_Type", "Video_vid_Source", "Video_Type", "Video_Destination"}

        For Each Value As String In Keys
            If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing) Is Nothing Then
                'create them all as default settings here.
                My.Computer.Registry.CurrentUser.CreateSubKey("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_SD_Volume", "EOS_DIGITAL")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_jpg_Destination", "P:\Canon T2i EOS")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_vid_Destination", "X:\Camera Backup\Family Videos\Bainbridge\DSLR Movies")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Camera_Vid_Type", "*.MOV")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_vid_Source", "PRIVATE\AVCHD\BDMV\STREAM")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_Type", "*.MTS")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", "Video_Destination", "X:\Camera Backup\Family Videos\Bainbridge\New Camcorder")
            Else
                'LOAD SETTINGS HERE!
                Select Case Value

                    Case "Camera_SD_Volume"
                        strSDVol = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Camera_jpg_Destination"
                        strJPGDest = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Camera_vid_Destination"
                        strSDVidDest = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Camera_Vid_Type"
                        strSDVidExt = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Video_vid_Source"
                        strVideoSrc = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Video_Type"
                        strVideoExt = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                    Case "Video_Destination"
                        strVideoDest = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Bainbridge Digital Backup", Value, Nothing)

                End Select

            End If
        Next
    End Sub

    Function GetFolderSize(ByVal DirPath As String, Optional ByVal IncludeSubFolders As Boolean = True) As Long

        Dim lngDirSize As Long
        Dim objFileInfo As FileInfo
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo

        Try

            'add length of each file
            For Each objFileInfo In objDir.GetFiles(strVideoExt)
                lngDirSize += objFileInfo.Length
            Next

            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    lngDirSize += GetFolderSize(objSubFolder.FullName)
                Next
            End If

        Catch Ex As Exception


        End Try

        Return lngDirSize
    End Function



    Function GetFolderSize1(ByVal DirPath As String) As Long
        Dim lngDirSize As Long

        Try
            Dim srcfiles As String() = Directory.GetFiles(DirPath, strVideoExt)

            For Each File As Object In srcfiles
                lngDirSize += File.length
            Next
        Catch Ex As Exception

        End Try
        Return lngDirSize

    End Function

End Module
