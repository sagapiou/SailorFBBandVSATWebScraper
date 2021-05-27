Imports System
Imports System.Text
Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Specialized
Imports System.Data.SqlClient

Enum vTypeOfSystem As Integer
    SailorVSat = 1
    SailorFBB = 2
End Enum

Module WebScraper
    Private Downloading As Boolean = False
    Private DownloadSpeed As String = ""
    Private DownloadFileSize As Long = 0

    Sub Main()
        Try
            Dim vVsatNet As Boolean = False
            Console.WriteLine("Welcome to Almi Tankers Weekly Checks for Backups and Comms")
            Console.WriteLine("===========================================================")
            Console.WriteLine("")
            Console.WriteLine("BACKUPS :")
            Console.WriteLine("===========================================================")
            OracleDump()
            SqlBackup()
            BackupExec()
            Console.WriteLine("")
            Console.WriteLine("SYNOLOGY :")
            Console.WriteLine("===========================================================")
            SynologyFreeSpace()
            Console.WriteLine("")
            Console.WriteLine("")
            Console.WriteLine("COMMUNICATIONS :")
            Console.WriteLine("===========================================================")
            If isFBBProvidingInternet() Then
                Console.WriteLine("FBB LIVE - FBB has open data sessions.")
                vVsatNet = False
            Else
                Console.WriteLine("FBB BACKUP - FBB doesn't have open data sessions.")
                vVsatNet = True
            End If
            Dim vSat As Collection = vSatLoginAndReturnData()
            If vSat Is Nothing Then
                vSat = vSatLoginAndReturnData() ' run once more in case it returns a null object
            End If
            Console.Write("VSAT : Strength - " & vSat("SignalStrength") & ", ")
            Console.Write(vSat("Tracking") & ", ")
            Console.Write("RX - " & vSat("RXLocked") & ", ")
            Console.Write("Net - " & vSat("Status"))
            If vVsatNet Then
                Console.WriteLine("")
                Console.WriteLine("")
                Console.WriteLine("SPEED TEST :")
                Console.WriteLine("===========================================================")
                CalculateDownloadSpeed()
            End If
            Console.ReadLine()

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.ReadLine()
        End Try

    End Sub

    Private Function vSatLoginAndReturnData() As Collection
        Try
            System.Net.ServicePointManager.Expect100Continue = False
            Dim vSATSetting As New Collection
            Dim GlobalCookieContainer As New CookieContainer

            Dim username As String = "admin"                                                    'Username
            Dim password As String = "1234"                                                'Password
            Dim URL As String = "http://192.168.239.1/?pageId=login"

            Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)
            request.CookieContainer = GlobalCookieContainer                   'Cookie wiederverwenden
            request.Method = "POST"                                                                 'Set the Method property of the request to POST.
            Dim postData As String = "user_login=" & username & "&pass_login=" & password           'Create POST data and convert it to a byte array.
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            request.ContentType = "application/x-www-form-urlencoded"                           ' Set the ContentType property of the WebRequest.    
            request.ContentLength = byteArray.Length                                            ' Set the ContentLength property of the WebRequest.

            'DATA STREAM
            Dim dataStream As Stream = request.GetRequestStream()                               ' Get the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length)                                    ' Write the data to the request stream.
            dataStream.Close()                                                                  ' Close the Stream object.

            'RESPONSE
            Dim resp As HttpWebResponse = request.GetResponse()                                 ' Get the response.
            'Console.WriteLine("resp.StatusCode ist: {0}, resp.StatusDescription ist: {1}", Int(resp.StatusCode), resp.StatusDescription)
            'Console.WriteLine("Server {0}, Uri {1}, Method {2}, Cookies {3}", resp.Server, resp.ResponseUri, resp.Method, resp.Cookies)
            'Console.ReadKey()                                                                     ' Display elements of repsonse that came from the request issued above.

            dataStream = resp.GetResponseStream()                                               ' Get the stream containing content returned by the server.
            Dim reader As New StreamReader(dataStream)                                          ' Open the stream using a StreamReader for easy access.
            Dim strOutput As String = reader.ReadToEnd()                               ' Read the content.
            'Console.WriteLine(responseFromServer)                                               ' Display the content.

            reader.Close()                                                                      ' Clean up the streams.
            dataStream.Close()
            resp.Close()

            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("rssi_blocks") + 14, 1), "SignalStrength")
            If strOutput.IndexOf(Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "Tracking" & Chr(34)) <> -1 Then
                vSATSetting.Add(Mid(strOutput, strOutput.IndexOf(Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "Tracking" & Chr(34)) + 1, 19), "Tracking")
            Else
                vSATSetting.Add("Not Tracking", "Tracking")
            End If
            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("'rx_locked_status'") + 20, 6), "RXLocked")
            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("'modem_status'") + 16, 5), "Status")
            If strOutput.Contains("No active data sessions") Then
                '    vIsFBBProvidingInternet = False
            Else
                '   vIsFBBProvidingInternet = True
            End If

            Return vSATSetting

        Catch e As WebException
            Console.WriteLine("Attempt to reconnect due to web page not properly opening.")
            Console.WriteLine(e.Message)
            If e.Status = WebExceptionStatus.ProtocolError Then
                Console.WriteLine("Status Code : {0}", CType(e.Response, HttpWebResponse).StatusCode)
                Console.WriteLine("Status Description : {0}", CType(e.Response, HttpWebResponse).StatusDescription)
                Console.WriteLine("Resonse URI: {0}", e.Response.ResponseUri)
                Console.WriteLine("IsFromCache: {0}", e.Response.IsFromCache)
                Console.WriteLine("IsMutuallyAuthenticated: {0}", e.Response.IsMutuallyAuthenticated)
                Console.WriteLine()
                Return Nothing
            End If

        Catch e As Exception
            Console.WriteLine(e.Message)
            Return Nothing
        End Try
    End Function

    Private Function isFBBProvidingInternet() As Boolean
        Dim vIsFBBProvidingInternet As Boolean = False
        Try
            Dim strURL As String = ""
            If My.Computer.Network.Ping("192.168.0.1") Then
                strURL = "http://192.168.0.1"
            ElseIf My.Computer.Network.Ping("172.18.128.57") Then
                strURL = "http://172.18.128.57:51080/"
            ElseIf My.Computer.Network.Ping("172.18.128.205") Then
                strURL = "http://172.18.128.205:51080/"
            ElseIf My.Computer.Network.Ping("172.18.128.58") Then
                strURL = "http://172.18.128.58:51080/"
            Else
                ' do something send email etc
            End If
            strURL = "http://192.168.0.1"
            Dim strOutput As String = ""
            Dim wrResponse As WebResponse
            Dim wrRequest As WebRequest = HttpWebRequest.Create(strURL)

            wrResponse = wrRequest.GetResponse()

            Using sr As New StreamReader(wrResponse.GetResponseStream())
                strOutput = sr.ReadToEnd()
                ' Close and clean up the StreamReader
                sr.Close()
            End Using

            'Console.WriteLine(strOutput)
            If strOutput.Contains("No active data sessions") Then
                vIsFBBProvidingInternet = False
            Else
                vIsFBBProvidingInternet = True
            End If
            Return vIsFBBProvidingInternet
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            ' Console.ReadLine()
            Return vIsFBBProvidingInternet
        End Try
    End Function

    Private Function returnVsatSettings() As Collection
        Try
            Dim vSATSetting As New Collection
            'Dim strURL As String = "http://212.165.106.96"
            Dim strURL As String = "http://192.168.239.1"
            'Dim strURL As String = "C:\Users\sagapiou.ALMITANKERS\Desktop\testVsat.htm"

            Dim strOutput As String = ""

            Dim wrResponse As WebResponse
            Dim wrRequest As WebRequest = HttpWebRequest.Create(strURL)

            Console.WriteLine("Extracting Web data ..." & Environment.NewLine)

            wrResponse = wrRequest.GetResponse()

            Using sr As New StreamReader(wrResponse.GetResponseStream())
                strOutput = sr.ReadToEnd()
                ' Close and clean up the StreamReader
                sr.Close()
            End Using
            Console.WriteLine(strOutput)
            Console.ReadLine()

            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("rssi_blocks") + 14, 1), "SignalStrength")
            If strOutput.IndexOf(Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "Tracking" & Chr(34)) <> -1 Then
                vSATSetting.Add(Mid(strOutput, strOutput.IndexOf(Chr(34) & "status" & Chr(34) & ":" & Chr(34) & "Tracking" & Chr(34)) + 1, 19), "Tracking")
            Else
                vSATSetting.Add("Not Tracking", "Tracking")
            End If
            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("'rx_locked_status'") + 20, 6), "RXLocked")
            vSATSetting.Add(Mid(strOutput, strOutput.IndexOf("'modem_status'") + 16, 5), "Status")
            If strOutput.Contains("No active data sessions") Then
                '    vIsFBBProvidingInternet = False
            Else
                '   vIsFBBProvidingInternet = True
            End If

            Return vSATSetting
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub OracleDump()
        Dim dirFound As Boolean = False
        Dim vFileFound As Boolean = False
        Try
            Dim oraPath As String = "C:\backup"

            If IO.Directory.Exists(oraPath) Then
                dirFound = True
            End If

            If dirFound Then
                Dim files() As String
                files = IO.Directory.GetFiles(oraPath, "ONBOARD_FULL.DMP")
                For Each vfile In files
                    Dim dt As DateTime = File.GetLastWriteTime(vfile)
                    If DateDiff(DateInterval.Day, Now, dt) = -1 Or DateDiff(DateInterval.Day, Now, dt) = 0 Then
                        vFileFound = True
                        Dim fi As FileInfo = New FileInfo(vfile)
                        Console.Write("Oracle :")
                        Console.Write(Path.GetFileName(vfile) & ", ")
                        Console.Write(dt & ", ")
                        Console.WriteLine(Math.Round(fi.Length / 1000000, 2) & " MB ")
                    End If
                Next
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub SqlBackup()
        Dim doesFileExist As Boolean = False
        Dim vFileFound As Boolean = False
        Dim vDirFound As Boolean = False
        Try
            Dim SQLPath As String = "C:\Program Files (x86)\Microsoft SQL Server\MSSQL$KAPA32SERVER\BACKUP"
            If Not IO.Directory.Exists(SQLPath) Then
                SQLPath = "C:\Program Files (x86)\Microsoft SQL Server\MSSQL$KAPA32SRV\BACKUP"
                If Not IO.Directory.Exists(SQLPath) Then
                    SQLPath = "C:\Program Files\Microsoft SQL Server\MSSQL10_50.KAPA32SERVER\MSSQL\Backup"
                    If Not IO.Directory.Exists(SQLPath) Then
                        ' check another directory
                    Else
                        vDirFound = True
                    End If
                Else
                    vDirFound = True
                End If
            Else
                vDirFound = True
            End If
            If vDirFound Then
                Dim files() As String
                files = IO.Directory.GetFiles(SQLPath, "*AlmiTank_383*.bak")
                For Each vfile In files
                    If Not vFileFound Then
                        Dim dt As DateTime = File.GetLastWriteTime(vfile)
                        If DateDiff(DateInterval.Day, Now, dt) = -1 Or DateDiff(DateInterval.Day, Now, dt) = 0 Then
                            vFileFound = True
                            Dim fi As FileInfo = New FileInfo(vfile)
                            Console.Write("KAPA : ")
                            Console.Write(Path.GetFileName(vfile) & ", ")
                            Console.Write(dt & ", ")
                            Console.WriteLine(Math.Round(fi.Length / 1000000, 2) & " MB ")
                            Exit For
                        End If
                    End If
                Next
            End If
            If doesFileExist Then
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub BackupExec()

        Dim vConnectionString As String
        Dim vSqlConn As SqlConnection
        Dim vSqlCmd As SqlCommand
        Dim vSql As String = ""
        ' Use this first connection string if using windows auth 
        ' connectionString = "Data Source=ServerName;Initial Catalog=DatabaseName;Integrated Security=True"
        'vConnectionString = "Data Source=SERVER\BKUPEXEC;Initial Catalog=BEDB;User ID=almitankers.local\administrator;Password=almitank3rs"
        vConnectionString = "Data Source=SERVER\BKUPEXEC;Initial Catalog=BEDB;integrated security=true"
        vSql = " "
        vSql = vSql & "Select TOP (1)  JobName, EndTime, ElapsedTimeSeconds, IsJobActive, FinalJobStatus, MediaSetName, "
        vSql = vSql & "TotalDataSizeBytes, TotalRateMBMin, TotalNumberOfFiles, "
        vSql = vSql & "TotalNumberOfDirectories, TotalSkippedFiles, TotalCorruptFiles, TotalInUseFiles "
        vSql = vSql & "From JobHistorySummary "
        vSql = vSql & "Where IsJobActive=0 "
        vSql = vSql & "Order By EndTime DESC "

        vSqlConn = New SqlConnection(vConnectionString)
        Try
            vSqlConn.Open()
            vSqlCmd = New SqlCommand(vSql, vSqlConn)
            Dim vSqlReader As SqlDataReader = vSqlCmd.ExecuteReader()
            While vSqlReader.Read()
                Console.Write("BKUPEXEC : ")
                Console.Write(vSqlReader.Item(0) & ", ")
                Console.Write(vSqlReader.Item(1) & ", ")
                Select Case vSqlReader.Item(4)
                    Case 3
                        Console.Write("SUCCESS WW, ")
                    Case 19
                        Console.Write("SUCCESS, ")
                    Case 6
                        Console.Write("FAILED, ")
                    Case Else
                        Console.Write("UNKNOWN, ")
                End Select
                Console.Write(Math.Round(vSqlReader.Item(2) / (3600), 2) & " hr, ")
                Console.WriteLine(Math.Round(vSqlReader.Item(6) / 1000000000, 2) & " GB")
                '                Console.WriteLine(vSqlReader.Item(7) & " MB/min ")
            End While
            vSqlReader.Close()
            vSqlCmd.Dispose()
            vSqlConn.Close()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub SynologyFreeSpace()
        Try
            For Each drive_info As DriveInfo In DriveInfo.GetDrives()
                If drive_info.Name = "X:\" Or drive_info.Name = "P:\" Then
                    If Math.Round(drive_info.TotalSize() / 1000000000, 2) > 1900 Then
                        Console.Write("Synology - " & drive_info.Name & ", ")
                        'Console.Write(drive_info.RootDirectory.ToString)
                        If drive_info.IsReady() Then
                            Console.Write("Total Space " & Math.Round(drive_info.TotalSize() / 1000000000, 2) & " GB ")
                            Console.WriteLine("Free Space " & Math.Round(drive_info.TotalFreeSpace() / 1000000000, 2) & " GB - around " & Math.Round(100 * drive_info.TotalFreeSpace() / drive_info.TotalSize(), 2) & " % ")
                        End If
                    End If
                End If

            Next drive_info
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub CalculateDownloadSpeed()
        If Downloading Then Exit Sub
        Downloading = True
        Dim FileToSave As String = "C:\temp\test.msu"
        Dim FileToDownload As String = ""
        'FileToDownload = "https://www.7-zip.org/a/7z1805-x64.exe"
        FileToDownload = "https://almitankers.gr/wp-content/uploads/2020/02/PA.jpg"
        If System.IO.File.Exists(FileToSave) = True Then
            System.IO.File.Delete(FileToSave)
        End If
        Dim wc As New WebClient
        Dim r As Net.WebRequest = Net.WebRequest.Create(FileToDownload)
        r.Method = Net.WebRequestMethods.Http.Head
        Using rsp = r.GetResponse()
            DownloadFileSize = rsp.ContentLength
        End Using
        AddHandler wc.DownloadProgressChanged, AddressOf wc_ProgressChanged
        AddHandler wc.DownloadFileCompleted, AddressOf wc_DownloadDone
        wc.DownloadFileAsync(New Uri(FileToDownload), FileToSave, Stopwatch.StartNew)
    End Sub

    Private Sub wc_DownloadDone(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs)
        Console.WriteLine("DownloadSpeed : " & Math.Round(DownloadSpeed / 1000, 2) & " Kbytes / second - > " & Math.Round(DownloadSpeed / 1000, 2) * 8 & " kbps ")
    End Sub

    Private Sub wc_ProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        DownloadSpeed = (e.BytesReceived / (DirectCast(e.UserState, Stopwatch).ElapsedMilliseconds / 1000.0#)).ToString("#")
        Console.Write(Math.Round(100 * e.BytesReceived / DownloadFileSize, 0) & " % " & vbCr)
    End Sub

End Module
