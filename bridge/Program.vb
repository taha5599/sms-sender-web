Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading

''' <summary>
''' SmsSenderBridge 1.0-web: a tiny local bridge so the SMS-sender
''' website can send through the USB modem on jony's PC.
'''
''' The website (Render, HTTPS) serves only static files and never sees
''' the numbers or the message: the browser page talks straight to this
''' bridge on http://127.0.0.1:18711/. The bridge reuses the desktop's
''' tested ModemSender.vb + PhoneNormalizer.vb + ExcelPhoneReader.vb
''' VERBATIM - no modem/PDU/counting logic is rewritten here.
'''
''' Multi-modem (Round 60 scope change): one send job distributes its
''' numbers round-robin across all chosen modems, one thread per modem,
''' each with its own ModemSender connection. A modem that dies mid-job
''' hands its remaining numbers to a single retry pass on the healthy
''' modems; with no healthy modem left those numbers fail loudly.
'''
''' The HTTP server is a minimal TcpListener implementation on purpose:
''' HttpListener on Windows needs a URL reservation (netsh / admin),
''' a raw socket on 127.0.0.1 needs none. It binds 127.0.0.1 ONLY.
'''
''' Safety: every /api/* call must carry the X-SmsBridge header and an
''' allowed Origin (localhost, or an origin listed in
''' bridge-origins.txt). A plain cross-site form post cannot set the
''' header, so no website jony visits can drive his modem.
''' </summary>
Public Module BridgeMain

    Public Const BridgeVersion As String = "1.0-web"
    Public Const BridgeHeaderName As String = "X-SmsBridge"
    Public Const BridgeHeaderValue As String = "sms-web-1.0"
    Public Const DefaultPort As Integer = 18711
    Private Const MaxBodyBytes As Integer = 50 * 1024 * 1024

    Private _listener As TcpListener
    Private _stopRequested As Boolean
    Private _port As Integer = DefaultPort
    Private _logPath As String = String.Empty
    Private _siteDir As String = String.Empty
    Private ReadOnly _logLock As New Object()
    Private ReadOnly _jobsLock As New Object()
    Private ReadOnly _jobs As New Dictionary(Of String, SendJob)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _allowedOrigins As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Public Sub Main()
        Dim args As String() = Environment.GetCommandLineArgs()
        Dim i As Integer = 0
        While i < args.Length
            If String.Equals(args(i), "--port", StringComparison.OrdinalIgnoreCase) AndAlso i + 1 < args.Length Then
                Dim p As Integer = 0
                If Integer.TryParse(args(i + 1), p) AndAlso p > 0 AndAlso p < 65536 Then _port = p
                i += 2
            Else
                i += 1
            End If
        End While
        LoadConfig()
        Dim appDir As String = AppContext.BaseDirectory
        _logPath = Path.Combine(appDir, "BridgeLog.txt")
        _siteDir = Path.Combine(appDir, "site")
        Log("SmsSenderBridge " & BridgeVersion & " starting on 127.0.0.1:" & _port.ToString())
        _listener = New TcpListener(IPAddress.Loopback, _port)
        Try
            _listener.Start()
        Catch ex As Exception
            Console.WriteLine("تعذر بدء الوسيط على المنفذ " & _port.ToString() & ": " & ex.Message)
            Return
        End Try
        Console.WriteLine("الوسيط شغال: http://127.0.0.1:" & _port.ToString() & "/")
        Console.WriteLine("للإيقاف اضغط Ctrl+C")
        AddHandler Console.CancelKeyPress, AddressOf OnCancelKeyPress
        While Not _stopRequested
            Try
                Dim client As TcpClient = _listener.AcceptTcpClient()
                ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf HandleClient), client)
            Catch ex As SocketException
                If _stopRequested Then Exit While
                Log("accept error: " & ex.Message)
            Catch ex As Exception
                If _stopRequested Then Exit While
                Log("accept error: " & ex.Message)
            End Try
        End While
        Try
            _listener.Stop()
        Catch ex As Exception
            Log("listener stop error: " & ex.Message)
        End Try
        Log("SmsSenderBridge stopped")
    End Sub

    Private Sub OnCancelKeyPress(sender As Object, e As ConsoleCancelEventArgs)
        e.Cancel = True
        _stopRequested = True
        Try
            If _listener IsNot Nothing Then _listener.Stop()
        Catch ex As Exception
            Log("listener stop error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>Reads bridge.ini (port=) and bridge-origins.txt (one origin per line, # = comment).</summary>
    Private Sub LoadConfig()
        Try
            Dim appDir As String = AppContext.BaseDirectory
            Dim iniPath As String = Path.Combine(appDir, "bridge.ini")
            If File.Exists(iniPath) Then
                For Each line As String In File.ReadAllLines(iniPath)
                    Dim t As String = line.Trim()
                    If t.StartsWith("port=", StringComparison.OrdinalIgnoreCase) Then
                        Dim p As Integer = 0
                        If Integer.TryParse(t.Substring("port=".Length).Trim(), p) AndAlso p > 0 AndAlso p < 65536 Then _port = p
                    End If
                Next
            End If
            Dim originsPath As String = Path.Combine(appDir, "bridge-origins.txt")
            If File.Exists(originsPath) Then
                For Each line As String In File.ReadAllLines(originsPath)
                    Dim t As String = line.Trim()
                    If t.Length = 0 OrElse t.StartsWith("#") Then Continue For
                    If Not _allowedOrigins.Contains(t) Then _allowedOrigins.Add(t)
                Next
            End If
        Catch ex As Exception
            Console.WriteLine("تعذر قراءة الإعدادات، سيتم استخدام الافتراضي: " & ex.Message)
        End Try
    End Sub

    ''' <summary>Appends a line to BridgeLog.txt next to the exe and to the console. Never throws.</summary>
    Public Sub Log(message As String)
        Try
            Dim line As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " " & message
            SyncLock _logLock
                If Not String.IsNullOrEmpty(_logPath) Then
                    File.AppendAllText(_logPath, line & vbCrLf, Encoding.UTF8)
                End If
            End SyncLock
            Console.WriteLine(line)
        Catch ex As Exception
            Try
                Console.WriteLine("log failed: " & ex.Message)
            Catch
            End Try
        End Try
    End Sub

    ' ------------------------------------------------------------------
    ' HTTP plumbing (minimal HTTP/1.1 over TcpListener)
    ' ------------------------------------------------------------------

    Public Class HttpRequest
        Public Property Method As String = String.Empty
        Public Property Path As String = String.Empty
        Public Property Query As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public Property Headers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public Property Body As Byte()

        Public ReadOnly Property BodyText As String
            Get
                If Body Is Nothing OrElse Body.Length = 0 Then Return String.Empty
                Return Encoding.UTF8.GetString(Body)
            End Get
        End Property

        Public Function Header(name As String) As String
            Dim value As String = Nothing
            If Headers.TryGetValue(name, value) Then Return value
            Return Nothing
        End Function

        Public Function QueryValue(name As String) As String
            Dim value As String = Nothing
            If Query.TryGetValue(name, value) Then Return value
            Return Nothing
        End Function
    End Class

    Public Class HttpResponse
        Public Property StatusCode As Integer = 200
        Public Property StatusText As String = "OK"
        Public Property ContentType As String = "application/json; charset=utf-8"
        Public Property Body As Byte()
        Public Property ExtraHeaders As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Class

    Private Sub HandleClient(state As Object)
        Dim client As TcpClient = DirectCast(state, TcpClient)
        Try
            Using client
                Using stream As NetworkStream = client.GetStream()
                    Dim req As HttpRequest = ReadRequest(stream)
                    If req Is Nothing Then Return
                    Dim resp As HttpResponse = Nothing
                    Try
                        resp = RouteRequest(req)
                    Catch ex As Exception
                        Log("handler error " & req.Method & " " & req.Path & ": " & ex.Message)
                        resp = ErrorResponse(500, "خطأ داخلي في الوسيط: " & ex.Message)
                    End Try
                    WriteResponse(stream, req, resp)
                End Using
            End Using
        Catch ex As Exception
            Try
                Log("client error: " & ex.Message)
            Catch
            End Try
        End Try
    End Sub

    Private Function ReadRequest(stream As NetworkStream) As HttpRequest
        Dim acc As New MemoryStream()
        Dim buf(8191) As Byte
        Dim headerEnd As Integer = -1
        Dim headerLen As Integer = 0
        Dim contentLength As Integer = 0
        Dim totalNeeded As Integer = -1
        Dim doubleCrlf As Byte() = Encoding.ASCII.GetBytes(vbCrLf & vbCrLf)
        While True
            Dim read As Integer = 0
            Try
                read = stream.Read(buf, 0, buf.Length)
            Catch ex As Exception
                Log("read error: " & ex.Message)
                Return Nothing
            End Try
            If read <= 0 Then
                If acc.Length = 0 Then Return Nothing
                Exit While
            End If
            acc.Write(buf, 0, read)
            If headerEnd < 0 Then
                Dim bytes As Byte() = acc.ToArray()
                headerEnd = FindBytes(bytes, doubleCrlf, 0)
                If headerEnd >= 0 Then
                    headerLen = headerEnd + doubleCrlf.Length
                    Dim headerText As String = Encoding.ASCII.GetString(bytes, 0, headerEnd)
                    contentLength = ParseContentLength(headerText)
                    If contentLength < 0 OrElse contentLength > MaxBodyBytes Then Return Nothing
                    totalNeeded = headerLen + contentLength
                ElseIf acc.Length > 65536 Then
                    Return Nothing
                End If
            End If
            If totalNeeded >= 0 AndAlso acc.Length >= totalNeeded Then Exit While
            If totalNeeded < 0 AndAlso acc.Length > 65536 Then Return Nothing
        End While
        Dim all As Byte() = acc.ToArray()
        If headerEnd < 0 Then Return Nothing
        Dim req As New HttpRequest()
        Dim headerText2 As String = Encoding.ASCII.GetString(all, 0, headerEnd)
        Dim lines As String() = headerText2.Split(New String() {vbCrLf}, StringSplitOptions.None)
        If lines.Length = 0 Then Return Nothing
        Dim requestLine As String() = lines(0).Split(" "c)
        If requestLine.Length < 2 Then Return Nothing
        req.Method = requestLine(0).Trim().ToUpperInvariant()
        Dim rawTarget As String = requestLine(1).Trim()
        Dim qIdx As Integer = rawTarget.IndexOf("?"c)
        If qIdx >= 0 Then
            req.Path = rawTarget.Substring(0, qIdx)
            ParseQuery(rawTarget.Substring(qIdx + 1), req.Query)
        Else
            req.Path = rawTarget
        End If
        req.Path = Uri.UnescapeDataString(req.Path)
        For li As Integer = 1 To lines.Length - 1
            Dim colon As Integer = lines(li).IndexOf(":"c)
            If colon > 0 Then
                req.Headers(lines(li).Substring(0, colon).Trim()) = lines(li).Substring(colon + 1).Trim()
            End If
        Next
        Dim bodyLen As Integer = Math.Min(contentLength, all.Length - headerLen)
        If bodyLen > 0 Then
            Dim body(bodyLen - 1) As Byte
            Array.Copy(all, headerLen, body, 0, bodyLen)
            req.Body = body
        Else
            req.Body = New Byte() {}
        End If
        Return req
    End Function

    Private Function ParseContentLength(headerText As String) As Integer
        For Each line As String In headerText.Split(New String() {vbCrLf}, StringSplitOptions.None)
            Dim colon As Integer = line.IndexOf(":"c)
            If colon > 0 AndAlso String.Equals(line.Substring(0, colon).Trim(), "Content-Length", StringComparison.OrdinalIgnoreCase) Then
                Dim n As Integer = 0
                If Integer.TryParse(line.Substring(colon + 1).Trim(), n) Then Return n
                Return -1
            End If
        Next
        Return 0
    End Function

    Private Sub ParseQuery(query As String, target As Dictionary(Of String, String))
        If String.IsNullOrEmpty(query) Then Return
        For Each pair As String In query.Split("&"c)
            If pair.Length = 0 Then Continue For
            Dim eq As Integer = pair.IndexOf("="c)
            If eq >= 0 Then
                target(Uri.UnescapeDataString(pair.Substring(0, eq).Replace("+"c, " "c))) = Uri.UnescapeDataString(pair.Substring(eq + 1).Replace("+"c, " "c))
            Else
                target(Uri.UnescapeDataString(pair.Replace("+"c, " "c))) = String.Empty
            End If
        Next
    End Sub

    Private Sub WriteResponse(stream As NetworkStream, req As HttpRequest, resp As HttpResponse)
        Dim body As Byte() = resp.Body
        If body Is Nothing Then body = New Byte() {}
        Dim sb As New StringBuilder()
        sb.Append("HTTP/1.1 ").Append(resp.StatusCode).Append(" ").Append(resp.StatusText).Append(vbCrLf)
        sb.Append("Content-Type: ").Append(resp.ContentType).Append(vbCrLf)
        sb.Append("Content-Length: ").Append(body.Length).Append(vbCrLf)
        sb.Append("Connection: close").Append(vbCrLf)
        For Each kvp As KeyValuePair(Of String, String) In resp.ExtraHeaders
            sb.Append(kvp.Key).Append(": ").Append(kvp.Value).Append(vbCrLf)
        Next
        sb.Append(vbCrLf)
        Dim headerBytes As Byte() = Encoding.ASCII.GetBytes(sb.ToString())
        stream.Write(headerBytes, 0, headerBytes.Length)
        If body.Length > 0 Then stream.Write(body, 0, body.Length)
        stream.Flush()
    End Sub

    ' ------------------------------------------------------------------
    ' CORS + bridge-header gate
    ' ------------------------------------------------------------------

    ''' <summary>True for origins the bridge answers: no Origin (curl/direct), localhost, or bridge-origins.txt.
    ''' Origin "null" (file://, data:, sandboxed iframes) is NOT allowed: the X-SmsBridge header value is public
    ''' (it sits in site/app.js and the README), so the Origin allowlist is the only protection. (Round 61) B-60-H1.
    ''' file:// testing keeps working through the bridge-served same-origin site (http://127.0.0.1:18711/).</summary>
    Private Function IsOriginAllowed(origin As String) As Boolean
        If String.IsNullOrEmpty(origin) Then Return True
        Dim u As Uri = Nothing
        If Not Uri.TryCreate(origin, UriKind.Absolute, u) Then Return False
        ' (Round 62) A-61-M1: check the scheme, not just the host - file://localhost, ftp://localhost and
        ' ws://localhost all have a localhost host but are not web origins; only http/https may use it.
        If Not (u.Scheme = "http" OrElse u.Scheme = "https") Then Return False
        Dim host As String = u.Host.ToLowerInvariant()
        If host = "localhost" OrElse host = "127.0.0.1" OrElse host = "[::1]" OrElse host = "::1" Then Return True
        For Each allowed As String In _allowedOrigins
            If String.Equals(allowed, origin, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Adds the CORS headers for an allowed origin, including the Private Network Access grant Chrome requires.</summary>
    Private Sub AddCorsHeaders(req As HttpRequest, resp As HttpResponse)
        Dim origin As String = req.Header("Origin")
        If String.IsNullOrEmpty(origin) Then Return
        ' (Round 62) A-61-L1: never echo a literal "null" origin - the request is rejected anyway and the echo
        ' contradicts the B-60-H1 fix. Real origins still get the echo (A-60-L2 readable 403).
        If String.Equals(origin, "null", StringComparison.OrdinalIgnoreCase) Then Return
        ' (Round 61) B-60-H1: no "null" branch any more - a null Origin is
        ' rejected by IsOriginAllowed, so there is nothing to echo for it.
        resp.ExtraHeaders("Access-Control-Allow-Origin") = origin
        resp.ExtraHeaders("Vary") = "Origin"
        Dim pna As String = req.Header("Access-Control-Request-Private-Network")
        If String.Equals(pna, "true", StringComparison.OrdinalIgnoreCase) Then
            resp.ExtraHeaders("Access-Control-Allow-Private-Network") = "true"
        End If
    End Sub

    Friend Function JsonResponse(status As Integer, json As String) As HttpResponse
        Dim resp As New HttpResponse()
        resp.StatusCode = status
        resp.StatusText = If(status = 200, "OK", If(status = 400, "Bad Request", If(status = 403, "Forbidden", If(status = 404, "Not Found", If(status = 409, "Conflict", "Error")))))
        resp.ContentType = "application/json; charset=utf-8"
        resp.Body = New UTF8Encoding(False).GetBytes(json)
        Return resp
    End Function

    Friend Function ErrorResponse(status As Integer, message As String) As HttpResponse
        Return JsonResponse(status, "{""ok"":false,""error"":" & JsonStr(message) & "}")
    End Function

    ' ------------------------------------------------------------------
    ' Routing
    ' ------------------------------------------------------------------

    Private Function RouteRequest(req As HttpRequest) As HttpResponse
        Log(req.Method & " " & req.Path)
        If req.Path.StartsWith("/api/", StringComparison.OrdinalIgnoreCase) Then
            If String.Equals(req.Method, "OPTIONS", StringComparison.OrdinalIgnoreCase) Then
                Return HandlePreflight(req)
            End If
            If Not IsOriginAllowed(req.Header("Origin")) Then
                ' (Round 61) A-60-L2: the CORS headers go on the gate-failure
                ' response too, so the page can READ the 403 body ("origin not
                ' allowed - ...") instead of seeing a generic fetch TypeError
                ' and misdiagnosing it as "bridge down". Echoing the Origin
                ' grants nothing: the call itself is still rejected (403) and
                ' the error body carries no secrets.
                Dim resp403a As HttpResponse = ErrorResponse(403, "origin not allowed - أضف رابط الموقع إلى bridge-origins.txt")
                AddCorsHeaders(req, resp403a)
                Return resp403a
            End If
            Dim hdr As String = req.Header(BridgeHeaderName)
            If Not String.Equals(hdr, BridgeHeaderValue, StringComparison.Ordinal) Then
                Dim resp403b As HttpResponse = ErrorResponse(403, "missing bridge header")
                AddCorsHeaders(req, resp403b)
                Return resp403b
            End If
            Dim resp As HttpResponse = DispatchApi(req)
            AddCorsHeaders(req, resp)
            Return resp
        End If
        If Not String.Equals(req.Method, "GET", StringComparison.OrdinalIgnoreCase) Then
            Return ErrorResponse(405, "method not allowed")
        End If
        Return ServeStatic(req)
    End Function

    Private Function HandlePreflight(req As HttpRequest) As HttpResponse
        Dim resp As New HttpResponse()
        resp.StatusCode = 204
        resp.StatusText = "No Content"
        resp.ContentType = "text/plain; charset=utf-8"
        resp.Body = New Byte() {}
        If Not IsOriginAllowed(req.Header("Origin")) Then
            Return resp
        End If
        AddCorsHeaders(req, resp)
        resp.ExtraHeaders("Access-Control-Allow-Methods") = "GET, POST, OPTIONS"
        resp.ExtraHeaders("Access-Control-Allow-Headers") = "X-SmsBridge, Content-Type"
        resp.ExtraHeaders("Access-Control-Max-Age") = "86400"
        Return resp
    End Function

    Private Function DispatchApi(req As HttpRequest) As HttpResponse
        Dim path As String = req.Path.ToLowerInvariant()
        Select Case req.Method & " " & path
            Case "GET /api/status"
                Return BridgeApi.ApiStatus(req)
            Case "GET /api/inbox"
                Return BridgeApi.ApiInbox(req)
            Case "GET /api/ports"
                Return BridgeApi.ApiPorts(req)
            Case "POST /api/detect"
                Return BridgeApi.ApiDetect(req)
            Case "POST /api/test"
                Return BridgeApi.ApiTest(req)
            Case "POST /api/preview"
                Return BridgeApi.ApiPreview(req)
            Case "POST /api/send"
                Return BridgeApi.ApiSend(req)
            Case "GET /api/progress"
                Return BridgeApi.ApiProgress(req)
            Case "POST /api/stop"
                Return BridgeApi.ApiStop(req)
            Case Else
                Return ErrorResponse(404, "not found")
        End Select
    End Function

    ''' <summary>Serves the website files from the site/ folder next to the exe (same-origin fallback for any browser, incl. Safari).</summary>
    Private Function ServeStatic(req As HttpRequest) As HttpResponse
        If String.IsNullOrEmpty(_siteDir) OrElse Not Directory.Exists(_siteDir) Then
            Return ErrorResponse(404, "site not found - ضع ملفات الموقع في مجلد site بجانب الوسيط")
        End If
        Dim rel As String = req.Path.TrimStart("/"c)
        If String.IsNullOrEmpty(rel) Then rel = "index.html"
        rel = rel.Replace("/"c, Path.DirectorySeparatorChar)
        If rel.Contains("..") OrElse rel.Contains(":") OrElse Path.IsPathRooted(rel) Then
            Return ErrorResponse(403, "forbidden")
        End If
        Dim full As String = Path.Combine(_siteDir, rel)
        If Not File.Exists(full) Then
            Return ErrorResponse(404, "not found")
        End If
        Dim resp As New HttpResponse()
        resp.StatusCode = 200
        resp.StatusText = "OK"
        Dim ext As String = Path.GetExtension(full).ToLowerInvariant()
        Select Case ext
            Case ".html", ".htm"
                resp.ContentType = "text/html; charset=utf-8"
            Case ".js"
                resp.ContentType = "text/javascript; charset=utf-8"
            Case ".css"
                resp.ContentType = "text/css; charset=utf-8"
            Case ".png"
                resp.ContentType = "image/png"
            Case ".ico"
                resp.ContentType = "image/x-icon"
            Case Else
                resp.ContentType = "application/octet-stream"
        End Select
        Try
            resp.Body = File.ReadAllBytes(full)
        Catch ex As Exception
            Return ErrorResponse(500, "تعذر قراءة الملف: " & ex.Message)
        End Try
        Return resp
    End Function

    Friend Function GetJob(jobId As String) As SendJob
        If String.IsNullOrEmpty(jobId) Then Return Nothing
        SyncLock _jobsLock
            Dim job As SendJob = Nothing
            If _jobs.TryGetValue(jobId, job) Then Return job
            Return Nothing
        End SyncLock
    End Function

    Friend Function HasActiveJob() As Boolean
        SyncLock _jobsLock
            For Each job As SendJob In _jobs.Values
                If Not job.Done Then Return True
            Next
            Return False
        End SyncLock
    End Function

    Friend Sub RegisterJob(job As SendJob)
        SyncLock _jobsLock
            _jobs(job.JobId) = job
        End SyncLock
    End Sub

End Module
