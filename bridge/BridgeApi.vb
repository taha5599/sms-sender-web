Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Threading

''' <summary>
''' API handlers for SmsSenderBridge. All handlers use the desktop's
''' tested ModemSender.vb / PhoneNormalizer.vb / ExcelPhoneReader.vb
''' VERBATIM - no modem/PDU/counting logic is rewritten here.
''' </summary>
Public Module BridgeApi

    ' ------------------------------------------------------------------
    ' Capped port enumeration (Round 61) A-60-M1
    ' ------------------------------------------------------------------

    Private ReadOnly _enumLock As New Object()
    Private _enumThread As Thread

    ''' <summary>Enumerates COM ports on a dedicated background thread with a 15-second Join cap (ported from the
    ''' desktop's EnumeratePortsCapped, Round 53): a hung WMI cannot hang /api/ports, /api/detect, /api/test(all) or
    ''' the send pre-flight. Returns Nothing and sets timedOut True on a Join timeout, or while a previous
    ''' enumeration is still alive (refused instead of leaking another hung thread). wmiError keeps its Round 14/17
    ''' meaning; a throw out of GetAvailablePortsWithWmiError itself is recorded separately in enumError.</summary>
    Private Function EnumeratePortsCapped(ByRef wmiError As Exception, ByRef enumError As Exception, ByRef timedOut As Boolean) As List(Of ModemSender.PortInfo)
        wmiError = Nothing
        enumError = Nothing
        timedOut = False
        SyncLock _enumLock
            Dim outstanding As Thread = _enumThread
            If outstanding IsNot Nothing AndAlso outstanding.IsAlive Then
                timedOut = True
                wmiError = New Exception("تعداد المنافذ ما زال جارياً في الخلفية — أعد المحاولة بعد قليل")
                Return Nothing
            End If
        End SyncLock
        Dim portsResult As List(Of ModemSender.PortInfo) = Nothing
        ' (Round 54 desktop lesson) BC36639: ByRef parameters may NOT be used inside a lambda - the worker thread
        ' reads/writes LOCALS only; after a successful Join the locals are copied back. On a Join timeout the locals
        ' are NEVER read (the abandoned thread may still be writing them) and the ByRef params stay Nothing.
        Dim wmiErrorLocal As Exception = Nothing
        Dim enumErrorLocal As Exception = Nothing
        Dim wmiThread As New Thread(Sub()
                                        Try
                                            portsResult = ModemSender.GetAvailablePortsWithWmiError(wmiErrorLocal)
                                        Catch ex As Exception
                                            enumErrorLocal = ex
                                        End Try
                                    End Sub)
        wmiThread.IsBackground = True
        SyncLock _enumLock
            _enumThread = wmiThread
        End SyncLock
        wmiThread.Start()
        If Not wmiThread.Join(TimeSpan.FromSeconds(15)) Then
            timedOut = True
            wmiError = New Exception("انتهت المهلة أثناء تعداد المنافذ (15 ثانية) — قد يكون WMI معلّقاً")
            Return Nothing
        End If
        wmiError = wmiErrorLocal
        enumError = enumErrorLocal
        Return portsResult
    End Function

    ' ------------------------------------------------------------------
    ' Single send-slot claim (Round 61) C-60-M2 / B-60-M1
    ' ------------------------------------------------------------------

    Private ReadOnly _sendClaimLock As New Object()
    Private _sendClaimed As Boolean

    ''' <summary>Atomically claims the single send slot. The check (no active job) and the claim must be one locked
    ''' step: ApiSend used to check HasActiveJob, then probe ports for seconds, then RegisterJob - two concurrent
    ''' sends both passed the check and sent the same numbers twice. Returns False when a send is already running
    ''' or claimed. The claim is released by ApiSend's Finally on EVERY return (Round 62 B-61-H1/C-61-H1: the old
    ''' "release only before registration" left it held forever after a successful send, so every later send got
    ''' a stuck 409); once a job is registered, HasActiveJob guards the running job instead.</summary>
    Friend Function TryClaimSendSlot() As Boolean
        SyncLock _sendClaimLock
            If _sendClaimed OrElse BridgeMain.HasActiveJob() Then Return False
            _sendClaimed = True
            Return True
        End SyncLock
    End Function

    Friend Sub ReleaseSendSlot()
        SyncLock _sendClaimLock
            _sendClaimed = False
        End SyncLock
    End Sub


    ' (R75W) Read-only inbox: a fresh AT+CMGL=4 read on every call, nothing cached, nothing deleted/marked read.
    ' (Round 76W) D2/B-L1: the inbox read claims the single send slot for its whole duration (probe + open +
    ' CMGL), so a concurrent send gets an honest 409 instead of failing 400 on a port this read holds. The
    ' claim is released in Finally on every return, as ApiSend does (Round 62). TryClaimSendSlot also refuses
    ' while a send job is active, so it subsumes the old HasActiveJob check.
    Public Function ApiInbox(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        If Not TryClaimSendSlot() Then
            Return ErrorResponse(409, "يوجد إرسال أو قراءة وارد جارية — أعد المحاولة بعد انتهائها")
        End If
        Try
            Return ApiInboxCore(req)
        Finally
            ReleaseSendSlot()
        End Try
    End Function

    Private Function ApiInboxCore(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim portName As String = req.QueryValue("port")
        Dim requested As List(Of String) = Nothing
        If Not String.IsNullOrEmpty(portName) Then
            requested = New List(Of String)()
            requested.Add(portName.Trim())
        End If
        Dim outcome As ProbeOutcome = ProbePorts(requested)
        Dim target As ModemSender.PortInfo = Nothing
        Try
            target = ModemSender.ChooseProbedPort(outcome.Results)
        Catch ex As Exception
            BridgeMain.Log("choose inbox port error: " & ex.Message)
        End Try
        If target Is Nothing AndAlso outcome.Usable.Count > 0 Then target = outcome.Usable(0)
        If target Is Nothing Then
            If requested IsNot Nothing AndAlso outcome.Excluded.Count > 0 Then
                Return ErrorResponse(400, "المنفذ " & portName & " غير متاح — " & outcome.Excluded(0).Reason)
            End If
            Dim errText As String = If(outcome.TimedOut, "تعذر تعداد المنافذ: " & outcome.WmiError, "لا يوجد مودم سليم")
            Return ErrorResponse(400, errText)
        End If
        Dim messages As List(Of ModemSender.InboxMessage) = Nothing
        Dim inboxError As String = String.Empty
        Dim inboxNotes As String = String.Empty ' (R82W) per-storage counts from the last ReadInboxSms call
        Try
            Using modem As New ModemSender()
                Dim openOk As Boolean = False
                Try
                    openOk = modem.Open(target.PortName, Nothing)
                Catch ex As Exception
                    Return ErrorResponse(500, "تعذر فتح المنفذ " & target.PortName & ": " & ex.Message)
                End Try
                If Not openOk Then
                    Return ErrorResponse(500, "فشل فتح المنفذ " & target.PortName & " (المودم لم يرد على أوامر AT/ATE0)")
                End If
                messages = modem.ReadInboxSms(inboxError, Nothing)
                inboxNotes = modem.InboxReadNotes
            End Using
        Catch ex As Exception
            Return ErrorResponse(500, "خطأ أثناء قراءة الوارد: " & ex.Message)
        End Try
        If messages Is Nothing Then
            Return ErrorResponse(500, inboxError)
        End If
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,""port"":").Append(JsonStr(target.PortName)).Append(",""messages"":[")
        Dim first As Boolean = True
        For Each m As ModemSender.InboxMessage In messages
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append("""index"":").Append(m.Index).Append(",")
            sb.Append("""status"":").Append(JsonStr(m.Status)).Append(",")
            sb.Append("""from"":").Append(JsonStr(m.FromNumber)).Append(",")
            sb.Append("""time"":").Append(JsonStr(m.TimeText)).Append(",")
            sb.Append("""text"":").Append(JsonStr(m.Text)).Append(",")
            ' (R82W) Storage tag: the storage this message came from (SM/ME/MT; empty on the CPMS-fallback read).
            sb.Append("""storage"":").Append(JsonStr(m.Storage))
            sb.Append("}")
        Next
        ' (R82W) Per-storage counts/notes (empty when the read fell back to the single default storage).
        sb.Append("],").Append("""notes"":").Append(JsonStr(inboxNotes)).Append("}")
        Return JsonResponse(200, sb.ToString())
    End Function

    Public Function ApiStatus(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Return JsonResponse(200, "{""ok"":true,""version"":" & JsonStr(BridgeMain.BridgeVersion) & ",""bridge"":""SmsSenderBridge""}")
    End Function

    Public Function ApiPorts(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim wmiError As Exception = Nothing
        Dim enumError As Exception = Nothing
        Dim enumTimedOut As Boolean = False
        Dim ports As List(Of ModemSender.PortInfo) = Nothing
        Try
            ' (Round 61) A-60-M1: capped enumeration - a hung WMI returns a timeout error instead of hanging the call.
            ports = EnumeratePortsCapped(wmiError, enumError, enumTimedOut)
        Catch ex As Exception
            Return ErrorResponse(500, "تعذر تعداد المنافذ: " & ex.Message)
        End Try
        If enumTimedOut Then
            Return ErrorResponse(500, "تعذر تعداد المنافذ: " & wmiError.Message)
        End If
        If enumError IsNot Nothing Then
            Return ErrorResponse(500, "تعذر تعداد المنافذ: " & enumError.Message)
        End If
        If ports Is Nothing Then ports = New List(Of ModemSender.PortInfo)()
        Dim preferred As ModemSender.PortInfo = Nothing
        Try
            preferred = ModemSender.SelectPreferredSmsPort(ports)
        Catch ex As Exception
            BridgeMain.Log("preferred port error: " & ex.Message)
        End Try
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,""ports"":[")
        Dim first As Boolean = True
        For Each p As ModemSender.PortInfo In ports
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append("""port"":").Append(JsonStr(p.PortName)).Append(",")
            sb.Append("""name"":").Append(JsonStr(p.Name)).Append(",")
            sb.Append("""description"":").Append(JsonStr(p.Description)).Append(",")
            sb.Append("""display"":").Append(JsonStr(p.DisplayText)).Append(",")
            sb.Append("""isZte"":").Append(JsonBool(p.IsZteDiagnostics)).Append(",")
            sb.Append("""isHuaweiSms"":").Append(JsonBool(p.IsHuaweiSmsPort)).Append(",")
            sb.Append("""isHuaweiData"":").Append(JsonBool(p.IsHuaweiDataPort)).Append(",")
            sb.Append("""isHuaweiModem"":").Append(JsonBool(p.IsHuaweiModemPort)).Append(",")
            sb.Append("""isNmea"":").Append(JsonBool(p.IsNmeaGps)).Append(",")
            sb.Append("""preferred"":").Append(JsonBool(preferred IsNot Nothing AndAlso String.Equals(preferred.PortName, p.PortName, StringComparison.OrdinalIgnoreCase)))
            sb.Append("}")
        Next
        sb.Append("],""wmiError"":")
        If wmiError IsNot Nothing Then sb.Append(JsonStr(wmiError.Message)) Else sb.Append("null")
        sb.Append("}")
        Return JsonResponse(200, sb.ToString())
    End Function

    Public Function ApiDetect(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim outcome As ProbeOutcome = ProbePorts(Nothing)
        Dim preferred As ModemSender.PortInfo = Nothing
        Try
            preferred = ModemSender.ChooseProbedPort(outcome.Results)
        Catch ex As Exception
            BridgeMain.Log("choose probed port error: " & ex.Message)
        End Try
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,""modems"":[")
        Dim first As Boolean = True
        For Each u As ModemSender.PortInfo In outcome.Usable
            If Not first Then sb.Append(",")
            first = False
            sb.Append(JsonStr(u.PortName))
        Next
        sb.Append("],""preferred"":")
        If preferred IsNot Nothing Then sb.Append(JsonStr(preferred.PortName)) Else sb.Append("null")
        sb.Append(",""wmiError"":")
        If String.IsNullOrEmpty(outcome.WmiError) Then sb.Append("null") Else sb.Append(JsonStr(outcome.WmiError))
        sb.Append(",""results"":[")
        first = True
        For Each r As ModemSender.PortProbeResult In outcome.Results
            If r Is Nothing OrElse r.Port Is Nothing Then Continue For
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append("""port"":").Append(JsonStr(r.Port.PortName)).Append(",")
            sb.Append("""kind"":").Append(JsonStr(r.Kind.ToString())).Append(",")
            sb.Append("""label"":").Append(JsonStr(r.Label)).Append(",")
            sb.Append("""detail"":").Append(JsonStr(r.Detail)).Append(",")
            sb.Append("""isZte"":").Append(JsonBool(r.Port.IsZteDiagnostics)).Append(",")
            sb.Append("""isHuaweiSms"":").Append(JsonBool(r.Port.IsHuaweiSmsPort)).Append(",")
            sb.Append("""isHuaweiData"":").Append(JsonBool(r.Port.IsHuaweiDataPort)).Append(",")
            sb.Append("""isHuaweiModem"":").Append(JsonBool(r.Port.IsHuaweiModemPort)).Append(",")
            sb.Append("""isNmea"":").Append(JsonBool(r.Port.IsNmeaGps)).Append(",")
            sb.Append("""display"":").Append(JsonStr(r.Port.DisplayText)).Append(",")
            sb.Append("""reason"":").Append(JsonStr(DetectReason(r)))
            sb.Append("}")
        Next
        sb.Append("]}")
        Return JsonResponse(200, sb.ToString())
    End Function

    Public Class TestPortResult
        Public Property Port As String = String.Empty
        Public Property Ok As Boolean
        Public Property AtOk As Boolean
        Public Property Sim As String = String.Empty
        Public Property Signal As String = String.Empty
        Public Property [Error] As String = String.Empty
    End Class

    Private Function TestOnePort(portName As String) As TestPortResult
        Dim r As New TestPortResult() With {.Port = portName}
        Try
            Using modem As New ModemSender()
                Dim openOk As Boolean = False
                Try
                    openOk = modem.Open(portName, Nothing)
                Catch ex As Exception
                    r.Error = "تعذر فتح المنفذ " & portName & ": " & ex.Message
                    Return r
                End Try
                If Not openOk Then
                    r.Error = "فشل فتح المنفذ " & portName & " (المودم لم يرد على أوامر AT/ATE0)"
                    Return r
                End If
                r.AtOk = modem.TestConnection(Nothing)
                r.Sim = modem.CheckSim(Nothing)
                r.Signal = modem.GetSignal(Nothing)
                r.Ok = r.AtOk
            End Using
        Catch ex As Exception
            r.Error = "خطأ أثناء الاختبار: " & ex.Message
        End Try
        Return r
    End Function

    Private Function TestPortResultJson(r As TestPortResult) As String
        Dim sb As New StringBuilder()
        sb.Append("{")
        sb.Append("""port"":").Append(JsonStr(r.Port)).Append(",")
        sb.Append("""ok"":").Append(JsonBool(r.Ok)).Append(",")
        sb.Append("""atOk"":").Append(JsonBool(r.AtOk)).Append(",")
        sb.Append("""sim"":").Append(JsonStr(r.Sim)).Append(",")
        sb.Append("""signal"":").Append(JsonStr(r.Signal)).Append(",")
        sb.Append("""error"":").Append(JsonStr(r.Error))
        sb.Append("}")
        Return sb.ToString()
    End Function

    Public Function ApiTest(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim body As String = req.BodyText
        Dim singlePort As String = JsonGetString(body, "port")
        Dim portsField As String = JsonGetString(body, "ports")
        Dim targets As New List(Of String)()
        If Not String.IsNullOrEmpty(portsField) Then
            For Each t As String In portsField.Split(","c)
                Dim n As String = t.Trim()
                If n.Length > 0 Then targets.Add(n)
            Next
        ElseIf Not String.IsNullOrEmpty(singlePort) Then
            If String.Equals(singlePort, "all", StringComparison.OrdinalIgnoreCase) Then
                ' (Round 61) A-60-M1: capped enumeration - a hung WMI returns a timeout error instead of hanging the call.
                Dim wmiError2 As Exception = Nothing
                Dim enumError2 As Exception = Nothing
                Dim enumTimedOut2 As Boolean = False
                Dim allPorts As List(Of ModemSender.PortInfo) = Nothing
                Try
                    allPorts = EnumeratePortsCapped(wmiError2, enumError2, enumTimedOut2)
                Catch ex As Exception
                    Return ErrorResponse(500, "تعذر تعداد المنافذ: " & ex.Message)
                End Try
                If enumTimedOut2 Then
                    Return ErrorResponse(500, "تعذر تعداد المنافذ: " & wmiError2.Message)
                End If
                If enumError2 IsNot Nothing Then
                    Return ErrorResponse(500, "تعذر تعداد المنافذ: " & enumError2.Message)
                End If
                If allPorts IsNot Nothing Then
                    For Each p As ModemSender.PortInfo In allPorts
                        targets.Add(p.PortName)
                    Next
                End If
            Else
                targets.Add(singlePort)
            End If
        Else
            Return ErrorResponse(400, "حدد المنفذ (port) أو all")
        End If
        If targets.Count = 0 Then
            Return JsonResponse(200, "{""ok"":true,""results"":[]}")
        End If
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,""results"":[")
        Dim first As Boolean = True
        Dim firstResult As TestPortResult = Nothing
        For Each target As String In targets
            Dim r As TestPortResult = TestOnePort(target)
            If firstResult Is Nothing Then firstResult = r
            If Not first Then sb.Append(",")
            first = False
            sb.Append(TestPortResultJson(r))
        Next
        sb.Append("]")
        If targets.Count = 1 AndAlso firstResult IsNot Nothing Then
            sb.Append(",""atOk"":").Append(JsonBool(firstResult.AtOk))
            sb.Append(",""sim"":").Append(JsonStr(firstResult.Sim))
            sb.Append(",""signal"":").Append(JsonStr(firstResult.Signal))
            sb.Append(",""error"":").Append(JsonStr(firstResult.Error))
        End If
        sb.Append("}")
        Return JsonResponse(200, sb.ToString())
    End Function

    ''' <summary>Loads the uploaded Excel/CSV file with the desktop's reader. Returns Nothing + loadError on failure.</summary>
    Private Function LoadPhonesFromFile(f As MultipartFile, ByRef loadError As String) As ExcelPhoneReader.LoadResult
        loadError = String.Empty
        If f Is Nothing OrElse f.Data Is Nothing Then
            loadError = "لم يصل ملف"
            Return Nothing
        End If
        Dim ext As String = Path.GetExtension(f.FileName).ToLowerInvariant()
        If ext.Length = 0 Then ext = ".xlsx"
        If ext <> ".xlsx" AndAlso ext <> ".xlsm" AndAlso ext <> ".csv" Then
            loadError = "صيغة الملف غير مدعومة. استخدم .xlsx أو .xlsm أو .csv"
            Return Nothing
        End If
        Dim tempPath As String = Path.Combine(Path.GetTempPath(), "smsweb_" & Guid.NewGuid().ToString("N") & ext)
        Try
            File.WriteAllBytes(tempPath, f.Data)
        Catch ex As Exception
            loadError = "تعذر حفظ الملف المؤقت: " & ex.Message
            Return Nothing
        End Try
        Try
            Dim reader As New ExcelPhoneReader()
            Dim lr As ExcelPhoneReader.LoadResult = reader.LoadPhones(tempPath)
            If lr.HasError Then
                loadError = lr.ErrorMessage
                Return Nothing
            End If
            Return lr
        Catch ex As Exception
            loadError = "تعذر قراءة الملف: " & ex.Message
            Return Nothing
        Finally
            Try
                File.Delete(tempPath)
            Catch ex As Exception
                BridgeMain.Log("temp delete failed: " & ex.Message)
            End Try
        End Try
    End Function

    Private Function RequireMultipartFile(req As BridgeMain.HttpRequest, ByRef mp As MultipartData) As String
        Dim contentType As String = req.Header("Content-Type")
        If String.IsNullOrEmpty(contentType) OrElse contentType.IndexOf("multipart/form-data", StringComparison.OrdinalIgnoreCase) < 0 Then
            Return "اطلب multipart/form-data مع ملف باسم file"
        End If
        mp = ParseMultipart(req.Body, contentType)
        If mp Is Nothing OrElse Not mp.Files.ContainsKey("file") Then
            Return "لم يصل ملف — أرفق ملف .xlsx أو .xlsm أو .csv باسم file"
        End If
        Return Nothing
    End Function

    Public Function ApiPreview(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim mp As MultipartData = Nothing
        Dim err As String = RequireMultipartFile(req, mp)
        If err IsNot Nothing Then Return ErrorResponse(400, err)
        Dim loadError As String = String.Empty
        Dim lr As ExcelPhoneReader.LoadResult = LoadPhonesFromFile(mp.Files("file"), loadError)
        If lr Is Nothing Then Return ErrorResponse(400, loadError)
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,")
        sb.Append("""total"":").Append(lr.TotalFound).Append(",")
        sb.Append("""valid"":").Append(lr.ValidUnique.Count).Append(",")
        sb.Append("""duplicates"":").Append(lr.DuplicatesRemoved).Append(",")
        sb.Append("""invalid"":").Append(lr.InvalidSkipped).Append(",")
        sb.Append("""sheet"":").Append(JsonStr(lr.SheetName)).Append(",")
        sb.Append("""numbers"":[")
        Dim first As Boolean = True
        For Each n As String In lr.ValidUnique
            If Not first Then sb.Append(",")
            first = False
            sb.Append(JsonStr(n))
        Next
        sb.Append("]}")
        Return JsonResponse(200, sb.ToString())
    End Function

    Public Function ApiSend(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim mp As MultipartData = Nothing
        Dim err As String = RequireMultipartFile(req, mp)
        If err IsNot Nothing Then Return ErrorResponse(400, err)
        ' (Round 61) C-60-M2/B-60-M1: claim the single send slot BEFORE any slow work (file parse, probe). The old
        ' HasActiveJob check sat after the file parse and seconds before RegisterJob, so two concurrent sends both
        ' passed it and sent the same numbers twice. The Finally below releases the claim on EVERY return from
        ' ApiSend (Round 62) B-61-H1/C-61-H1; once registered, HasActiveJob guards the running job. (The whole body
        ' below runs inside this Try, so its indentation is one level shallower than the Try - cosmetic only.)
        If Not TryClaimSendSlot() Then
            Return ErrorResponse(409, "يوجد إرسال أو قراءة وارد جارية — أعد المحاولة بعد انتهائها")
        End If
        Try
        Dim message As String = Nothing
        If mp.Fields.TryGetValue("message", message) Then message = message.Trim()
        If String.IsNullOrEmpty(message) Then
            Return ErrorResponse(400, "الرسالة فارغة — اكتب نص الرسالة")
        End If
        Dim delaySeconds As Integer = 5
        Dim delayText As String = Nothing
        If mp.Fields.TryGetValue("delaySeconds", delayText) Then
            Dim d As Integer = 0
            If Integer.TryParse(delayText, d) Then delaySeconds = d
        End If
        If delaySeconds < 2 Then delaySeconds = 2
        ' (R76b) No product cap anymore: the user types any delay. 86400
        ' (24h) is only the overflow guard — DelaySeconds * 1000 must fit
        ' in an Integer for SleepDelay.
        If delaySeconds > 86400 Then delaySeconds = 86400
        Dim portsText As String = Nothing
        mp.Fields.TryGetValue("ports", portsText)
        Dim requested As List(Of String) = Nothing
        If Not String.IsNullOrEmpty(portsText) AndAlso Not String.Equals(portsText.Trim(), "all", StringComparison.OrdinalIgnoreCase) Then
            requested = New List(Of String)()
            For Each t As String In portsText.Split(","c)
                Dim n As String = t.Trim()
                If n.Length > 0 Then requested.Add(n)
            Next
        End If
        Dim loadError As String = String.Empty
        Dim lr As ExcelPhoneReader.LoadResult = LoadPhonesFromFile(mp.Files("file"), loadError)
        If lr Is Nothing Then Return ErrorResponse(400, loadError)
        If lr.ValidUnique.Count = 0 Then
            Return ErrorResponse(400, "لا توجد أرقام صالحة في الملف")
        End If
        ' Pre-flight: probe every requested port; a modem that does not
        ' answer AT is excluded BEFORE the job starts, with its reason.
        Dim outcome As ProbeOutcome = ProbePorts(requested)
        If outcome.Usable.Count = 0 Then
            Dim sb0 As New StringBuilder()
            ' (Round 61) A-60-M1: on an enumeration timeout the honest error is the timeout, not "no healthy modem".
            Dim errText As String = If(outcome.TimedOut, "تعذر تعداد المنافذ: " & outcome.WmiError, "لا يوجد مودم سليم")
            sb0.Append("{""ok"":false,""error"":").Append(JsonStr(errText)).Append(",""excluded"":")
            sb0.Append(ExcludedJson(outcome.Excluded))
            sb0.Append("}")
            Return JsonResponse(400, sb0.ToString())
        End If
        Dim portNames As New List(Of String)()
        For Each u As ModemSender.PortInfo In outcome.Usable
            portNames.Add(u.PortName)
        Next
        Dim job As New SendJob(Guid.NewGuid().ToString("N"), message, delaySeconds, lr.ValidUnique, portNames, outcome.Excluded)
        BridgeMain.RegisterJob(job)
        job.Start()
        BridgeMain.Log("send job " & job.JobId & " started: " & lr.ValidUnique.Count & " numbers on " & String.Join(",", portNames) & " (msg: " & TruncateText(message, 30) & ")")
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,")
        sb.Append("""jobId"":").Append(JsonStr(job.JobId)).Append(",")
        sb.Append("""total"":").Append(lr.ValidUnique.Count).Append(",")
        sb.Append("""ports"":[")
        Dim first As Boolean = True
        For Each pn As String In portNames
            If Not first Then sb.Append(",")
            first = False
            sb.Append(JsonStr(pn))
        Next
        sb.Append("],""excluded"":").Append(ExcludedJson(outcome.Excluded))
        sb.Append("}")
        Return JsonResponse(200, sb.ToString())
        Finally
            ' (Round 62) B-61-H1/C-61-H1: release the slot claim on EVERY return from ApiSend, success or
            ' failure. The old "If Not registered" left the claim held forever after a successful send
            ' (registered=True), so every later send got a stuck 409 and the bridge was single-use per
            ' restart. The claim only needs to cover this method's own slow body (parse + probe + register);
            ' once registered, HasActiveJob guards the running job instead.
            ReleaseSendSlot()
        End Try
    End Function

    Public Function ApiProgress(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim jobId As String = req.QueryValue("job")
        Dim job As SendJob = BridgeMain.GetJob(jobId)
        If job Is Nothing Then Return ErrorResponse(404, "job not found")
        Return JsonResponse(200, job.SnapshotJson())
    End Function

    Public Function ApiStop(req As BridgeMain.HttpRequest) As BridgeMain.HttpResponse
        Dim jobId As String = JsonGetString(req.BodyText, "jobId")
        If String.IsNullOrEmpty(jobId) Then jobId = JsonGetString(req.BodyText, "job")
        Dim job As SendJob = BridgeMain.GetJob(jobId)
        If job Is Nothing Then Return ErrorResponse(404, "job not found")
        job.RequestStop()
        BridgeMain.Log("stop requested for job " & job.JobId)
        Return JsonResponse(200, "{""ok"":true,""stopped"":true}")
    End Function

    Private Function ExcludedJson(excluded As List(Of ExcludedPort)) As String
        Dim sb As New StringBuilder()
        sb.Append("[")
        Dim first As Boolean = True
        For Each e As ExcludedPort In excluded
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{""port"":").Append(JsonStr(e.Port)).Append(",""reason"":").Append(JsonStr(e.Reason)).Append("}")
        Next
        sb.Append("]")
        Return sb.ToString()
    End Function

    ' ------------------------------------------------------------------
    ' Probing helpers
    ' ------------------------------------------------------------------

    Public Class ExcludedPort
        Public Property Port As String = String.Empty
        Public Property Reason As String = String.Empty
    End Class

    Public Class ProbeOutcome
        Public Property Usable As New List(Of ModemSender.PortInfo)()
        Public Property Excluded As New List(Of ExcludedPort)()
        Public Property Results As New List(Of ModemSender.PortProbeResult)()
        Public Property WmiError As String = String.Empty
        ''' <summary>(Round 61) A-60-M1: True when the port enumeration itself timed out (hung WMI) - no port was
        ''' probed, so exclusions carry the timeout reason instead of "port not found".</summary>
        Public Property TimedOut As Boolean
    End Class

    ''' <summary>Same filter as ChooseProbedPort (Answered + not a data/GPS port, ranked), but returns ALL of them.</summary>
    Public Function UsableProbedPorts(results As List(Of ModemSender.PortProbeResult)) As List(Of ModemSender.PortInfo)
        Dim usable As New List(Of ModemSender.PortInfo)()
        If results Is Nothing Then Return usable
        For Each r As ModemSender.PortProbeResult In results
            If r Is Nothing OrElse r.Port Is Nothing Then Continue For
            If r.Kind <> ModemSender.PortProbeKind.Answered Then Continue For
            If r.Port.IsHuaweiDataPort Then Continue For
            If r.Port.IsHuaweiModemPort Then Continue For
            If r.Port.IsNmeaGps Then Continue For
            usable.Add(r.Port)
        Next
        usable.Sort(AddressOf ModemSender.ComparePortOrder)
        ' (Round 71W) Dedupe by port name: a ghost WMI row can share
        ' a COM number with the live row, so the same physical port
        ' would otherwise be sent to twice and ApiDetect would list
        ' it twice. Keep the first = highest-ranked twin.
        Dim deduped As New List(Of ModemSender.PortInfo)()
        For Each m As ModemSender.PortInfo In usable
            Dim seen As Boolean = False
            For Each x As ModemSender.PortInfo In deduped
                If String.Equals(x.PortName, m.PortName, StringComparison.OrdinalIgnoreCase) Then
                    seen = True
                    Exit For
                End If
            Next
            If Not seen Then deduped.Add(m)
        Next
        Return deduped
    End Function

    Private Function ProbeReason(pr As ModemSender.PortProbeResult) As String
        If pr Is Nothing OrElse pr.Port Is Nothing Then Return "لا يوجد رد"
        If pr.Port.IsHuaweiDataPort Then Return "منفذ بيانات/إنترنت، ليس للإرسال"
        If pr.Port.IsHuaweiModemPort Then Return "مودم (اتصال)، ليس للإرسال"
        If pr.Port.IsNmeaGps Then Return "منفذ GPS، ليس للإرسال"
        Select Case pr.Kind
            Case ModemSender.PortProbeKind.Answered
                Return String.Empty
            Case ModemSender.PortProbeKind.NoReply
                Return "لا يرد على AT"
            Case ModemSender.PortProbeKind.Busy
                Return "مشغول (برنامج آخر يستخدمه)"
            Case ModemSender.PortProbeKind.Error
                If String.IsNullOrEmpty(pr.Detail) Then Return "خطأ أثناء الفحص"
                Return "خطأ أثناء الفحص: " & TruncateText(pr.Detail, 100)
            Case Else
                Return "لا يوجد رد"
        End Select
    End Function

    ''' <summary>(Round 70W) Why this probed port is not usable, for the detect list (B-65W-L1).
    ''' Empty when the port IS usable (Answered + not data/modem/GPS) and when it simply
    ''' did not reply (the "لم يرد" tag already says that). Answered-but-excluded ports get
    ''' the exclusion reason; Busy/Error get theirs. ProbeReason itself is untouched
    ''' (it still feeds the send pre-flight Excluded list).</summary>
    Private Function DetectReason(pr As ModemSender.PortProbeResult) As String
        If pr Is Nothing OrElse pr.Port Is Nothing Then Return String.Empty
        If pr.Kind = ModemSender.PortProbeKind.Answered Then
            If pr.Port.IsHuaweiDataPort OrElse pr.Port.IsHuaweiModemPort OrElse pr.Port.IsNmeaGps Then
                Return ProbeReason(pr)
            End If
            Return String.Empty
        End If
        Select Case pr.Kind
            Case ModemSender.PortProbeKind.Busy, ModemSender.PortProbeKind.Error
                Return ProbeReason(pr)
            Case Else
                Return String.Empty
        End Select
    End Function
    ''' <summary>Probes the requested ports (Nothing/empty = all enumerated ports) and splits them into usable + excluded.</summary>
    Public Function ProbePorts(requested As List(Of String)) As ProbeOutcome
        Dim outcome As New ProbeOutcome()
        Dim wmiError As Exception = Nothing
        Dim enumError As Exception = Nothing
        Dim enumTimedOut As Boolean = False
        Dim ports As List(Of ModemSender.PortInfo) = Nothing
        Try
            ' (Round 61) A-60-M1: capped enumeration - a hung WMI sets TimedOut instead of hanging the call.
            ports = EnumeratePortsCapped(wmiError, enumError, enumTimedOut)
        Catch ex As Exception
            outcome.WmiError = ex.Message
            ports = New List(Of ModemSender.PortInfo)()
        End Try
        If enumTimedOut Then
            outcome.TimedOut = True
            outcome.WmiError = wmiError.Message
            ports = New List(Of ModemSender.PortInfo)()
        End If
        If enumError IsNot Nothing Then outcome.WmiError = enumError.Message
        If wmiError IsNot Nothing Then outcome.WmiError = wmiError.Message
        Dim wanted As New List(Of ModemSender.PortInfo)()
        If requested Is Nothing OrElse requested.Count = 0 Then
            wanted.AddRange(ports)
        Else
            For Each name As String In requested
                Dim found As ModemSender.PortInfo = Nothing
                For Each p As ModemSender.PortInfo In ports
                    If String.Equals(p.PortName, name, StringComparison.OrdinalIgnoreCase) Then found = p : Exit For
                Next
                If found Is Nothing Then
                    ' (Round 61) A-60-M1: on a timeout the port was never enumerated, so the honest reason is the
                    ' timeout itself, not "port not found".
                    Dim reason As String = If(outcome.TimedOut, outcome.WmiError, "المنفذ غير موجود")
                    outcome.Excluded.Add(New ExcludedPort() With {.Port = name, .Reason = reason})
                Else
                    wanted.Add(found)
                End If
            Next
        End If
        For Each info As ModemSender.PortInfo In wanted
            Dim pr As ModemSender.PortProbeResult = Nothing
            Try
                pr = ModemSender.ProbePortForAt(info)
            Catch ex As Exception
                pr = New ModemSender.PortProbeResult() With {.Port = info, .Kind = ModemSender.PortProbeKind.Error, .Detail = ex.Message}
            End Try
            outcome.Results.Add(pr)
        Next
        outcome.Usable = UsableProbedPorts(outcome.Results)
        Dim usableNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each u As ModemSender.PortInfo In outcome.Usable
            usableNames.Add(u.PortName)
        Next
        For Each pr As ModemSender.PortProbeResult In outcome.Results
            If pr Is Nothing OrElse pr.Port Is Nothing Then Continue For
            If Not usableNames.Contains(pr.Port.PortName) Then
                outcome.Excluded.Add(New ExcludedPort() With {.Port = pr.Port.PortName, .Reason = ProbeReason(pr)})
            End If
        Next
        Return outcome
    End Function

End Module
