Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading

''' <summary>
''' Multi-modem send job: numbers are distributed round-robin across the
''' chosen modems (number 1 -&gt; modem A, 2 -&gt; B, 3 -&gt; A, ...), one
''' thread per modem, each with its own ModemSender connection and its own
''' delay pacing. A modem that dies mid-job hands its remaining numbers
''' to ONE retry pass on the healthy modems; with no healthy modem left
''' those numbers fail loudly. The single-modem case is the same code
''' with one worker.
'''
''' The per-number send + late-+CMGS: credit logic mirrors the desktop
''' send loop (FrmMain BgwSend_DoWork): credit a pending late
''' confirmation to the previous row first (bounded), then to the
''' current row; the re-send warning clause and "(بدون تأكيد OK)" marker
''' follow the same standing rules.
''' </summary>

Public Class JobRow
    Public Property Number As String = String.Empty
    Public Property Port As String = String.Empty
    Public Property Ok As Boolean
    ''' <summary>(Round 61) C-60-M1/A-60-L1: True for a row that was never attempted because the job was stopped -
    ''' it counts in the cancelled bucket, not in failed, so sent+failed+cancelled = total.</summary>
    Public Property Cancelled As Boolean
    Public Property Message As String = String.Empty
    Public Property SegmentsSent As Integer
    Public Property SegmentsTotal As Integer
End Class

''' <summary>One modem's share of a send job: its own thread, its own ModemSender connection.</summary>
Public Class ModemWorker
    Public Property Port As String = String.Empty
    Public Property Queue As List(Of String)
    Public Property QueueIndex As Integer
    Public Property Status As String = "pending"
    Public Property Reason As String = String.Empty
    Public Property Current As String = String.Empty
    ''' <summary>True when this worker was created for the retry pass (its numbers are retried, never re-retried).</summary>
    Public Property Retried As Boolean
    ''' <summary>True once this worker's remaining numbers were moved into the job's retry list.</summary>
    Public Property Collected As Boolean
    Public Property Job As SendJob

    Private _thread As Thread

    Public Sub New(job As SendJob, port As String, queue As List(Of String), retried As Boolean)
        Me.Job = job
        Me.Port = port
        Me.Queue = queue
        Me.Retried = retried
    End Sub

    Public Sub Start()
        _thread = New Thread(AddressOf Run)
        _thread.IsBackground = True
        _thread.Start()
    End Sub

    Public Sub Join()
        If _thread IsNot Nothing Then _thread.Join()
    End Sub

    ''' <summary>Numbers this worker never attempted (QueueIndex only advances after a row is recorded).</summary>
    Public Function Remaining() As List(Of String)
        Dim rest As New List(Of String)()
        If Queue Is Nothing Then Return rest
        For idx As Integer = QueueIndex To Queue.Count - 1
            rest.Add(Queue(idx))
        Next
        Return rest
    End Function

    Private Function JobStop() As Boolean
        Return Job.StopRequested()
    End Function

    Private Sub MarkDead(reason As String)
        Status = "dead"
        ' (Round 61) C-60-M3: "Reason = reason" was a self-assignment (VB is case-insensitive, so the parameter
        ' shadows the property and it stayed "" forever). Me.Reason sets the property.
        Me.Reason = reason
        Job.AddLog(reason)
    End Sub

    Private Sub SleepDelay(totalMs As Integer)
        Dim waited As Integer = 0
        While waited < totalMs
            If JobStop() Then Return
            Thread.Sleep(100)
            waited += 100
        End While
    End Sub

    Public Sub Run()
        Status = "sending"
        Try
            Using modem As New ModemSender()
                Dim openOk As Boolean = False
                Try
                    openOk = modem.Open(Port, AddressOf JobStop)
                Catch ex As Exception
                    MarkDead("تعذر فتح المنفذ " & Port & ": " & ex.Message)
                    Return
                End Try
                If Not openOk Then
                    MarkDead("فشل فتح المنفذ " & Port & " (المودم لم يرد على أوامر AT/ATE0)")
                    Return
                End If
                Dim staleWithoutOk As Boolean = False
                modem.TakePendingLateCmgs(staleWithoutOk)
                Dim prevRow As JobRow = Nothing
                While QueueIndex < Queue.Count
                    If JobStop() Then Exit While
                    Dim number As String = Queue(QueueIndex)
                    Current = number
                    Dim r As ModemSender.SmsResult = Nothing
                    Dim threw As Exception = Nothing
                    Try
                        r = modem.SendSmsConnected(number, Job.Message, AddressOf JobStop)
                    Catch ex As Exception
                        threw = ex
                    End Try
                    Dim modemDead As Boolean = False
                    Dim deadReason As String = String.Empty
                    If threw IsNot Nothing Then
                        ' The send itself threw: the port may have died.
                        ' A plain AT tells a dead modem from a bad number.
                        Dim alive As Boolean = False
                        Try
                            alive = modem.TestConnection(AddressOf JobStop)
                        Catch
                            alive = False
                        End Try
                        If Not alive Then
                            modemDead = True
                            deadReason = "انقطع الاتصال بالمودم " & Port & " أثناء الإرسال: " & threw.Message
                        Else
                            r = New ModemSender.SmsResult() With {.Success = False, .Message = "خطأ أثناء الإرسال: " & threw.Message}
                        End If
                    End If
                    If r Is Nothing Then
                        r = New ModemSender.SmsResult() With {.Success = False, .Message = "لا يوجد رد من المودم"}
                    End If
                    ' Late-confirmation credit: previous row first, then this row (bounded).
                    Dim pw As Boolean = False
                    Dim pending As Integer = 0
                    Try
                        pending = modem.TakePendingLateCmgs(pw)
                    Catch
                        pending = 0
                    End Try
                    If pending > 0 AndAlso prevRow IsNot Nothing AndAlso Not prevRow.Ok AndAlso prevRow.SegmentsTotal > 0 AndAlso prevRow.SegmentsSent < prevRow.SegmentsTotal Then
                        Dim credit As Integer = Math.Min(pending, prevRow.SegmentsTotal - prevRow.SegmentsSent)
                        prevRow.SegmentsSent += credit
                        pending -= credit
                        If prevRow.SegmentsSent >= prevRow.SegmentsTotal Then
                            Job.CreditRow(prevRow, "تم (تأكيد متأخر)" & If(pw OrElse prevRow.Message.Contains("(بدون تأكيد OK)"), " (بدون تأكيد OK)", String.Empty))
                        End If
                    End If
                    If pending > 0 AndAlso Not r.Success AndAlso r.SegmentsTotal > 0 AndAlso r.SegmentsSent < r.SegmentsTotal Then
                        Dim credit As Integer = Math.Min(pending, r.SegmentsTotal - r.SegmentsSent)
                        r.SegmentsSent += credit
                        pending -= credit
                        If r.SegmentsSent >= r.SegmentsTotal Then
                            r.Success = True
                            r.Message = "تم (تأكيد متأخر)" & If(pw OrElse r.Message.Contains("(بدون تأكيد OK)"), " (بدون تأكيد OK)", String.Empty)
                        End If
                    End If
                    ' A failed send can also mean the modem died: check with AT.
                    If Not modemDead AndAlso Not r.Success Then
                        Dim alive As Boolean = False
                        Try
                            alive = modem.TestConnection(AddressOf JobStop)
                        Catch
                            alive = False
                        End Try
                        If Not alive Then
                            modemDead = True
                            deadReason = "انقطع الاتصال بالمودم " & Port & " أثناء الإرسال"
                        End If
                    End If
                    Dim row As New JobRow() With {.Number = number, .Port = Port, .Ok = r.Success, .Message = r.Message, .SegmentsSent = r.SegmentsSent, .SegmentsTotal = r.SegmentsTotal}
                    Job.RecordResult(row)
                    If r.Success Then
                        Job.AddLog("✓ " & number & " (" & Port & ")")
                    Else
                        Job.AddLog("✗ " & number & " (" & Port & "): " & r.Message)
                    End If
                    prevRow = row
                    QueueIndex += 1
                    If modemDead Then
                        MarkDead(deadReason)
                        Return
                    End If
                    If QueueIndex < Queue.Count Then SleepDelay(Job.DelaySeconds * 1000)
                End While
                Status = "done"
            End Using
        Catch ex As Exception
            MarkDead("خطأ في المودم " & Port & ": " & ex.Message)
        End Try
    End Sub
End Class

Public Class SendJob
    Public Property JobId As String = String.Empty
    Public Property Message As String = String.Empty
    Public Property DelaySeconds As Integer = 5
    Public Property Numbers As List(Of String)
    Public Property Workers As New List(Of ModemWorker)()
    Public Property Excluded As New List(Of ExcludedPort)()
    Public Property Done As Boolean
    Public Property Stopped As Boolean
    Public Property FatalError As String = String.Empty

    Private _results As New List(Of JobRow)()
    Private _sent As Integer
    Private _failed As Integer
    ''' <summary>(Round 61) C-60-M1/A-60-L1: rows never attempted because the job was stopped.</summary>
    Private _cancelled As Integer
    Private _logLines As New List(Of String)()
    Private _retryNumbers As New List(Of String)()
    Private _stopFlag As Integer
    Private _stateLock As New Object()
    Private _coordinator As Thread

    Public Sub New(jobId As String, message As String, delaySeconds As Integer, numbers As List(Of String), portNames As List(Of String), excluded As List(Of ExcludedPort))
        Me.JobId = jobId
        Me.Message = message
        Me.DelaySeconds = delaySeconds
        Me.Numbers = numbers
        Me.Excluded = excluded
        Dim queues As List(Of List(Of String)) = AssignRoundRobin(numbers, portNames.Count)
        For idx As Integer = 0 To portNames.Count - 1
            Workers.Add(New ModemWorker(Me, portNames(idx), queues(idx), False))
        Next
    End Sub

    Public Sub Start()
        _coordinator = New Thread(AddressOf RunCoordinator)
        _coordinator.IsBackground = True
        _coordinator.Start()
    End Sub

    Public Function StopRequested() As Boolean
        Return Threading.Volatile.Read(_stopFlag) = 1
    End Function

    Public Sub RequestStop()
        Threading.Interlocked.Exchange(_stopFlag, 1)
    End Sub

    Public Sub AddLog(line As String)
        SyncLock _stateLock
            _logLines.Add(DateTime.Now.ToString("HH:mm:ss") & " " & line)
            If _logLines.Count > 200 Then _logLines.RemoveRange(0, _logLines.Count - 200)
        End SyncLock
        BridgeMain.Log(line)
    End Sub

    Public Sub RecordResult(row As JobRow)
        SyncLock _stateLock
            _results.Add(row)
            If row.Ok Then _sent += 1 Else _failed += 1
        End SyncLock
    End Sub

    Public Sub CreditRow(row As JobRow, newMessage As String)
        SyncLock _stateLock
            If Not row.Ok Then
                row.Ok = True
                row.Message = newMessage
                _sent += 1
                _failed -= 1
            End If
        End SyncLock
    End Sub

    Public Function SentCount() As Integer
        SyncLock _stateLock
            Return _sent
        End SyncLock
    End Function

    Public Function FailedCount() As Integer
        SyncLock _stateLock
            Return _failed
        End SyncLock
    End Function

    ''' <summary>(Round 61) C-60-M1/A-60-L1: how many numbers were never attempted because the job was stopped.</summary>
    Public Function CancelledCount() As Integer
        SyncLock _stateLock
            Return _cancelled
        End SyncLock
    End Function

    Public Function RetryCount() As Integer
        SyncLock _stateLock
            Return _retryNumbers.Count
        End SyncLock
    End Function

    Public Sub RunCoordinator()
        Try
            AddLog("بدء الإرسال: " & Numbers.Count & " رقم على " & Workers.Count & " مودم")
            For Each w As ModemWorker In Workers
                w.Start()
            Next
            For Each w As ModemWorker In Workers
                w.Join()
            Next
            If StopRequested() Then
                ' (Round 61) C-60-M1/A-60-L1: a stop used to leave every unattempted number without a row
                ' (sent+failed ended below total, results missing rows). Every unattempted number gets a
                ' stopped row instead, so sent+failed+cancelled = total.
                RecordStoppedRows()
            Else
                CollectRetryFromDeadWorkers()
                If RetryCount() > 0 Then
                    Dim healthy As New List(Of ModemWorker)()
                    For Each w As ModemWorker In Workers
                        If w.Status = "done" Then healthy.Add(w)
                    Next
                    If healthy.Count > 0 Then
                        AddLog("إعادة محاولة " & RetryCount() & " رقم على مودم سليم")
                        Dim retryWorkers As List(Of ModemWorker) = CreateRetryWorkers(healthy)
                        For Each w As ModemWorker In retryWorkers
                            w.Start()
                        Next
                        For Each w As ModemWorker In retryWorkers
                            w.Join()
                        Next
                        SyncLock _stateLock
                            Workers.AddRange(retryWorkers)
                        End SyncLock
                        ' (Round 61) C-60-M1: a stop DURING the retry pass accounts the same way: the retry
                        ' workers' unattempted numbers become stopped rows, not "no healthy modem" failures.
                        If StopRequested() Then
                            RecordStoppedRows()
                        Else
                            CollectRetryFromDeadWorkers()
                        End If
                    End If
                End If
                If Not StopRequested() Then
                    ' Leftovers: no healthy modem (or the retry pass died too) - fail loudly, exactly once.
                    Dim leftovers As New List(Of String)()
                    SyncLock _stateLock
                        leftovers.AddRange(_retryNumbers)
                        _retryNumbers.Clear()
                    End SyncLock
                    For Each n As String In leftovers
                        RecordResult(New JobRow() With {.Number = n, .Port = String.Empty, .Ok = False, .Message = "تعذر الإرسال: لا يوجد مودم سليم متاح"})
                    Next
                End If
            End If
        Catch ex As Exception
            ' (Round 61) B-60-L1: the coordinator writes FatalError under the same lock SnapshotJson reads it under.
            SyncLock _stateLock
                FatalError = "خطأ في مهمة الإرسال: " & ex.Message
            End SyncLock
            AddLog(FatalError)
        Finally
            ' (Round 61) B-60-L1: Done/Stopped are written under the same lock SnapshotJson reads them under.
            SyncLock _stateLock
                If StopRequested() Then Stopped = True
                Done = True
            End SyncLock
            Dim doneLine As String = "اكتمل الإرسال: نجح " & SentCount() & "، فشل " & FailedCount() & " (من " & Numbers.Count & ")"
            If CancelledCount() > 0 Then doneLine &= "، أُوقف " & CancelledCount()
            AddLog(doneLine)
        End Try
    End Sub

    ''' <summary>(Round 61) C-60-M1/A-60-L1: records every number the job never attempted as a stopped row, so
    ''' sent+failed+cancelled = total and no number vanishes from the results. Runs only on the coordinator thread
    ''' after all workers joined, so QueueIndex is stable. A worker's remaining numbers are recorded once: workers
    ''' already collected into the retry list are skipped (their numbers live in the retry workers' queues or in
    ''' _retryNumbers, both drained here too).</summary>
    Public Sub RecordStoppedRows()
        SyncLock _stateLock
            For Each w As ModemWorker In Workers
                If Not w.Collected Then
                    w.Collected = True
                    For Each n As String In w.Remaining()
                        _results.Add(New JobRow() With {.Number = n, .Port = w.Port, .Ok = False, .Cancelled = True, .Message = "تم الإيقاف بواسطة المستخدم"})
                        _cancelled += 1
                    Next
                End If
            Next
            For Each n As String In _retryNumbers
                _results.Add(New JobRow() With {.Number = n, .Port = String.Empty, .Ok = False, .Cancelled = True, .Message = "تم الإيقاف بواسطة المستخدم"})
                _cancelled += 1
            Next
            _retryNumbers.Clear()
        End SyncLock
    End Sub

    ''' <summary>Moves the unattempted numbers of dead workers into the retry list (each worker collected once).</summary>
    Public Sub CollectRetryFromDeadWorkers()
        Dim workersCopy As List(Of ModemWorker)
        SyncLock _stateLock
            workersCopy = New List(Of ModemWorker)(Workers)
        End SyncLock
        For Each w As ModemWorker In workersCopy
            If w.Status = "dead" AndAlso Not w.Collected Then
                w.Collected = True
                SyncLock _stateLock
                    For Each n As String In w.Remaining()
                        _retryNumbers.Add(n)
                    Next
                End SyncLock
            End If
        Next
    End Sub

    ''' <summary>Splits the retry numbers round-robin across the healthy modems and builds retry workers.</summary>
    Friend Function CreateRetryWorkers(healthy As List(Of ModemWorker)) As List(Of ModemWorker)
        Dim result As New List(Of ModemWorker)()
        Dim retryCopy As New List(Of String)()
        SyncLock _stateLock
            retryCopy.AddRange(_retryNumbers)
            _retryNumbers.Clear()
        End SyncLock
        If retryCopy.Count = 0 OrElse healthy.Count = 0 Then Return result
        Dim ports As New List(Of String)()
        For Each w As ModemWorker In healthy
            ports.Add(w.Port)
        Next
        Dim queues As List(Of List(Of String)) = AssignRoundRobin(retryCopy, ports.Count)
        For idx As Integer = 0 To ports.Count - 1
            If queues(idx).Count > 0 Then
                result.Add(New ModemWorker(Me, ports(idx), queues(idx), True))
            End If
        Next
        Return result
    End Function

    Public Function SnapshotJson() As String
        Dim resultsCopy As List(Of JobRow)
        Dim logCopy As List(Of String)
        Dim workersCopy As List(Of ModemWorker)
        Dim sent As Integer, failed As Integer, cancelled As Integer
        Dim doneVal As Boolean, stoppedVal As Boolean
        Dim fatalVal As String = String.Empty
        SyncLock _stateLock
            resultsCopy = New List(Of JobRow)(_results)
            logCopy = New List(Of String)(_logLines)
            workersCopy = New List(Of ModemWorker)(Workers)
            sent = _sent
            failed = _failed
            cancelled = _cancelled
            ' (Round 61) B-60-L1: Done/Stopped/FatalError are read under the same lock the coordinator writes them
            ' under, so a snapshot never mixes values from two moments.
            doneVal = Done
            stoppedVal = Stopped
            fatalVal = FatalError
        End SyncLock
        Dim sb As New StringBuilder()
        sb.Append("{""ok"":true,")
        sb.Append("""jobId"":").Append(JsonStr(JobId)).Append(",")
        sb.Append("""total"":").Append(Numbers.Count).Append(",")
        sb.Append("""sent"":").Append(sent).Append(",")
        sb.Append("""failed"":").Append(failed).Append(",")
        sb.Append("""cancelled"":").Append(cancelled).Append(",")
        sb.Append("""done"":").Append(JsonBool(doneVal)).Append(",")
        sb.Append("""stopped"":").Append(JsonBool(stoppedVal)).Append(",")
        sb.Append("""fatalError"":").Append(JsonStr(fatalVal)).Append(",")
        sb.Append("""modems"":[")
        Dim first As Boolean = True
        ' (Round 61) A-60-M2/C-60-L1: aggregate workers by port into one entry each. A retry worker for the same
        ' port used to appear as a second modem entry with the same (double-counted) numbers. Status merges by
        ' rank dead > sending > done > pending; the dead worker's reason survives. The merged entries are COPIES -
        ' the live workers are never mutated here.
        Dim merged As New Dictionary(Of String, MergedModem)(StringComparer.OrdinalIgnoreCase)
        Dim mergedOrder As New List(Of String)()
        For Each w As ModemWorker In workersCopy
            Dim e As MergedModem = Nothing
            If Not merged.TryGetValue(w.Port, e) Then
                e = New MergedModem() With {.Port = w.Port, .Status = w.Status, .Current = w.Current, .Reason = w.Reason}
                merged(w.Port) = e
                mergedOrder.Add(w.Port)
            Else
                If ModemStatusRank(w.Status) > ModemStatusRank(e.Status) Then e.Status = w.Status
                If w.Status = "dead" Then e.Reason = w.Reason
                If w.Status = "sending" Then e.Current = w.Current
                If String.IsNullOrEmpty(e.Current) AndAlso Not String.IsNullOrEmpty(w.Current) Then e.Current = w.Current
                If e.Status = "dead" AndAlso String.IsNullOrEmpty(e.Reason) Then e.Reason = w.Reason
            End If
        Next
        For Each portKey As String In mergedOrder
            Dim w As MergedModem = merged(portKey)
            If Not first Then sb.Append(",")
            first = False
            Dim wSent As Integer = 0, wFailed As Integer = 0, wCancelled As Integer = 0
            For Each r As JobRow In resultsCopy
                If String.Equals(r.Port, w.Port, StringComparison.OrdinalIgnoreCase) Then
                    ' (Round 62) B-61-L1: a stopped row is cancelled, not failed - counting every non-Ok row as
                    ' failed made the modem line read "فشل 2" while the totals said "فشل 0، أُوقف 2".
                    ' (Block If: a single-line If cannot carry ElseIf - BC30205.)
                    If r.Ok Then
                        wSent += 1
                    ElseIf r.Cancelled Then
                        wCancelled += 1
                    Else
                        wFailed += 1
                    End If
                End If
            Next
            sb.Append("{")
            sb.Append("""port"":").Append(JsonStr(w.Port)).Append(",")
            sb.Append("""sent"":").Append(wSent).Append(",")
            sb.Append("""failed"":").Append(wFailed).Append(",")
            sb.Append("""cancelled"":").Append(wCancelled).Append(",")
            sb.Append("""status"":").Append(JsonStr(w.Status)).Append(",")
            sb.Append("""current"":").Append(JsonStr(w.Current)).Append(",")
            sb.Append("""reason"":").Append(JsonStr(w.Reason))
            sb.Append("}")
        Next
        For Each e As ExcludedPort In Excluded
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{""port"":").Append(JsonStr(e.Port)).Append(",""sent"":0,""failed"":0,""cancelled"":0,""status"":""excluded"",""current"":"""",""reason"":").Append(JsonStr(e.Reason)).Append("}")
        Next
        sb.Append("],""log"":[")
        first = True
        For Each line As String In logCopy
            If Not first Then sb.Append(",")
            first = False
            sb.Append(JsonStr(line))
        Next
        sb.Append("],""results"":[")
        first = True
        For Each r As JobRow In resultsCopy
            If Not first Then sb.Append(",")
            first = False
            sb.Append("{")
            sb.Append("""number"":").Append(JsonStr(r.Number)).Append(",")
            sb.Append("""port"":").Append(JsonStr(r.Port)).Append(",")
            sb.Append("""ok"":").Append(JsonBool(r.Ok)).Append(",")
            sb.Append("""cancelled"":").Append(JsonBool(r.Cancelled)).Append(",")
            sb.Append("""message"":").Append(JsonStr(r.Message)).Append(",")
            sb.Append("""segmentsSent"":").Append(r.SegmentsSent).Append(",")
            sb.Append("""segmentsTotal"":").Append(r.SegmentsTotal)
            sb.Append("}")
        Next
        sb.Append("]}")
        Return sb.ToString()
    End Function

    ''' <summary>(Round 61) A-60-M2/C-60-L1: merge rank for modem statuses - dead beats sending beats done beats pending.</summary>
    Private Shared Function ModemStatusRank(status As String) As Integer
        Select Case status
            Case "dead"
                Return 4
            Case "sending"
                Return 3
            Case "done"
                Return 2
            Case Else
                Return 1
        End Select
    End Function
End Class

''' <summary>(Round 61) A-60-M2/C-60-L1: one merged modem entry per port for SnapshotJson - workers sharing a port
''' (first pass + retry pass) fold into a single entry instead of double-counting.</summary>
Public Class MergedModem
    Public Property Port As String = String.Empty
    Public Property Status As String = "pending"
    Public Property Current As String = String.Empty
    Public Property Reason As String = String.Empty
End Class

''' <summary>Splits numbers round-robin: item i goes to worker (i Mod workerCount). Pure function - unit-tested in the harness.</summary>
Public Module BridgeRoundRobin
Public Function AssignRoundRobin(numbers As List(Of String), workerCount As Integer) As List(Of List(Of String))
    Dim result As New List(Of List(Of String))()
    If workerCount <= 0 OrElse numbers Is Nothing Then Return result
    For i As Integer = 0 To workerCount - 1
        result.Add(New List(Of String)())
    Next
    For i As Integer = 0 To numbers.Count - 1
        result(i Mod workerCount).Add(numbers(i))
    Next
    Return result
End Function
End Module
