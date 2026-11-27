Imports System.IO.Ports
Imports System.Management
Imports System.Text.RegularExpressions

''' <summary>
''' Sends SMS through a ZTE MF190 (or any AT-command USB modem).
''' The SMS port is the "ZTE Diagnostics Interface" (COM18 in jony's
''' screenshot - the COM number may change after re-plug/reboot, so trust
''' the "ZTE Diagnostics" label, not the number). COM17 (ZTE NMEA Device)
''' is GPS - NOT for SMS. COM3 (Intel SOL) is irrelevant.
''' (Round 48) Huawei E1550 is also supported: its SMS port is the
''' "PC UI Interface" (COM16 in jony's screenshot); the "Application
''' Interface" (COM15) is the internet/data port, not for SMS.
''' (Round 72) COM14 ("HUAWEI Mobile Connect - 3G Modem") is the
''' dial-up/RAS port: it lives under Device Manager's Modems node,
''' not Ports, carries no "(COM14)" in its name, and is excluded
''' via Win32_POTSModem.AttachedTo (see ApplyPotsModemRow), not
''' via its name - on real hardware its Name/Description are empty.
''' </summary>
Public Class ModemSender
    Implements IDisposable

    Public Class SmsResult
        Public Property Success As Boolean
        Public Property Message As String = String.Empty
        ''' <summary>True when the send was stopped by the user (Stop).</summary>
        Public Property Cancelled As Boolean
        ''' <summary>Standalone messages (segments) already sent before a failure/stop.</summary>
        Public Property SegmentsSent As Integer
        ''' <summary>Total standalone messages (segments) for this SMS.</summary>
        Public Property SegmentsTotal As Integer
        ''' <summary>(Round 57) C-55-L1: the concatenation reference
        ''' actually used for this send (1..255), or 0 when the
        ''' message was a single segment (no UDH, no ref). A retry
        ''' of this row passes it back via SendSmsConnected's
        ''' reuseConcatRef so the phone completes the SAME
        ''' concatenated message instead of starting a new one.</summary>
        Public Property ConcatRef As Integer
        ''' <summary>(Round 55) True when this send failed because the
        ''' modem's port itself is dead (a port I/O exception escaped
        ''' to the outer Catch, or the segment body write threw a
        ''' port exception): the modem cannot send anything more in
        ''' this run, so the caller stops this modem and requeues its
        ''' remaining numbers to a healthy modem. False for ordinary
        ''' failures (CMS ERROR, timeout, no credit) — those leave
        ''' the modem usable.</summary>
        Public Property PortDead As Boolean
        ''' <summary>(Round 75) LIVE 2026-11-23: True when Ctrl-Z
        ''' went out and the terminal reply was OK but no +CMGS:
        ''' arrived in the wait + grace + drain windows — the
        ''' segment MAY have been sent; its fate is UNKNOWN. Never
        ''' counted as sent (SegmentsSent stays 0 for it) and never
        ''' reported as a definite failure: the caller counts it
        ''' separately as unconfirmed, and a late +CMGS: credited
        ''' later clears it (see CreditPrevRowLateConfirmations).
        ''' False for ordinary failures (CMS ERROR, timeout with
        ''' no reply, prompt failure) — those are known-not-sent.</summary>
        Public Property Unconfirmed As Boolean
    End Class

    Public Class PortInfo
        Public Property PortName As String = String.Empty
        ''' <summary>WMI Name, e.g. "ZTE Diagnostics Interface (COM18)".</summary>
        Public Property Name As String = String.Empty
        Public Property Description As String = String.Empty
        Public ReadOnly Property DisplayText As String
            Get
                Dim label As String = If(String.IsNullOrEmpty(Description), Name, Description)
                Dim text As String
                If String.IsNullOrEmpty(label) Then
                    text = PortName
                Else
                    text = PortName & " - " & label
                End If
                If IsNmeaGps Then text &= " - GPS، ليس للإرسال"
                ' (Round 48) Huawei data port, same pattern as NMEA:
                ' flag it in the list so it is never picked by mistake.
                If IsHuaweiDataPort Then text &= " - بيانات/إنترنت، ليس للإرسال"
                ' (Round 64) Huawei dial-up modem port, same pattern:
                ' flag it in the list so it is never picked by mistake.
                If IsHuaweiModemPort Then text &= " - مودم (اتصال)، ليس للإرسال"
                Return text
            End Get
        End Property
        ''' <summary>(Round 64) F2: the CheckedListBox shows ToString()
        ''' (no DisplayMember is set), so without this override every
        ''' row showed "SmsSender.ModemSender+PortInfo" instead of the
        ''' port's name — the user could not tell COM14 from COM16.</summary>
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
        ''' <summary>True when this is the ZTE Diagnostics (AT/SMS) port.
        ''' Checks the WMI Name OR the Description - the Name is the one
        ''' that holds "ZTE Diagnostics Interface (COM18)".</summary>
        Public ReadOnly Property IsZteDiagnostics As Boolean
            Get
                Return Name.IndexOf("ZTE Diagnostics", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("ZTE Diagnostics", StringComparison.OrdinalIgnoreCase) >= 0
            End Get
        End Property
        ''' <summary>True for the GPS (NMEA) port - never use it for SMS.</summary>
        Public ReadOnly Property IsNmeaGps As Boolean
            Get
                Return Name.IndexOf("NMEA", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Name.IndexOf("GPS", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("NMEA", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("GPS", StringComparison.OrdinalIgnoreCase) >= 0
            End Get
        End Property
        ''' <summary>True for the Huawei E1550 SMS (AT) port: the "PC UI
        ''' Interface". (Round 48) Matches by substring, so both the
        ''' "HUAWEI Mobile Connect - 3G PC UI Interface" form and the
        ''' driver-INF "HUAWEI Mobile Connect - PC UI Interface" form
        ''' (no "3G") are recognised. The SMS/AT port on the E1550 is
        ''' the PC UI Interface, NOT the Application Interface
        ''' (Patriot GSM troubleshooting docs; CodeProject 3G dialer
        ''' article: open the "HUAWEI Mobile Connect - 3G PC UI"
        ''' port and send your commands).</summary>
        Public ReadOnly Property IsHuaweiSmsPort As Boolean
            Get
                Return Name.IndexOf("PC UI Interface", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("PC UI Interface", StringComparison.OrdinalIgnoreCase) >= 0
            End Get
        End Property
        ''' <summary>True for the Huawei internet/data port: the
        ''' "Application Interface" - never use it for SMS. (Round 48)
        ''' Matches by substring, so both the "3G Application
        ''' Interface" form and the bare "Application Interface"
        ''' form are recognised. (Round 49) Requires a HUAWEI vendor
        ''' marker in the Name or Description: without it, other
        ''' vendors' genuine AT/SMS ports (e.g. "Sierra Wireless
        ''' Application Interface") were wrongly flagged and
        ''' hard-blocked (B-47-L1/C-47-M1). Conscious trade-off:
        ''' the Huawei INF variant "FC - Application Interface"
        ''' (no HUAWEI in the name) is no longer flagged - it stays
        ''' a plain unflagged port, testable on its merits, exactly
        ''' like the v1.0 baseline.</summary>
        Public ReadOnly Property IsHuaweiDataPort As Boolean
            Get
                Dim hasApp As Boolean = Name.IndexOf("Application Interface", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("Application Interface", StringComparison.OrdinalIgnoreCase) >= 0
                If Not hasApp Then Return False
                ' (Round 49) Huawei vendor guard - see doc comment above.
                Return Name.IndexOf("HUAWEI", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("HUAWEI", StringComparison.OrdinalIgnoreCase) >= 0
            End Get
        End Property
        ''' <summary>(Round 64) True for the Huawei dial-up/RAS modem
        ''' port: "HUAWEI Mobile Connect - 3G Modem" (COM14 on jony's
        ''' E1550) — it lives under Device Manager's Modems node, not
        ''' Ports, and answers AT, so auto-detect counted it as a
        ''' second modem. Requires a HUAWEI marker and "Modem" in the
        ''' name/description, and is False for the PC UI Interface
        ''' (SMS) and the Application Interface (data) ports. The
''' "FC - Application Interface" variant has no HUAWEI marker,
''' so it stays unflagged (Round 49 trade-off, unchanged).
''' (Round 72) Also True when the port is a modem-class
''' (dial-up/RAS) device per Win32_POTSModem (IsModemClassPort),
''' even with an empty name: on real hardware the modem port
''' never carries a name (see ApplyPotsModemRow), so the
''' name-based check below could never fire for it.</summary>
        Public ReadOnly Property IsHuaweiModemPort As Boolean
            Get
                If IsHuaweiSmsPort OrElse IsHuaweiDataPort OrElse IsZteDiagnostics Then Return False
                ' (Round 72) Modem-class port (Win32_POTSModem):
                ' the dial-up port by Windows classification,
                ' whatever its name says. The SMS/data/ZTE guard
                ' above still wins, so a genuine SMS port is
                ' never excluded this way.
                ' (Round 73) The guard also covers ZTE
                ' Diagnostics: a ZTE SMS port marked
                ' modem-class must never be excluded
                ' either (B-72-L1 = C-72-L1).
                If IsModemClassPort Then Return True
                Dim hasModem As Boolean = Name.IndexOf("Modem", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("Modem", StringComparison.OrdinalIgnoreCase) >= 0
                If Not hasModem Then Return False
                Return Name.IndexOf("HUAWEI", StringComparison.OrdinalIgnoreCase) >= 0 _
                    OrElse Description.IndexOf("HUAWEI", StringComparison.OrdinalIgnoreCase) >= 0
            End Get
        End Property
        ''' <summary>(Round 72) True when Windows lists this COM
        ''' port as a modem-class (dial-up/RAS) device
        ''' (Win32_POTSModem.AttachedTo = this port). Set by
        ''' ApplyPotsModemRow during enumeration - it does NOT
        ''' depend on the port's name, which is empty for
        ''' modem-class ports on real hardware. A plain settable
        ''' property (not name-derived) so the sim can pin the
        ''' real flag path with fake modem rows.</summary>
        Public Property IsModemClassPort As Boolean = False
        ''' <summary>Numeric part of the port name (COM18 -> 18), for numeric sorting.</summary>
        Public ReadOnly Property PortNumber As Integer
            Get
                Dim m As Match = Regex.Match(PortName, "\d+")
                Dim n As Integer = 0
                If m.Success Then Integer.TryParse(m.Value, n)
                Return n
            End Get
        End Property
    End Class

    Private _port As SerialPort
    Private _disposed As Boolean
    ''' <summary>True when the last reply proved the modem is in command
    ''' mode (a final result line was seen), so no ESC is needed before
    ''' the next command. False after a "&gt;" prompt (text mode).</summary>
    Private _commandModeKnown As Boolean
    ' Bytes read ahead of a command (e.g. a very-late +CMGS: that
    ' arrived between commands) that must not be discarded: they are
    ' prepended to the next read's result instead. Written only by
    ' SendCommand's preserve read; consumed by the next SendCommand
    ' reply or by SendEscapeAndDrainCore's drain.
    Private _preservedText As String = String.Empty

    ''' <summary>Returns any preserved bytes not yet prepended to a
    ''' reply and clears the buffer.</summary>
    Private Function TakePreservedText() As String
        Dim text As String = _preservedText
        _preservedText = String.Empty
        Return text
    End Function
    ' D1: very-late +CMGS: confirmations that arrived BEFORE the
    ' current segment's Ctrl-Z (e.g. during the inter-number delay
    ' or a setup command) belong to a PREVIOUS send, whose SmsResult
    ' is already returned, so they must never be folded into the
    ' current segment's count. They are parked here instead: a
    ' +CMGS: is counted exactly once — either in its own send's
    ' WaitForResult / grace / drain path, or via this store, never
    ' both. TakePendingLateCmgs (below) consumes the store once;
    ' FrmMain calls it after each SendSmsConnected returns.
    Private _pendingLateCmgs As Integer
    ' FIX B-L2: unconfirmed parked confirmations are tracked as a
    ' COUNT, not a Boolean, so one stale OK can clear at most one
    ' of them (see StripStaleFinalResultLines). withoutOk for a
    ' Take is (count > 0).
    Private _pendingLateCmgsWithoutOkCount As Integer
    ' (Round 57) C-55-L1: concatenation reference counter, SHARED
    ' across all ModemSender instances in this process (one
    ' incrementing byte per multi-segment message, 1..255, wraps
    ' to 1, never 0 — 0 is rejected by the PDU builders). The
    ' Round-56 per-instance counter restarted at 1 for every new
    ' instance (one per modem per wave, and per run), so the first
    ' long message of every wave/run reused ref 1 for the same
    ' SIM — a new message's part 2 could join onto an incomplete
    ' part 1 still held by the phone. Two different modems may
    ' still share a value safely (concatenation is keyed by
    ' sender + reference). Guarded by _concatRefLock: the
    ' multi-modem waves draw refs from parallel Tasks.
    Private Shared _concatRef As Integer = 0
    Private Shared ReadOnly _concatRefLock As New Object()

    ''' <summary>Adds the +CMGS: lines found in a drain/probe pair to
    ''' the D1 pending late-confirmation store (see _pendingLateCmgs).
    ''' FIX B-L2: adds the EXACT number of counted lines that lack
    ''' their own trailing OK to _pendingLateCmgsWithoutOkCount
    ''' (DrainHasCmgs reports it), instead of setting one Boolean
    ''' for the whole store.</summary>
    Private Sub RecordPendingLateCmgs(drainedText As String, probeText As String)
        Dim drainedOk As Boolean = True
        Dim probeOk As Boolean = True
        Dim withoutOkCount As Integer = 0
        Dim found As Integer = DrainHasCmgs(drainedText, probeText, drainedOk, probeOk, withoutOkCount)
        If found <= 0 Then Return
        _pendingLateCmgs += found
        _pendingLateCmgsWithoutOkCount += withoutOkCount
        System.Diagnostics.Debug.WriteLine("D1 late +CMGS: recorded " & found.ToString() & " pending confirmation(s)")
    End Sub

    ''' <summary>D1/D2: parks count +CMGS: lines that were found in a
    ''' reply but could not be added to the current segment's count
    ''' (the bounded add was already full) in the pending
    ''' late-confirmation store, instead of dropping them: they
    ''' belong to an earlier segment, not to this one.
    ''' FIX B-L2: withoutOkCount is the EXACT number of parked lines
    ''' lacking their OK (from SplitCmgsFlags); the store's
    ''' without-OK count grows by that, not by a single Boolean.
    ''' FIX D-L1: the withoutOk Boolean parameter is removed - it
    ''' was never read by the body (only withoutOkCount was), and
    ''' SplitCmgsFlags guarantees it equalled
    ''' (withoutOkCount > 0), so it carried no information.</summary>
    Private Sub ParkPendingLateCmgs(count As Integer, withoutOkCount As Integer)
        If count <= 0 Then Return
        _pendingLateCmgs += count
        _pendingLateCmgsWithoutOkCount += withoutOkCount
        System.Diagnostics.Debug.WriteLine("D1/D2 late +CMGS: parked " & count.ToString() & " excess confirmation(s) as pending")
    End Sub

    ''' <summary>Consumes the D1 pending late-confirmation store
    ''' (count of very-late +CMGS: lines parked by
    ''' RecordPendingLateCmgs / ParkPendingLateCmgs) and clears it.
    ''' withoutOk is True when any parked line lacked its own
    ''' trailing OK.</summary>
    Public Function TakePendingLateCmgs(ByRef withoutOk As Boolean) As Integer
        Dim count As Integer = _pendingLateCmgs
        ' FIX B-L2: the ByRef Boolean boundary is kept (FrmMain's
        ' pendingWithoutOk / drainWithoutOk locals are unchanged);
        ' it is derived from the count inside ModemSender.
        withoutOk = (_pendingLateCmgsWithoutOkCount > 0)
        _pendingLateCmgs = 0
        _pendingLateCmgsWithoutOkCount = 0
        Return count
    End Function

    ''' <summary>A-M1: short post-loop grace drain. Reads any bytes
    ''' still sitting unread in the serial buffer (e.g. a +CMGS:
    ''' that arrived during the inter-number delay, when nobody was
    ''' reading) via the existing drain core with forceDrain, and
    ''' parks any +CMGS: lines found in the D1 pending store, so
    ''' FrmMain can credit them after the loop instead of Close()
    ''' destroying them. Short on purpose: a ~500ms probe, not a
    ''' full 3s one, so a normal run end is not delayed noticeably.
    ''' Returns True when the drain's probe proved command mode
    ''' (B-B6: the caller can tell a proven drain from an unproven
    ''' one). A +CMGS: parked by the probe's own StripStale call is
    ''' already removed from the probe text, so RecordPendingLateCmgs
    ''' below cannot count it twice.</summary>
    Public Function DrainAndParkLateCmgs(Optional stopPredicate As Func(Of Boolean) = Nothing) As Boolean
        Dim probeText As String = String.Empty
        ' FIX B-D1: end-of-run drain - parkRejoinedCmgs stays True
        ' (default): there is no current segment, so a rejoined
        ' +CMGS: is parked for the previous row, not counted here.
        Dim drainedText As String = SendEscapeAndDrainCore(stopPredicate, 500, probeText, forceDrain:=True)
        RecordPendingLateCmgs(drainedText, probeText)
        ' FIX B-L3: flush a +CMGS:-related fragment still stranded in
        ' _preservedText (see FlushStrandedCmgsFragment) — this is the
        ' last read of the run, so there is no next read to complete
        ' it and Close() would otherwise discard it uncounted.
        FlushStrandedCmgsFragment()
        Return _commandModeKnown
    End Function

    ''' <summary>FIX B-L3: a +CMGS:-related fragment still unterminated
    ''' at the end of the final grace drain (DrainAndParkLateCmgs) is
    ''' flushed into the pending store as ONE unconfirmed late
    ''' confirmation (withoutOk count +1) instead of being left
    ''' stranded for Close() to discard. The fragment is TAKEN out of
    ''' _preservedText (consumed, not copied), so it cannot be counted
    ''' twice. It cannot double-count a line RecordPendingLateCmgs
    ''' already parked on this call: that call only sees drainedText /
    ''' probeText, from which the fragment was already held back. A
    ''' non-+CMGS: fragment is restored untouched for the next read.
    ''' By construction _preservedText holds a single trailing
    ''' fragment, so any head before its last line break (normally
    ''' empty) is restored as-is.</summary>
    Private Sub FlushStrandedCmgsFragment()
        Dim leftover As String = TakePreservedText()
        If String.IsNullOrEmpty(leftover) Then Return
        Dim lastBreak As Integer = leftover.LastIndexOfAny(New Char() {ChrW(13), ChrW(10)})
        Dim head As String = If(lastBreak < 0, String.Empty, leftover.Substring(0, lastBreak + 1))
        Dim trailing As String = If(lastBreak < 0, leftover, leftover.Substring(lastBreak + 1))
        ' FIX B-D4: flush only an unambiguous "+CMGS:" fragment -
        ' a bare "+CM" / "+CMG" / "+C" prefix may be a "+CMS ERROR"
        ' (a FAILED segment) and must be left stranded/dropped,
        ' never counted as a confirmation.
        If Not IsUnambiguousCmgsFragmentText(trailing) Then
            _preservedText = leftover
            Return
        End If
        _preservedText = head
        _pendingLateCmgs += 1
        _pendingLateCmgsWithoutOkCount += 1
        System.Diagnostics.Debug.WriteLine("B-L3 late +CMGS: flushed stranded fragment as 1 pending confirmation (without OK)")
    End Sub

    ''' <summary>D5: strips stale final-result lines (OK / ERROR /
    ''' +CMS ERROR / +CME ERROR) from text read BEFORE a command was
    ''' written, so they can never satisfy that command's
    ''' expected/error match. Any +CMGS: line in the stale text is
    ''' NOT discarded: it is routed to the D1 pending
    ''' late-confirmation store (with its confirming OK, when one
    ''' follows it in the same text) and removed from the returned
    ''' text, so a +CMGS: is counted exactly once. Other lines
    ''' (echo, +CPIN: READY, ...) are kept as-is.</summary>
    ' FIX B-D1: parkCmgs = False (in-send recovery drain) keeps a
    ' rejoined +CMGS: line in the returned text instead of parking
    ' it, so the caller can count it for the current segment.
    ' FIX B-M1: keptCmgsText returns the +CMGS: lines (and their kept
    ' confirming OKs) that Strip kept instead of parking (parkCmgs =
    ' False only) — the caller re-appends them after the
    ' staleKeptLength cut, so they are not discarded by it.
    Private Function StripStaleFinalResultLines(preservedText As String, Optional parkCmgs As Boolean = True, Optional ByRef keptCmgsText As String = Nothing) As String
        keptCmgsText = String.Empty
        If String.IsNullOrEmpty(preservedText) Then Return preservedText
        ' B-B4: hold back a trailing line fragment that has no line
        ' terminator — it is not a complete line yet, so do not
        ' count, strip, or keep it as one: leave it in
        ' _preservedText for the next read, where the rest of the
        ' line may complete it. A trailing "+CMGS: 12" with its OK
        ' still in flight must not be counted (confirmed=0) and
        ' stripped, and a trailing partial "O" must not be kept for
        ' a live "K" to complete into a stale OK (nor a partial
        ' "ERR" into a stale ERROR). _preservedText is empty here
        ' (the caller passed TakePreservedText()), so re-preserving
        ' the fragment here is safe and it is never counted twice.
        Dim fragment As String = String.Empty
        Dim completeText As String = preservedText
        Dim lastBreak As Integer = preservedText.LastIndexOfAny(New Char() {ChrW(13), ChrW(10)})
        If lastBreak < 0 Then
            fragment = preservedText
            completeText = String.Empty
        ElseIf lastBreak < preservedText.Length - 1 Then
            fragment = preservedText.Substring(lastBreak + 1)
            completeText = preservedText.Substring(0, lastBreak + 1)
        End If
        _preservedText = fragment
        If String.IsNullOrEmpty(completeText) Then Return String.Empty
        ' FIX B-D2: do NOT park here, before the stale-line loop.
        ' An OK that PRECEDES a +CMGS: in this same text (e.g.
        ' preserved "OK" + vbCrLf + "+CMGS: 12") must never flip
        ' that +CMGS: to confirmed: Record would add it with
        ' withoutOkCount = 1, then the loop below would hit the
        ' earlier OK first and decrement the count to 0, while
        ' CmgsLineConfirmations on the same text says [False].
        ' The loop runs first, so a stale OK here can only
        ' decrement a without-OK count parked by an EARLIER call
        ' (whose +CMGS: precedes this whole text). An OK after a
        ' just-parked +CMGS: needs no decrement: either it is that
        ' line's own confirming OK (consumed by the j-scan below,
        ' and Record then marks the line confirmed anyway) or an
        ' ERROR / next +CMGS: broke the window first. Parking runs
        ' after the loop (see below).
        Dim lines As String() = SplitLines(completeText)
        Dim kept As New System.Text.StringBuilder()
        ' FIX B-M1: the kept (not parked) +CMGS: lines and their
        ' confirming OKs, returned to the caller via keptCmgsText.
        Dim keptCmgs As New System.Text.StringBuilder()
        ' FIX B-D1: True after a kept (not parked) +CMGS: line, until
        ' its confirming OK is kept or an ERROR / next +CMGS: ends it.
        Dim keepCmgsOk As Boolean = False
        Dim i As Integer = 0
        While i < lines.Length
            Dim t As String = lines(i).Trim()
            If t.Length = 0 Then
                kept.Append(lines(i))
                kept.Append(vbCrLf)
                i += 1
                Continue While
            End If
            If t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                If Not parkCmgs Then
                    ' FIX B-D1: in-send recovery drain - this +CMGS:
                    ' is NOT parked here; keep the line (and, via
                    ' keepCmgsOk below, its confirming OK) in the
                    ' returned text so the caller's CountGuardedCmgs
                    ' can count it for the current segment.
                    ' FIX B-M1: also record it in keptCmgsText, so
                    ' SendCommand can re-append it after the
                    ' staleKeptLength cut below.
                    kept.Append(lines(i))
                    kept.Append(vbCrLf)
                    keptCmgs.Append(lines(i))
                    keptCmgs.Append(vbCrLf)
                    keepCmgsOk = True
                    i += 1
                    Continue While
                End If
                ' D1: the line itself is parked after this loop
                ' (FIX B-D2) — drop it here.
                ' Its confirming OK (the first OK after it, before
                ' any ERROR / +CMS ERROR / +CME ERROR / +CMGS:
                ' line) belongs to it, not to the next command:
                ' drop that too.
                Dim j As Integer = i + 1
                While j < lines.Length
                    Dim u As String = lines(j).Trim()
                    If u.Length = 0 Then
                        j += 1
                        Continue While
                    End If
                    ' B-B3: skip any non-OK line (e.g. an unsolicited
                    ' +CMTI: / +CREG: URC between the +CMGS: and its
                    ' OK) instead of stopping at the first non-empty
                    ' line — mirroring CmgsLineConfirmations, so the OK is
                    ' really dropped and cannot satisfy the next
                    ' command.
                    If LineMatchesExpected(u, "OK") Then
                        i = j
                        Exit While
                    End If
                    If u.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                        OrElse u.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                        OrElse u.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) _
                        OrElse u.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                        Exit While
                    End If
                    j += 1
                End While
                i += 1
                Continue While
            End If
            If LineMatchesExpected(t, "OK") _
                OrElse t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                OrElse t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then
                ' FIX B-L2: an OK dropped here as stale may be the
                ' confirmation of a +CMGS: parked WITHOUT its OK by an
                ' earlier Strip call (its OK had not arrived yet) — the
                ' OK window for that confirmation has now closed with
                ' an OK in hand, so decrement the without-OK COUNT by
                ' one (floor 0) instead of clearing a Boolean for the
                ' whole store. Residual ambiguity, NOT assumed away:
                ' this OK may instead belong to an unrelated command
                ' (e.g. a late reply to a previous AT whose SendCommand
                ' already timed out) — nothing here proves it confirms
                ' a parked +CMGS:. Decrementing by one bounds that
                ' damage to a single parked confirmation per stale OK.
                ' FIX B-D1: when a +CMGS: line was kept above (not
                ' parked), its confirming OK is kept too, not dropped
                ' as stale and not decremented - without it the kept
                ' line would look unconfirmed to the caller.
                If Not parkCmgs AndAlso keepCmgsOk AndAlso LineMatchesExpected(t, "OK") Then
                    kept.Append(lines(i))
                    kept.Append(vbCrLf)
                    keptCmgs.Append(lines(i))
                    keptCmgs.Append(vbCrLf)
                    keepCmgsOk = False
                    i += 1
                    Continue While
                End If
                If Not parkCmgs Then keepCmgsOk = False
                If LineMatchesExpected(t, "OK") AndAlso _pendingLateCmgs > 0 AndAlso _pendingLateCmgsWithoutOkCount > 0 Then
                    _pendingLateCmgsWithoutOkCount -= 1
                End If
                i += 1
                Continue While
            End If
            kept.Append(lines(i))
            kept.Append(vbCrLf)
            i += 1
        End While
        ' FIX B-D2 (continued): park the +CMGS: lines only now,
        ' after the stale-OK loop above - an OK in this text can
        ' no longer decrement a count this very call just added.
        ' FIX B-D1: skipped when parkCmgs is False - the lines were
        ' kept in the returned text instead (see above).
        If parkCmgs Then RecordPendingLateCmgs(completeText, String.Empty)
        ' FIX B-M1: hand the kept +CMGS: portion to the caller (only
        ' non-empty when parkCmgs is False — see above).
        keptCmgsText = keptCmgs.ToString()
        Return kept.ToString()
    End Function
    ' Lazily built once, thread-safe by construction: no double-checked
    ' locking (which needs Volatile to be correct) on these maps.
    ' (VB has no Volatile field modifier, so Lazy(Of T) is used.)
    Private Shared ReadOnly _gsmDefaultMapLazy As New Lazy(Of Dictionary(Of Char, Integer))(AddressOf BuildGsmDefaultMap)
    Private Shared ReadOnly _gsmExtMapLazy As New Lazy(Of Dictionary(Of Char, Integer))(AddressOf BuildGsmExtensionMap)
    ' (R75W) Reverse maps for the inbox PDU decoder (septet -> character).
    Private Shared ReadOnly _gsmDefaultRevLazy As New Lazy(Of Dictionary(Of Integer, Char))(AddressOf BuildGsmDefaultRevMap)
    Private Shared ReadOnly _gsmExtRevLazy As New Lazy(Of Dictionary(Of Integer, Char))(AddressOf BuildGsmExtRevMap)

    Private Shared Function BuildGsmDefaultRevMap() As Dictionary(Of Integer, Char)
        Dim rev As New Dictionary(Of Integer, Char)()
        For Each kv As KeyValuePair(Of Char, Integer) In GsmDefaultMap()
            If Not rev.ContainsKey(kv.Value) Then rev.Add(kv.Value, kv.Key)
        Next
        Return rev
    End Function

    Private Shared Function BuildGsmExtRevMap() As Dictionary(Of Integer, Char)
        Dim rev As New Dictionary(Of Integer, Char)()
        For Each kv As KeyValuePair(Of Char, Integer) In GsmExtensionMap()
            If Not rev.ContainsKey(kv.Value) Then rev.Add(kv.Value, kv.Key)
        Next
        Return rev
    End Function

    Public ReadOnly Property IsOpen As Boolean
        Get
            Return _port IsNot Nothing AndAlso _port.IsOpen
        End Get
    End Property

    ''' <summary>Auto-preselect priority for the SMS port: the first ZTE
    ''' Diagnostics port; if none exists, the first Huawei "PC UI
    ''' Interface" port (E1550 SMS port). (Round 48) The Huawei
    ''' "Application Interface" is the internet/data port and is never
    ''' preselected. (Round 49) A port flagged as a data port is
    ''' skipped even if it also matches the SMS substring, so a
    ''' both-substrings description can never be preselected.
    ''' Detection is by description substring only - never
    ''' by COM number (COM15/COM16 are jony's machine only). With only
    ''' ZTE ports present this returns the same port as the old
    ''' ZTE-only loop in FrmMain.PopulatePorts.</summary>
    Public Shared Function SelectPreferredSmsPort(ports As List(Of PortInfo)) As PortInfo
        If ports Is Nothing Then Return Nothing
        For Each info As PortInfo In ports
            ' (Round 67) Skip a port that is both a ZTE Diagnostics
            ' port and a refused port (Huawei dial-up modem, Huawei
            ' data, or NMEA GPS): preselecting it would contradict
            ' the refusal backstops. (Round 66 guarded modem only.)
            If info.IsZteDiagnostics AndAlso Not info.IsHuaweiModemPort AndAlso Not info.IsHuaweiDataPort AndAlso Not info.IsNmeaGps Then Return info
        Next
        For Each info As PortInfo In ports
            ' (Round 49) Skip a port that is both an SMS port and a
            ' data port (e.g. "HUAWEI Mobile Connect - 3G PC UI
            ' Interface (Application Interface)"): preselecting it
            ' would contradict the data-port refusal backstops.
            ' (Round 67) Also skip an NMEA GPS port wearing the SMS
            ' substring; no modem guard is needed here (Sms implies
            ' Not Modem by the flag definition).
            If info.IsHuaweiSmsPort AndAlso Not info.IsHuaweiDataPort AndAlso Not info.IsNmeaGps Then Return info
        Next
        Return Nothing
    End Function

    ''' <summary>Port list order: ZTE Diagnostics first, then Huawei
    ''' "PC UI Interface", then the rest numerically by COM number.
    ''' (Round 48) A strict tier order: a port in an earlier tier
    ''' always sorts before a port in a later tier; inside a tier the
    ''' numeric-then-name order is unchanged.</summary>
    Public Shared Function ComparePortOrder(a As PortInfo, b As PortInfo) As Integer
        If a.IsZteDiagnostics AndAlso Not b.IsZteDiagnostics Then Return -1
        If b.IsZteDiagnostics AndAlso Not a.IsZteDiagnostics Then Return 1
        If a.IsHuaweiSmsPort AndAlso Not b.IsHuaweiSmsPort Then Return -1
        If b.IsHuaweiSmsPort AndAlso Not a.IsHuaweiSmsPort Then Return 1
        Dim cmp As Integer = a.PortNumber.CompareTo(b.PortNumber)
        If cmp <> 0 Then Return cmp
        Return String.Compare(a.PortName, b.PortName, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>One probe outcome for a single COM port.</summary>
    Public Enum PortProbeKind
        ''' <summary>The port answered a plain AT with an OK (or ERROR) final-result line.</summary>
        Answered
        ''' <summary>The port opened but nothing answered AT within the timeout.</summary>
        NoReply
        ''' <summary>The port could not be opened because another program holds it (e.g. Mobile Partner).</summary>
        Busy
        ''' <summary>Any other open / I/O failure while probing.</summary>
        [Error]
    End Enum

    ''' <summary>Result of probing one COM port for an AT-speaking modem.</summary>
    Public Class PortProbeResult
        ''' <summary>The enumerated port (full Name / Description), never a bare name.</summary>
        Public Property Port As PortInfo
        Public Property Kind As PortProbeKind
        ''' <summary>Best-effort ATI label (first reply line, capped at 60 chars); empty when unknown.</summary>
        Public Property Label As String = String.Empty
        ''' <summary>Open / I/O failure message for the log; empty otherwise.</summary>
        Public Property Detail As String = String.Empty
    End Class

    ''' <summary>(Round 52) Proves a reply comes from an AT-speaking
    ''' device: a standalone "OK" line, or an "ERROR" line (a device
    ''' that rejects AT still speaks AT; the connection test
    ''' afterwards is the real health check). Echo ("AT"),
    ''' unsolicited text and "OK"/"ERROR" inside a longer word or
    ''' line (e.g. "ATOK", "+CME ERROR: 10") do NOT count: lines are
    ''' split on CR/LF and matched exactly after trimming. Pure
    ''' logic (no SerialPort), so the sim pins it exactly.</summary>
    Public Shared Function ReplyProvesAtDevice(reply As String) As Boolean
        If String.IsNullOrEmpty(reply) Then Return False
        Dim lines As String() = reply.Split(New Char() {ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        For Each line As String In lines
            Dim t As String = line.Trim()
            If t = "OK" OrElse t = "ERROR" Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 52) Reads from an open probe port until the
    ''' deadline, accumulating whatever arrives. Polls with
    ''' ReadExisting + a short sleep (ReadExisting never blocks, so
    ''' the loop - not ReadTimeout - enforces the budget) and stops
    ''' early as soon as an OK/ERROR final-result line arrives.</summary>
    Private Shared Function ReadReplyUntil(sp As SerialPort, timeoutMs As Integer) As String
        Dim sb As New System.Text.StringBuilder()
        Dim deadline As System.DateTime = System.DateTime.UtcNow.AddMilliseconds(timeoutMs)
        While System.DateTime.UtcNow < deadline
            Dim chunk As String = String.Empty
            Try
                chunk = sp.ReadExisting()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Probe read failed: " & ex.Message)
                Exit While
            End Try
            If Not String.IsNullOrEmpty(chunk) Then sb.Append(chunk)
            If ReplyProvesAtDevice(sb.ToString()) Then Exit While
            System.Threading.Thread.Sleep(50)
        End While
        Return sb.ToString()
    End Function

    ''' <summary>(Round 52) Best-effort ATI label for a port that
    ''' already answered AT: sends "ATI", waits up to ~500 ms, takes
    ''' the first non-empty reply line that is not the echo ("ATI")
    ''' and not a final result ("OK"/"ERROR"), capped at 60 chars.
    ''' Any failure (or no useful line) returns String.Empty - the
    ''' label is display-only and must never break a successful
    ''' AT answer.</summary>
    Private Shared Function BestEffortAtiLabel(sp As SerialPort) As String
        Try
            sp.Write("ATI" & vbCr)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Probe ATI write failed: " & ex.Message)
            Return String.Empty
        End Try
        Dim reply As String = ReadReplyUntil(sp, 500)
        If String.IsNullOrEmpty(reply) Then Return String.Empty
        Dim lines As String() = reply.Split(New Char() {ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
        For Each line As String In lines
            Dim t As String = line.Trim()
            If t.Length = 0 Then Continue For
            If String.Equals(t, "ATI", System.StringComparison.OrdinalIgnoreCase) Then Continue For
            If t = "OK" OrElse t = "ERROR" Then Continue For
            If t.Length > 60 Then t = t.Substring(0, 60)
            Return t
        Next
        Return String.Empty
    End Function

    ''' <summary>(Round 52) True when the system currently lists
    ''' the port name (SerialPort.GetPortNames). Used only to tell a
    ''' held port (Busy) from a nonexistent one (Error) when Open
    ''' raises UnauthorizedAccessException - the runtime raises that
    ''' same exception for both on some platforms. Any failure here
    ''' answers True (assume held) so a transient enumeration fault
    ''' can never turn a held port into a misleading "does not
    ''' exist".</summary>
    Private Shared Function PortExistsInSystemList(portName As String) As Boolean
        Try
            For Each n As String In SerialPort.GetPortNames()
                If String.Equals(n, portName, System.StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Probe port-list check failed: " & ex.Message)
            Return True
        End Try
    End Function

    ''' <summary>(Round 53) Splits one probe's AT budget (timeoutMs)
    ''' across the two AT attempts of ProbePortForAt: the first AT
    ''' gets the first half, the ESC + drain gets a fixed 250 ms, and
    ''' the retry AT gets the rest (never below 100 ms, so a tiny
    ''' budget still gives the retry a chance). For the default 1500
    ''' ms budget that is 750 / 250 / 500, so the worst case per
    ''' silent port stays ~1500 ms and the README "حوالي ثانية ونصف
    ''' لكل منفذ صامت" still holds. Public Shared so the sim pins the
    ''' split exactly, like ReplyProvesAtDevice.</summary>
    Public Shared Sub ProbeRetryBudgets(timeoutMs As Integer, ByRef firstWaitMs As Integer, ByRef drainMs As Integer, ByRef secondWaitMs As Integer)
        If timeoutMs < 100 Then timeoutMs = 100
        firstWaitMs = timeoutMs \ 2
        drainMs = 250
        secondWaitMs = timeoutMs - firstWaitMs - drainMs
        If secondWaitMs < 100 Then secondWaitMs = 100
    End Sub

    ''' <summary>(Round 53) Probe-local ESC + drain: writes ESC
    ''' followed by CR (so it can never glue onto the next command —
    ''' same rule as Open()) to leave a stale "&gt;" prompt, then
    ''' drains whatever arrives for up to drainMs, exiting early
    ''' after ~100 ms of silence. All failures are swallowed: the
    ''' retry AT that follows is the real test. Mirrors the ESC step
    ''' of Open()'s AT retry, but on the probe's own short-lived
    ''' port — SendEscapeAndDrain works on the sender instance's
    ''' _port and cannot be reused here.</summary>
    Private Shared Sub ProbeEscapeAndDrain(sp As SerialPort, drainMs As Integer)
        Try
            sp.Write(New Byte() {CByte(&H1B), CByte(&HD)}, 0, 2)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Probe ESC write failed: " & ex.Message)
            Return
        End Try
        Dim watch As New System.Diagnostics.Stopwatch()
        watch.Start()
        Dim lastDataMs As Long = 0
        While watch.ElapsedMilliseconds < drainMs
            Dim chunk As String = String.Empty
            Try
                chunk = sp.ReadExisting()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Probe ESC drain failed: " & ex.Message)
                Exit While
            End Try
            If Not String.IsNullOrEmpty(chunk) Then
                lastDataMs = watch.ElapsedMilliseconds
            ElseIf watch.ElapsedMilliseconds - lastDataMs >= 100 Then
                Exit While
            End If
            System.Threading.Thread.Sleep(20)
        End While
        watch.Stop()
    End Sub

    ''' <summary>(Round 52) Probes one COM port for an AT-speaking
    ''' modem: opens it at 115200 8N1 with SHORT timeouts, sends a
    ''' plain "AT" (and, on no reply, ESC+CR + drain and ONE more AT — mirroring Open() for autobaud modems and stale ">" prompts), waits up to timeoutMs (default 1500 ms) in total for an
    ''' OK (or ERROR) final-result line. The port is closed and
    ''' disposed in a Finally on every path, and this function never
    ''' touches the sender instance's _port: it opens its own
    ''' short-lived SerialPort, so a probe can never disturb a
    ''' send/test and a send/test can never disturb a probe (the UI
    ''' also refuses a probe while either is running). Only AT (and
    ''' best-effort ATI for a display label) is ever sent - never
    ''' CMGF/CSCS/CMGS - so the probe is harmless to any serial
    ''' device. A full sweep stays fast: silent ports cost only the
    ''' 1500 ms AT wait (~9 s for 6 ports), and ATI only runs on
    ''' ports that already answered.</summary>
    Public Shared Function ProbePortForAt(info As PortInfo, Optional timeoutMs As Integer = 1500) As PortProbeResult
        Dim result As New PortProbeResult()
        If info Is Nothing Then
            result.Kind = PortProbeKind.[Error]
            result.Detail = "لا يوجد منفذ"
            Return result
        End If
        result.Port = info
        If timeoutMs < 100 Then timeoutMs = 100
        Dim sp As SerialPort = Nothing
        Try
            Try
                sp = New SerialPort(info.PortName, 115200, Parity.None, 8, StopBits.One)
                sp.DtrEnable = True
                sp.RtsEnable = True
                sp.ReadTimeout = 250
                sp.WriteTimeout = 1000
                sp.NewLine = vbCrLf
            Catch ex As Exception
                ' (Round 53) D3: the construction itself can throw
                ' (ArgumentException for an invalid port name) — it
                ' used to sit outside the open Try, so the throw
                ' escaped ProbePortForAt and aborted the whole sweep
                ' mid-list. Report this port as Error instead; the
                ' sweep continues with the next port.
                result.Kind = PortProbeKind.[Error]
                result.Detail = ex.Message
                Return result
            End Try
            Try
                sp.Open()
            Catch ex As UnauthorizedAccessException
                ' (Round 52) Usually the classic "port held by
                ' another program" failure (e.g. Mobile Partner
                ' holding the Huawei PC UI port) - Busy, not Error,
                ' so the caller can skip it with the "مشغول" line.
                ' But the runtime raises this same exception for a
                ' port that does not exist at all on some platforms
                ' (Linux: a nonexistent COM999 is "Access denied"),
                ' so tell them apart by asking the system whether
                ' the port exists right now: listed -> held (Busy);
                ' not listed -> does not exist (Error).
                result.Kind = If(PortExistsInSystemList(info.PortName), PortProbeKind.Busy, PortProbeKind.[Error])
                result.Detail = ex.Message
                Return result
            Catch ex As Exception
                result.Kind = PortProbeKind.[Error]
                result.Detail = ex.Message
                Return result
            End Try
            Try
                sp.DiscardInBuffer()
            Catch ex As Exception
                ' Stale bytes are harmless: the read loop below only
                ' looks for OK/ERROR lines, so they cannot fake an
                ' answer (an old OK line would only repeat a real
                ' earlier answer from this same port).
                System.Diagnostics.Debug.WriteLine("Probe discard failed: " & ex.Message)
            End Try
            ' (Round 53) Mirror Open(): a modem in autobaud mode
            ' answers the FIRST AT with nothing and the second with
            ' OK, and a modem stuck at a stale ">" prompt swallows
            ' "AT\r" as SMS text — the single bare AT missed both,
            ' while Open() accepts them. The budget is split (see
            ' ProbeRetryBudgets): first AT, then ESC+CR + drain,
            ' then ONE more AT, all inside the same ~timeoutMs.
            Dim firstWaitMs As Integer = 0
            Dim drainMs As Integer = 0
            Dim secondWaitMs As Integer = 0
            ProbeRetryBudgets(timeoutMs, firstWaitMs, drainMs, secondWaitMs)
            Try
                sp.Write("AT" & vbCr)
            Catch ex As Exception
                result.Kind = PortProbeKind.[Error]
                result.Detail = ex.Message
                Return result
            End Try
            Dim reply As String = ReadReplyUntil(sp, firstWaitMs)
            If Not ReplyProvesAtDevice(reply) Then
                ' No proof from the first AT: ESC (+CR, so it can
                ' never glue onto the next command) to leave a stale
                ' ">" prompt, drain whatever the ESC flushed, then
                ' one more AT in the remaining budget. The reply
                ' accumulator keeps both attempts: an OK split across
                ' the drain boundary still counts.
                ProbeEscapeAndDrain(sp, drainMs)
                Try
                    sp.Write("AT" & vbCr)
                Catch ex As Exception
                    result.Kind = PortProbeKind.[Error]
                    result.Detail = ex.Message
                    Return result
                End Try
                reply &= ReadReplyUntil(sp, secondWaitMs)
            End If
            If ReplyProvesAtDevice(reply) Then
                result.Kind = PortProbeKind.Answered
                result.Label = BestEffortAtiLabel(sp)
            Else
                result.Kind = PortProbeKind.NoReply
            End If
            Return result
        Finally
            If sp IsNot Nothing Then
                Try
                    If sp.IsOpen Then sp.Close()
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("Probe port close failed: " & ex.Message)
                End Try
                Try
                    sp.Dispose()
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("Probe port dispose failed: " & ex.Message)
                End Try
            End If
        End Try
    End Function

    ''' <summary>(Round 52) Picks the SMS port from probe results:
    ''' among ports that ANSWERED AT, apply the existing ranking
    ''' (ComparePortOrder: ZTE Diagnostics first, then Huawei PC UI,
    ''' then numerically) - but NEVER a Huawei data port or an
    ''' NMEA/GPS port, even if it answers AT: a data/NMEA port may
    ''' speak AT, and the Round 47/48 refusal backstops are absolute.
    ''' (Round 64) The Huawei dial-up modem port is refused the same way.
    ''' No answer, all busy, or an empty list returns Nothing. Pure
    ''' logic (no SerialPort), so the sim can pin it exactly.</summary>
    Public Shared Function ChooseProbedPort(results As List(Of PortProbeResult)) As PortInfo
        If results Is Nothing Then Return Nothing
        Dim candidates As New List(Of PortInfo)()
        For Each r As PortProbeResult In results
            If r Is Nothing OrElse r.Port Is Nothing Then Continue For
            If r.Kind <> PortProbeKind.Answered Then Continue For
            If r.Port.IsHuaweiDataPort Then Continue For
            If r.Port.IsHuaweiModemPort Then Continue For
            If r.Port.IsNmeaGps Then Continue For
            candidates.Add(r.Port)
        Next
        If candidates.Count = 0 Then Return Nothing
        candidates.Sort(AddressOf ComparePortOrder)
        Return candidates(0)
    End Function

    ''' <summary>(Round 55) Multi-modem version of ChooseProbedPort:
    ''' returns EVERY usable modem that answered AT, sorted by
    ''' ComparePortOrder (ZTE Diagnostics first, then Huawei PC UI,
    ''' then numerically) — never a Huawei data port, an NMEA/GPS
    ''' port or a Huawei dial-up modem port, even if it answers AT
    ''' (same absolute backstops).
    ''' No answer, all busy, or an empty list returns an empty list.
    ''' Pure logic (no SerialPort), so the sim can pin it exactly.</summary>
    Public Shared Function ChooseSendModems(results As List(Of PortProbeResult)) As List(Of PortInfo)
        Dim modems As New List(Of PortInfo)()
        If results Is Nothing Then Return modems
        For Each r As PortProbeResult In results
            If r Is Nothing OrElse r.Port Is Nothing Then Continue For
            If r.Kind <> PortProbeKind.Answered Then Continue For
            If r.Port.IsHuaweiDataPort Then Continue For
            If r.Port.IsHuaweiModemPort Then Continue For
            If r.Port.IsNmeaGps Then Continue For
            modems.Add(r.Port)
        Next
        modems.Sort(AddressOf ComparePortOrder)
        ' (Round 70) Dedupe by port name: a ghost WMI row can share
        ' a COM number with the live row, so the same physical port
        ' would otherwise be counted twice ("2 modems" for one
        ' flash drive). Keep the first = highest-ranked twin.
        Dim deduped As New List(Of PortInfo)()
        For Each m As PortInfo In modems
            Dim seen As Boolean = False
            For Each d As PortInfo In deduped
                If String.Equals(d.PortName, m.PortName, StringComparison.OrdinalIgnoreCase) Then
                    seen = True
                    Exit For
                End If
            Next
            If Not seen Then deduped.Add(m)
        Next
        Return deduped
    End Function

    ''' <summary>(Round 55) Round-robin distribution of row
    ''' indexes over modems: row i goes to modem (i Mod
    ''' modemCount), so the rows interleave evenly and the order
    ''' inside each modem stays the file order. modemCount = 1
    ''' returns every row on modem 0 — identical to v1.3.1.
    ''' modemCount &lt;= 0 returns an empty list (the caller must
    ''' pass at least one modem). Pure logic, pinned by the sim.</summary>
    Public Shared Function DistributeRows(totalRows As Integer, modemCount As Integer) As List(Of List(Of Integer))
        Dim share As New List(Of List(Of Integer))()
        If modemCount <= 0 Then Return share
        For m As Integer = 0 To modemCount - 1
            share.Add(New List(Of Integer)())
        Next
        For i As Integer = 0 To totalRows - 1
            share(i Mod modemCount).Add(i)
        Next
        Return share
    End Function

    ''' <summary>(Round 55) Failover requeue: a dead modem's
    ''' remaining (never attempted) rows are dealt round-robin
    ''' over the healthy modems, keeping their relative order.
    ''' healthyCount = 0 returns an empty list (the caller marks
    ''' the rows failed). Pure logic, pinned by the sim.</summary>
    Public Shared Function RequeueFailover(deadRows As List(Of Integer), healthyCount As Integer) As List(Of List(Of Integer))
        Dim share As New List(Of List(Of Integer))()
        If healthyCount <= 0 Then Return share
        For m As Integer = 0 To healthyCount - 1
            share.Add(New List(Of Integer)())
        Next
        If deadRows Is Nothing Then Return share
        For j As Integer = 0 To deadRows.Count - 1
            share(j Mod healthyCount).Add(deadRows(j))
        Next
        Return share
    End Function

    ''' <summary>
    ''' Lists COM ports with their device descriptions (via WMI when
    ''' possible) and preselects the ZTE Diagnostics Interface.
    ''' SerialPort.GetPortNames() is ALWAYS merged in, so a partial WMI
    ''' result can never hide a port. ZTE Diagnostics first, then the
    ''' rest sorted numerically by COM number (COM3 before COM17
    ''' before COM18).
    ''' FIX A-D2 (Round 14): the WMI failure is NOT swallowed here any
    ''' more — use GetAvailablePortsWithWmiError to receive it (the
    ''' parameterless GetAvailablePorts below keeps the old
    ''' swallow-and-return behaviour for compatibility).
    ''' </summary>
    Public Shared Function GetAvailablePorts() As List(Of PortInfo)
        Dim ignoredError As Exception = Nothing
        Return GetAvailablePortsWithWmiError(ignoredError)
    End Function

    ''' <summary>
    ''' FIX A-D2 (Round 14): same list as GetAvailablePorts, but a WMI
    ''' failure is also reported in wmiError (Nothing when WMI worked):
    ''' a fault is distinguishable from "no ports", so the caller can
    ''' report the fault instead of misreporting it as an unplugged
    ''' dongle. The GetPortNames() merge still runs after a WMI fault,
    ''' so a partial WMI result never hides a port.
    ''' </summary>
    Public Shared Function GetAvailablePortsWithWmiError(ByRef wmiError As Exception) As List(Of PortInfo)
        wmiError = Nothing
        Dim ports As New List(Of PortInfo)()
        Try
            Using searcher As New ManagementObjectSearcher(
                    "SELECT Name, Description FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'")
                Using results As ManagementObjectCollection = searcher.Get()
                    For Each obj As ManagementObject In results
                        Try
                            Dim name As String = String.Empty
                            Dim description As String = String.Empty
                            If obj("Name") IsNot Nothing Then name = obj("Name").ToString()
                            If obj("Description") IsNot Nothing Then description = obj("Description").ToString()
                            Dim m As Match = Regex.Match(name, "\(COM\d+\)")
                            If m.Success Then
                                Dim info As New PortInfo()
                                info.PortName = m.Value.Trim("("c, ")"c)
                                info.Name = name
                                info.Description = description
                                ports.Add(info)
                            End If
                        Finally
                            obj.Dispose()
                        End Try
                    Next
                End Using
            End Using
        Catch ex As Exception
            ' FIX A-D2: surface the WMI fault to the caller (via
            ' wmiError) instead of swallowing it here — the old
            ' swallow fell through to the merge and returned normally,
            ' so the caller's WMI-fault branch could never fire.
            ' The GetPortNames() merge below still runs.
            wmiError = ex
            System.Diagnostics.Debug.WriteLine("WMI port lookup failed: " & ex.Message)
        End Try

        ' Always merge GetPortNames() - a partial WMI success must never
        ' suppress the fallback list. Dedupe by port name.
        For Each portName As String In SerialPort.GetPortNames()
            Dim exists As Boolean = False
            For Each info As PortInfo In ports
                If String.Equals(info.PortName, portName, StringComparison.OrdinalIgnoreCase) Then
                    exists = True
                    Exit For
                End If
            Next
            If Not exists Then
                Dim info As New PortInfo()
                info.PortName = portName
                info.Description = String.Empty
                ports.Add(info)
            End If
        Next

        ' (Round 72) Modem-class devices ("HUAWEI Mobile Connect -
        ' 3G Modem", COM14 on jony's E1550) live under Device
        ' Manager's Modems node, NOT Ports: their PnP name carries
        ' no "(COMn)", so the Win32_PnPEntity query above never
        ' sees them and they only arrive via the GetPortNames()
        ' merge with an empty Name/Description - the name-based
        ' IsHuaweiModemPort flag could never fire on real hardware
        ' (it only worked in the sim, where test data supplied the
        ' name). Win32_POTSModem.AttachedTo gives the COM port for
        ' every modem-class device: mark the matching row (and
        ' fill its Name/Description when empty, so the list shows
        ' the real name). Runs inside the same WMI call, so the
        ' EnumeratePortsCapped thread + timeout still bound it.
        ' A failure here never hides a port: the rows stay listed
        ' (and the name-based flag still applies when a name is
        ' present); wmiError is set only if the PnP query did not
        ' already report a fault.
        Try
            Using modemSearcher As New ManagementObjectSearcher(
                    "SELECT AttachedTo, Name, Description FROM Win32_POTSModem")
                Using modemResults As ManagementObjectCollection = modemSearcher.Get()
                    For Each modemObj As ManagementObject In modemResults
                        Try
                            Dim attachedTo As String = String.Empty
                            Dim modemName As String = String.Empty
                            Dim modemDesc As String = String.Empty
                            If modemObj("AttachedTo") IsNot Nothing Then attachedTo = modemObj("AttachedTo").ToString()
                            If modemObj("Name") IsNot Nothing Then modemName = modemObj("Name").ToString()
                            If modemObj("Description") IsNot Nothing Then modemDesc = modemObj("Description").ToString()
                            ApplyPotsModemRow(ports, attachedTo, modemName, modemDesc)
                        Finally
                            modemObj.Dispose()
                        End Try
                    Next
                End Using
            End Using
        Catch ex As Exception
            If wmiError Is Nothing Then wmiError = ex
            System.Diagnostics.Debug.WriteLine("WMI modem lookup failed: " & ex.Message)
        End Try

        ' (Round 48) ZTE Diagnostics Interface first, then the Huawei
        ' "PC UI Interface" (E1550 SMS port), then the rest numerically.
        ' ZTE-only behaviour is unchanged: with no Huawei port present
        ' the new tier never fires.
        ports.Sort(AddressOf ComparePortOrder)
        Return ports
    End Function

    ''' <summary>(Round 72) Applies one Win32_POTSModem row to the
    ''' port list: marks the matching COM port as a modem-class
    ''' (dial-up/RAS) port and fills its Name/Description when
    ''' empty, so the list shows the real device name and every
    ''' IsHuaweiModemPort gate fires for it. Shared so the sim
    ''' can exercise the REAL code with fake modem rows (WMI
    ''' itself cannot run on Linux). Matching is by PortName
    ''' (OrdinalIgnoreCase); a row whose AttachedTo matches no
    ''' listed port is ignored, and a row with an empty
    ''' AttachedTo never marks anything.</summary>
    Friend Shared Sub ApplyPotsModemRow(ports As List(Of PortInfo), attachedTo As String, modemName As String, modemDesc As String)
        If ports Is Nothing OrElse String.IsNullOrEmpty(attachedTo) Then Return
        For Each info As PortInfo In ports
            If String.Equals(info.PortName, attachedTo, StringComparison.OrdinalIgnoreCase) Then
                info.IsModemClassPort = True
                If String.IsNullOrEmpty(info.Name) AndAlso Not String.IsNullOrEmpty(modemName) Then info.Name = modemName
                If String.IsNullOrEmpty(info.Description) AndAlso Not String.IsNullOrEmpty(modemDesc) Then info.Description = modemDesc
                ' (Round 73) No Exit For: mark every matching
                ' row, so a twin row is never left unmarked (B-O1/C-O1).
            End If
        Next
    End Sub

    ''' <summary>Opens the serial port: 115200 baud, 8 data bits, no parity, 1 stop bit.
    ''' Sends a plain AT first (no speculative bare ESC: a strict modem
    ''' can buffer it into the next command line); ESC (followed by CR,
    ''' so it can never glue onto the next command) is sent only when a
    ''' stale "&gt;" prompt is actually waiting or AT gets no reply.
    ''' Then ATE0 so command echo is off (echoed message text can never
    ''' be mistaken for a modem reply), retried once without ESC.
    ''' Split contract: a failure to open the port itself still THROWS
    ''' (the half-built port is disposed first); a failure in the
    ''' command phase (AT / ATE0 gets no OK, or an I/O error mid-open)
    ''' returns False with the port closed instead of throwing.</summary>
    Public Function Open(portName As String, Optional stopPredicate As Func(Of Boolean) = Nothing) As Boolean
        Close()
        _port = New SerialPort(portName, 115200, Parity.None, 8, StopBits.One)
        _port.DtrEnable = True
        _port.RtsEnable = True
        _port.ReadTimeout = 5000
        _port.WriteTimeout = 5000
        _port.NewLine = vbCrLf
        Try
            _port.Open()
        Catch ex As Exception
            Dim broken As SerialPort = _port
            _port = Nothing
            Try
                broken.Dispose()
            Catch ex2 As Exception
                System.Diagnostics.Debug.WriteLine("Port dispose failed: " & ex2.Message)
            End Try
            Throw
        End Try
        _commandModeKnown = False
        ' Peek for a stale ">" prompt BEFORE discarding the buffers:
        ' the discard would erase the only evidence of it. A bare ESC
        ' is never sent speculatively (see the summary above).
        Try
            Dim stalePrompt As Boolean = HasPromptWaiting()
        ' Clear any stale data the modem may have sent on open.
        Try
            _port.DiscardInBuffer()
            _port.DiscardOutBuffer()
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Discard buffers failed: " & ex.Message)
        End Try
        If stalePrompt Then
            SendEscapeAndDrain(stopPredicate)
            If Not _commandModeKnown Then
                ' The ESC probe did not prove command mode: the modem
                ' may still be stuck at a ">" prompt. One more ESC+CR
                ' before ATE0; if it is still stuck, ATE0 fails and
                ' Open returns False below.
                Try
                    _port.Write(New Byte() {CByte(&H1B), CByte(&HD)}, 0, 2)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("Stale-prompt ESC retry failed: " & ex.Message)
                End Try
                ' Drain after the retry ESC before sending ATE0: a
                ' still-stuck modem would swallow ATE0 as SMS text.
                ' If the ">" prompt persists, fail promptly instead.
                Dim retryText As New System.Text.StringBuilder()
                Dim retryWatch As New System.Diagnostics.Stopwatch()
                retryWatch.Start()
                Dim retryLastDataMs As Long = 0
                While retryWatch.ElapsedMilliseconds < 500
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                    Dim retryChunk As String = String.Empty
                    Try
                        retryChunk = _port.ReadExisting()
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("Stale-prompt retry drain failed: " & ex.Message)
                        Exit While
                    End Try
                    If Not String.IsNullOrEmpty(retryChunk) Then
                        retryText.Append(retryChunk)
                        retryLastDataMs = retryWatch.ElapsedMilliseconds
                    ElseIf retryWatch.ElapsedMilliseconds - retryLastDataMs >= 100 Then
                        Exit While
                    End If
                    Threading.Thread.Sleep(20)
                End While
                retryWatch.Stop()
                If LineMatchesExpected(retryText.ToString(), "OK") Then _commandModeKnown = True
                If HasPromptLine(retryText.ToString()) Then
                    System.Diagnostics.Debug.WriteLine("Stale prompt persists after ESC retry")
                    Close()
                    Return False
                End If
            End If
        Else
            ' Plain AT first: proves command mode with no ESC at all.
            Dim atReply As String = SendCommand("AT", "OK", 3000, stopPredicate)
            If Not ReplyMatchesExpected(atReply, "OK") Then
                ' No reply: the modem may be stuck at a ">" prompt
                ' whose evidence was already consumed. ESC (+CR, so
                ' it can never glue onto the next command) + drain,
                ' then one more AT attempt.
                SendEscapeAndDrain(stopPredicate)
                atReply = SendCommand("AT", "OK", 3000, stopPredicate)
            End If
        End If
        Dim ate0Reply As String = SendCommand("ATE0", "OK", 3000, stopPredicate)
        If Not ReplyMatchesExpected(ate0Reply, "OK") Then
            ' One retry without ESC before giving up.
            ate0Reply = SendCommand("ATE0", "OK", 3000, stopPredicate)
        End If
        If Not ReplyMatchesExpected(ate0Reply, "OK") Then
            System.Diagnostics.Debug.WriteLine("ATE0 failed: " & ate0Reply)
            Close()
            Return False
        End If
        Catch ex As Exception
            ' USB unplug (or any I/O failure) mid-open must honour the
            ' Boolean contract: dispose the port and return False
            ' instead of throwing out of Open().
            System.Diagnostics.Debug.WriteLine("Open command phase failed: " & ex.Message)
            Close()
            Return False
        End Try
        Return _port IsNot Nothing AndAlso _port.IsOpen
    End Function

    Public Sub Close()
        If _port IsNot Nothing Then
            Try
                If _port.IsOpen Then _port.Close()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Port close failed: " & ex.Message)
            Finally
                _port.Dispose()
                _port = Nothing
            End Try
        End If
        _commandModeKnown = False
        ' Preserved bytes belong to the closed session: never let a
        ' stale +CMGS: (or OK) leak into the next session's replies.
        _preservedText = String.Empty
        ' D1: pending late confirmations belong to the closed
        ' session too: never let one leak into the next session
        ' and get credited to a row it has nothing to do with.
        _pendingLateCmgs = 0
        _pendingLateCmgsWithoutOkCount = 0
    End Sub

    ''' <summary>Briefly reads (consumes) bytes already waiting in the
    ''' input buffer looking for a stale "&gt;" prompt, without waiting
    ''' for a live reply. Anything read here is stale input at open
    ''' time. NOTE: this is not a peek - ReadExisting() consumes the
    ''' bytes it returns.</summary>
    Private Function HasPromptWaiting() As Boolean
        If _port Is Nothing OrElse Not _port.IsOpen Then Return False
        Try
            Dim watch As New System.Diagnostics.Stopwatch()
            watch.Start()
            Dim pending As New System.Text.StringBuilder()
            While watch.ElapsedMilliseconds < 300
                Dim chunk As String = _port.ReadExisting()
                If Not String.IsNullOrEmpty(chunk) Then
                    pending.Append(chunk)
                    If HasPromptLine(pending.ToString()) Then Return True
                End If
                Threading.Thread.Sleep(20)
            End While
            watch.Stop()
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Prompt peek failed: " & ex.Message)
        End Try
        Return False
    End Function

    Private Shared Function SplitLines(text As String) As String()
        If String.IsNullOrEmpty(text) Then Return New String() {}
        Return text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
    End Function

    ''' <summary>True when any line of a reply starts with the given prefix
    ''' (ignoring leading whitespace). Line-anchored checks only: message
    ''' text echoed back must never count as an ERROR / +CMGS: result.</summary>
    Private Shared Function LineStartsWith(text As String, prefix As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each line As String In SplitLines(text)
            If line.TrimStart().StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 57) C-55-L2: the final-result lines
    ''' (ERROR / +CMS ERROR / +CME ERROR, plus — Round 77 — the
    ''' other non-OK final codes NO CARRIER / NO DIALTONE / BUSY /
    ''' NO ANSWER) found in a drained text,
    ''' joined with vbCrLf — NOT spaces — so BuildTimeoutReply can
    ''' split and dedupe them per line (Round 59 B-57-L1: a space
    ''' join fused two distinct lines into one pseudo-line and the
    ''' cross-source dedupe could no longer match either).
    ''' Line-anchored like LineStartsWith: message text echoed back
    ''' must never count. The probe's OK is never a final-result
    ''' line, so it is never included. includeBareError=False
    ''' (Round 59 C-57-L1) keeps only +CMS ERROR / +CME ERROR
    ''' lines: plain AT is never answered with a CMGS-family
    ''' error, so those can only be the submit's late answer,
    ''' while a bare ERROR answers the probe itself.
    ''' (Round 77) R76-A-L2/D-O1: a BUSY / NO CARRIER arriving only
    ''' in the drain was invisible — the predicate and the suffix
    ''' never saw it, so a definite refusal counted as unconfirmed.
    ''' These codes are mined exactly like the error lines (plain
    ''' AT is never answered with them, so they can only be the
    ''' submit's late answer, probe window included). OK is never
    ''' mined here — the caller mines a drain OK separately, from
    ''' the late drain only, so the probe's OK can never leak in —
    ''' and +CMGS: never is either (CountGuardedCmgs owns it).</summary>
    Private Shared Function DrainFinalResultLines(text As String, Optional includeBareError As Boolean = True) As String
        If String.IsNullOrEmpty(text) Then Return String.Empty
        Dim found As New List(Of String)()
        For Each line As String In SplitLines(text)
            Dim t As String = line.Trim()
            Dim isCmsOrCme As Boolean = t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase)
            Dim isBareError As Boolean = t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase)
            ' (Round 77) R76-A-L2/D-O1: a non-OK final-result line
            ' other than +CMGS: and the error family is a definite
            ' refusal too — mine it, so it reaches the predicate
            ' and the suffix. The exclusions reuse the file's own
            ' line classifiers, so no new unresolvable name is
            ' introduced: OK stays out (ReplyHasTerminalOk on the
            ' single line — the caller mines a drain OK separately,
            ' late drain only), +CMGS: stays out (CountGuardedCmgs
            ' owns it), and the error lines stay behind the
            ' includeBareError gate (a bare ERROR can answer the
            ' probe itself).
            Dim isNonOkFinal As Boolean = IsFinalResultLineText(t) _
                AndAlso Not ReplyHasTerminalOk(t) _
                AndAlso Not LineStartsWith(t, "+CMGS:") _
                AndAlso Not isCmsOrCme AndAlso Not isBareError
            If isCmsOrCme OrElse (includeBareError AndAlso isBareError) OrElse isNonOkFinal Then
                ' (Round 58) C-56-L2: the same final-result line can
                ' arrive several times in one drain window — keep it
                ' once (different lines are all kept).
                If Not found.Contains(t) Then found.Add(t)
            End If
        Next
        Return String.Join(vbCrLf, found)
    End Function

    ''' <summary>(Round 58) C-56-L2: flattens the timeout-reply
    ''' sources (initial reply, grace read, drain final lines, and
    ''' — Round 77 — the raw drain text) into one space-separated
    ''' string with identical lines shown once: the same +CMS ERROR
    ''' can arrive in both the grace window and the drain window,
    ''' and repeating it wastes the 200-char window. First
    ''' occurrence wins, order kept; empty lines dropped. The
    ''' 200-char cap stays at the call site. The raw drain text is
    ''' a source because its unsolicited lines (BOOT:/^RSSI:) are
    ''' mined by nothing else — without it the suffix wrongly said
    ''' "لا يوجد رد" when the drain held bytes. The probe text is
    ''' deliberately NOT a source: its OK answers the probe.</summary>
    Friend Shared Function BuildTimeoutReply(response As String, graceText As String, drainFinal As String, Optional drainRaw As String = "") As String
        Dim seen As New System.Collections.Generic.HashSet(Of String)(StringComparer.Ordinal)
        Dim parts As New System.Collections.Generic.List(Of String)()
        For Each source As String In New String() {response, graceText, drainFinal, drainRaw}
            If String.IsNullOrEmpty(source) Then Continue For
            For Each line As String In SplitLines(source)
                Dim t As String = line.Trim()
                If t.Length = 0 Then Continue For
                If seen.Add(t) Then parts.Add(t)
            Next
        Next
        Return String.Join(" ", parts)
    End Function

    ''' <summary>FIX B-M1: the per-line half of LineMatchesExpected,
    ''' factored out so the stale-aware matchers below can apply the
    ''' exact same rule to one line. Behaviour unchanged: exact line,
    ''' line starting with it on a word boundary ("OKAY" must not match
    ''' "OK"), or line ending with it after a space. The "&gt;" prompt
    ''' matches only a line that is exactly "&gt;" after trim.</summary>
    Private Shared Function LineTextMatchesExpected(t As String, expected As String) As Boolean
        If t.Length = 0 Then Return False
        If expected = ">" Then Return t = ">"
        If t.Equals(expected, StringComparison.OrdinalIgnoreCase) Then Return True
        If t.StartsWith(expected, StringComparison.OrdinalIgnoreCase) Then
            ' Prefix match only on a word boundary: "OKAY" must
            ' not match "OK", "READYX" must not match "READY".
            If t.Length = expected.Length Then Return True
            Dim nextCh As Char = t(expected.Length)
            If Not Char.IsLetterOrDigit(nextCh) Then Return True
        ElseIf t.EndsWith(" " & expected, StringComparison.OrdinalIgnoreCase) Then
            ' (Round 44) B-M1: the suffix match is an info-line
            ' value match ("+CPIN: READY") only when the text
            ' before the suffix ends with ":" - without this
            ' "+CPIN: NOT READY" (and bare "NOT READY") matched
            ' "READY", so CheckSim reported a not-ready SIM as
            ' READY. A bare word before the suffix ("NOT") is a
            ' different value, not the header of an info line.
            ' No other caller relies on a non-colon suffix: the
            ' expected values used here are "OK", "READY",
            ' "SIM PIN[2]" / "SIM PUK[2]" (all "+CPIN:"-headed)
            ' and ">" (exact-only) - no " NOT OK"-shaped line
            ' exists in AT replies.
            Dim prefix As String = t.Substring(0, t.Length - expected.Length - 1)
            If prefix.EndsWith(":") Then Return True
        End If
        Return False
    End Function

    ''' <summary>True when a reply line matches the expected text:
    ''' exact line, line starting with it, or line ending with it after
    ''' a space (e.g. "+CPIN: READY"). The "&gt;" prompt matches only a
    ''' line that is exactly "&gt;" after trim.</summary>
    Private Shared Function LineMatchesExpected(text As String, expected As String) As Boolean
        If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(expected) Then Return False
        For Each line As String In SplitLines(text)
            If LineTextMatchesExpected(line.Trim(), expected) Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 75) LIVE 2026-11-23: True when a reply
    ''' holds a terminal OK line — a line that is exactly "OK"
    ''' after trim. "OKAY" must not match (exact-line rule, not
    ''' the word-boundary rule of LineMatchesExpected): only a
    ''' bare OK proves the modem answered in command mode.</summary>
    Friend Shared Function ReplyHasTerminalOk(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each line As String In SplitLines(text)
            If line.Trim().Equals("OK", StringComparison.OrdinalIgnoreCase) Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 75) LIVE 2026-11-23: True when a reply
    ''' holds a terminal error line — a line starting with ERROR,
    ''' +CMS ERROR or +CME ERROR (line-anchored like
    ''' LineStartsWith: message text echoed back must never
    ''' count). An error anywhere beats an OK anywhere: a
    ''' refusal is known-not-sent, an OK beside it changes
    ''' nothing.</summary>
    Friend Shared Function ReplyHasTerminalError(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        Return LineStartsWith(text, "ERROR") OrElse LineStartsWith(text, "+CMS ERROR") OrElse LineStartsWith(text, "+CME ERROR")
    End Function

    ''' <summary>True when a reply line is exactly "&gt;" after trim.</summary>
    Private Shared Function HasPromptLine(text As String) As Boolean
        Return LineMatchesExpected(text, ">")
    End Function

    ''' <summary>FIX B-M1: expected-match on a whole reply that may hold
    ''' a late +CMGS: pair. A +CMGS: line and the OK that confirms it
    ''' (the first OK after it, skipping blank lines and URCs, before
    ''' any ERROR / +CMS ERROR / +CME ERROR / next +CMGS: line — the
    ''' same skip rule as CmgsLineConfirmations) belong to that confirmation,
    ''' NOT to the live command, so they can never satisfy the command's
    ''' expected-match. Every other line is matched with the exact
    ''' LineMatchesExpected rule. Use this — never LineMatchesExpected —
    ''' for the OK check on a SendCommand reply.</summary>
    Private Shared Function ReplyMatchesExpected(text As String, expected As String) As Boolean
        If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(expected) Then Return False
        Dim awaitingCmgsOk As Boolean = False
        For Each line As String In SplitLines(text)
            Dim t As String = line.Trim()
            If t.Length = 0 Then Continue For
            If t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                awaitingCmgsOk = True
                Continue For
            End If
            If awaitingCmgsOk Then
                If LineTextMatchesExpected(t, "OK") Then
                    awaitingCmgsOk = False
                    Continue For
                End If
                If t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then
                    awaitingCmgsOk = False
                Else
                    ' (Round 39) B-Low-1: a live non-OK/non-ERROR line
                    ' that matches expected satisfies the match at once
                    ' (e.g. a late ">" / "READY" inside the +CMGS:
                    ' window) instead of burning the full timeout.
                    If LineTextMatchesExpected(t, expected) Then Return True
                    Continue For
                End If
            End If
            If LineTextMatchesExpected(t, expected) Then Return True
        Next
        Return False
    End Function

    ''' <summary>FIX B-M1: expected-match inside SendCommand's live read.
    ''' text is the joined buffer (stale kept prefix + rejoined held
    ''' fragment + live bytes); staleLength is the stale/live boundary
    ''' (bytes before it arrived before the command was written). A line
    ''' is a candidate only when it starts at or after staleLength — a
    ''' line that starts before the boundary is stale, or straddles it
    ''' (e.g. a held "O" completed by a live "K" into a stale OK), and
    ''' must never satisfy the match. A +CMGS: line (joined, even a
    ''' straddling one) opens the confirmation window, and the OK that
    ''' confirms it is consumed as that confirmation, never as the
    ''' command's reply — same skip rule as ReplyMatchesExpected.
    ''' TRACED CASE (i): held "+CM" + live "GS: 12" + vbCrLf + "OK" + vbCrLf
    ''' during AT+CMGF=1 — the joined "+CMGS: 12" line opens the window,
    ''' the OK is consumed as its confirmation, the loop keeps waiting,
    ''' and only the modem's real CMGF "OK" (a fully-live line) matches.
    ''' If no real OK ever arrives, the loop times out and the caller
    ''' sees no match via ReplyMatchesExpected either.
    ''' FIX B-D3: a straddling line's LIVE suffix (bytes after
    ''' staleLength) is also evaluated: held "+CM" + live "OK" + vbCrLf
    ''' glues into "+CMOK", and the suffix "OK" is the live AT reply,
    ''' so it satisfies the probe instead of burning the timeout.
    ''' A suffix that merely completes a stale line (held "O" + live
    ''' "K" -> suffix "K") never matches, so case (i) still holds.</summary>
    Private Shared Function JoinedMatchesExpected(text As String, staleLength As Integer, expected As String, ByRef matchedViaStaleWindow As Boolean) As Boolean
        If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(expected) Then Return False
        matchedViaStaleWindow = False
        Dim awaitingCmgsOk As Boolean = False
        ' FIX B-L1 (Round 17): True only while the open confirmation
        ' window was opened by a FULLY-STALE +CMGS: line (the whole
        ' line arrived before this command was written). A live OK
        ' arriving in that state is the probe's own reply, not that
        ' line's confirmation - the append design in SendCommand
        ' already decides this (the kept line is appended as
        ' [False]) - so the loop exits on it instead of burning the
        ' whole timeout. A window opened by a live (or straddling)
        ' +CMGS: line keeps the old rule: its OK is its own
        ' confirmation and the loop waits for the probe's own OK.
        Dim awaitingStaleCmgsOk As Boolean = False
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> ChrW(13) AndAlso text(lineEnd) <> ChrW(10)
                lineEnd += 1
            End While
            Dim line As String = text.Substring(lineStart, lineEnd - lineStart)
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(10) Then nextPos += 1
            Dim t As String = line.Trim()
            If t.Length > 0 Then
                If lineStart < staleLength Then
                    ' Stale or straddling line: never satisfies — except
                    ' a straddling +CMGS: line, which still opens the
                    ' confirmation window for the OK that follows it.
                    If t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                        awaitingCmgsOk = True
                        ' Only a line that ENDS at or before the
                        ' boundary is fully stale: a straddling line
                        ' (lineEnd > staleLength) was completed by
                        ' live bytes, so its OK is its own
                        ' confirmation (traced case (i) in the summary
                        ' above) and the stale flag stays False.
                        awaitingStaleCmgsOk = (lineEnd <= staleLength)
                    ElseIf lineEnd > staleLength Then
                        ' FIX B-D3: the line straddles the stale/live
                        ' boundary - its LIVE suffix (bytes after
                        ' staleLength) is a live reply line glued onto
                        ' a held fragment (held "+CM" + live "OK" ->
                        ' "+CMOK"). The suffix alone is evaluated: a
                        ' live "OK" here satisfies the probe even
                        ' though the joined line never terminates the
                        ' fragment. A suffix that merely completes a
                        ' stale line (held "O" + live "K" -> "K")
                        ' never matches, so case (i) still holds.
                        ' When the +CMGS: window is open the suffix OK
                        ' is that confirmation, not the command's reply
                        ' — UNLESS the window was opened by a fully-stale
                        ' +CMGS: line (awaitingStaleCmgsOk): then the
                        ' live OK is the probe's own reply (FIX B-L1,
                        ' Round 17 rule, applied here in Round 21) and
                        ' the loop exits instead of burning the timeout.
                        Dim liveSuffix As String = text.Substring(staleLength, lineEnd - staleLength).Trim()
                        If liveSuffix.Length > 0 Then
                            If awaitingCmgsOk Then
                                If LineTextMatchesExpected(liveSuffix, "OK") Then
                                    If awaitingStaleCmgsOk Then matchedViaStaleWindow = True : Return True
                                    awaitingCmgsOk = False
                                    awaitingStaleCmgsOk = False
                                ElseIf liveSuffix.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                                    OrElse liveSuffix.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                                    OrElse liveSuffix.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then
                                    awaitingCmgsOk = False
                                    awaitingStaleCmgsOk = False
                                    If LineTextMatchesExpected(liveSuffix, expected) Then Return True
                                Else
                                    ' (Round 39) B-Low-1: a live non-OK/non-ERROR
                                    ' suffix that matches expected satisfies the
                                    ' match at once, same as the fully-live line.
                                    If LineTextMatchesExpected(liveSuffix, expected) Then Return True
                                End If
                            ElseIf LineTextMatchesExpected(liveSuffix, expected) Then
                                Return True
                            End If
                        End If
                    End If
                ElseIf t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                    awaitingCmgsOk = True
                    awaitingStaleCmgsOk = False
                ElseIf awaitingCmgsOk Then
                    If LineTextMatchesExpected(t, "OK") Then
                        ' FIX B-L1 (Round 17): fully-live OK while the
                        ' window was opened by a fully-stale +CMGS:
                        ' line - this OK is the probe's own reply (see
                        ' awaitingStaleCmgsOk above), so the live loop
                        ' exits here instead of burning the timeout.
                        ' A window opened by a live +CMGS: line keeps
                        ' the old rule: this OK is that line's own
                        ' confirmation and the loop waits for the
                        ' probe's own second OK (normal two-OK case).
                        If awaitingStaleCmgsOk Then matchedViaStaleWindow = True : Return True
                        awaitingCmgsOk = False
                        awaitingStaleCmgsOk = False
                    ElseIf t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                        OrElse t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                        OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then
                        awaitingCmgsOk = False
                        awaitingStaleCmgsOk = False
                        If LineTextMatchesExpected(t, expected) Then Return True
                    Else
                        ' (Round 39) B-Low-1: a live non-OK/non-ERROR line
                        ' that matches expected satisfies the match at once
                        ' (e.g. a late ">" / "READY" inside the +CMGS:
                        ' window) instead of burning the full timeout.
                        If LineTextMatchesExpected(t, expected) Then Return True
                    End If
                ElseIf LineTextMatchesExpected(t, expected) Then
                    Return True
                End If
            End If
            pos = nextPos
        End While
        Return False
    End Function

    ''' <summary>FIX B-M1: error-line check inside SendCommand's live
    ''' read — same stale/live boundary as JoinedMatchesExpected: only a
    ''' fully-live ERROR / +CMS ERROR / +CME ERROR line ends the read.
    ''' A stale fragment completed by live bytes (e.g. "ERR" + "OR")
    ''' never ends it.</summary>
    ''' <summary>(Round 46) B-M2: the +CMS ERROR skip rule, shared
    ''' by JoinedHasErrorLine and FirstLiveSkippedCmsErrorLineStart
    ''' so the two can never drift apart: a +CMS ERROR is skipped
    ''' (treated as a previous segment's late result) while a parked
    ''' +CMGS: line is outstanding - +CMS ERROR can only ever be
    ''' produced by an SMS submission, which is never a SendCommand,
    ''' except the prompt path. It is the command's own error only
    ''' when no parked line is outstanding, or when the command's
    ''' expected token has already (provisionally) arrived - for
    ''' expected=OK the provisional OK IS the expected token.</summary>
    Private Shared Function ShouldSkipCmsError(orphanedOkAllowance As Integer, provisionalOkLineStart As Integer, expected As String) As Boolean
        Return ((orphanedOkAllowance + If(provisionalOkLineStart >= 0, 1, 0)) > 0) AndAlso Not (provisionalOkLineStart >= 0 AndAlso String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase))
    End Function

    ''' <summary>(Round 46) B-M2: the line start of the first live
    ''' +CMS ERROR line that JoinedHasErrorLine SKIPS under the
    ''' rule above, or -1 when there is none (no parked line
    ''' outstanding, or no live +CMS ERROR line). Mirrors
    ''' JoinedHasErrorLine's scan exactly, including the
    ''' straddling case (a live error glued onto a held fragment).
    ''' Used only by the prompt-path grace in SendCommand: a
    ''' skipped +CMS ERROR there starts a short grace - '>'
    ''' arriving means the error was late, grace expiry with no
    ''' '>' means it was the command's own.</summary>
    Private Shared Function FirstLiveSkippedCmsErrorLineStart(text As String, staleLength As Integer, expected As String, orphanedOkAllowance As Integer, provisionalOkLineStart As Integer) As Integer
        If String.IsNullOrEmpty(text) Then Return -1
        If Not ShouldSkipCmsError(orphanedOkAllowance, provisionalOkLineStart, expected) Then Return -1
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(10)
                lineEnd += 1
            End While
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(10) Then nextPos += 1
            If lineStart >= staleLength Then
                If text.Substring(lineStart, lineEnd - lineStart).Trim().StartsWith("+CMS ERROR", System.StringComparison.OrdinalIgnoreCase) Then Return lineStart
            End If
            If lineStart < staleLength AndAlso lineEnd > staleLength Then
                Dim liveSuffix As String = text.Substring(staleLength, lineEnd - staleLength).Trim()
                If liveSuffix.StartsWith("+CMS ERROR", System.StringComparison.OrdinalIgnoreCase) Then Return lineStart
            End If
            pos = nextPos
        End While
        Return -1
    End Function

    Private Shared Function JoinedHasErrorLine(text As String, staleLength As Integer, expected As String, orphanedOkAllowance As Integer, provisionalOkLineStart As Integer) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        ' (Round 45) B-H3: a +CMS ERROR is LATE (a previous
        ' segment's late result) while a parked +CMGS: line is
        ' outstanding and the command's own reply has not
        ' arrived yet - +CMS ERROR can only ever be produced
        ' by an SMS submission, which is never a SendCommand,
        ' except the prompt path. It is the command's own
        ' error only when no parked line is outstanding, or
        ' when the command's expected token has already
        ' (provisionally) arrived - for expected=OK the
        ' provisional OK IS the expected token. +CME ERROR
        ' and bare ERROR always end the read, on every path.
        ' (Round 46) B-M2: the rule lives in ShouldSkipCmsError
        ' above, shared with FirstLiveSkippedCmsErrorLineStart.
        Dim skipCmsError As Boolean = ShouldSkipCmsError(orphanedOkAllowance, provisionalOkLineStart, expected)
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> ChrW(13) AndAlso text(lineEnd) <> ChrW(10)
                lineEnd += 1
            End While
            Dim line As String = text.Substring(lineStart, lineEnd - lineStart)
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(10) Then nextPos += 1
            If lineStart >= staleLength Then
                Dim t As String = line.Trim()
                If t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse (Not skipCmsError AndAlso t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase)) _
                    OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            End If
            If lineStart < staleLength AndAlso lineEnd > staleLength Then
                ' FIX B-D3: a live ERROR glued onto a held fragment
                ' (held "+CM" + live "ERROR") ends the read; a suffix
                ' that merely completes a stale fragment ("ERR" + "OR"
                ' -> "OR") never does.
                Dim liveSuffix As String = text.Substring(staleLength, lineEnd - staleLength).Trim()
                If liveSuffix.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse (Not skipCmsError AndAlso liveSuffix.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase)) _
                    OrElse liveSuffix.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            End If
            pos = nextPos
        End While
        Return False
    End Function

    ''' <summary>(Round 40) B-L3, reworked (Round 41) B-1: True when
    ''' text holds a fully-live OK line that is a definitive final
    ''' answer for a command whose expected token is not "OK" itself:
    ''' the OK line starts at or after staleLength (fully live, not a
    ''' stale fragment completed by live bytes), and NO +CMGS: line
    ''' claims it as its own confirmation. Claiming mirrors
    ''' CmgsLineConfirmations exactly: the forward scan from a +CMGS:
    ''' line skips blank lines, URCs (+CMTI:, +CREG:) and other
    ''' non-OK/non-ERROR lines, and stops at OK / ERROR / +CMS ERROR /
    ''' +CME ERROR / the next +CMGS: line - the OK it lands on is
    ''' that line's confirmation, never a definitive final OK. The
    ''' Round 40 rule (only the immediately-preceding non-blank line)
    ''' missed a URC or info line between the +CMGS: and its OK, so a
    ''' previous segment's very-late confirmation fired the
    ''' definitive-OK exit on the prompt path (the recovery drain then
    ''' sent ESC and the row failed spuriously) and on the CheckSim
    ''' path (SendCommand returned before +CPIN: arrived). A stale
    ''' +CMGS: line still claims a live OK (e2), exactly as before.
    ''' Traces: "+CPIN: SIM PIN + OK" expected READY -> True;
    ''' "+CPIN: READY + OK" -> True (also exits on READY); SIM PUK /
    ''' PIN2 / PUK2 + OK -> True; ERROR -> True via the error check,
    ''' not here; prompt "+CMGS: 12 + OK" -> False; "+CMGS: 12 +
    ''' +CMTI: + OK" / "+CREG:" / "+CPIN: SIM PIN" + OK -> False.</summary>
    Private Shared Function JoinedHasDefinitiveFinalOk(text As String, staleLength As Integer, command As String) As Boolean
        ' (Round 43) B-M1/B-L1/B-L2: same rule as before — an OK no
        ' +CMGS: line claims is definitive — now answered by the
        ' finder below, so the claim rule lives in exactly one place.
        Return FirstLiveUnclaimedOkLineStart(text, staleLength, staleLength, command) >= 0
    End Function

    ''' <summary>(Round 43) B-M1/B-L1/B-L2: the line START of the first
    ''' live unclaimed OK at or after fromPos (and at or after
    ''' staleLength), or -1 when there is none. The line collection
    ''' and the +CMGS:-claim marking are the Round 41 B-1 rule
    ''' verbatim (collect the lines with their start offsets, then
    ''' mark every OK a +CMGS: line claims under the
    ''' CmgsLineConfirmations forward scan — skip blank/URC/other
    ''' lines, stop at OK / ERROR / +CMS ERROR / +CME ERROR / next
    ''' +CMGS:). Only an OK no +CMGS: line claims is definitive.
    ''' fromPos is the re-arm floor: the buffer offset just past the
    ''' last swallowed OK, so a swallowed OK is never seen again.</summary>
    Private Shared Function FirstLiveUnclaimedOkLineStart(text As String, staleLength As Integer, fromPos As Integer, command As String) As Integer
        If String.IsNullOrEmpty(text) Then Return -1
        Dim starts As New System.Collections.Generic.List(Of Integer)()
        Dim trimmed As New System.Collections.Generic.List(Of String)()
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(10)
                lineEnd += 1
            End While
            starts.Add(lineStart)
            trimmed.Add(text.Substring(lineStart, lineEnd - lineStart).Trim())
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(10) Then nextPos += 1
            pos = nextPos
        End While
        Dim claimedOk As New System.Collections.Generic.List(Of Boolean)()
        For i As Integer = 0 To trimmed.Count - 1
            claimedOk.Add(False)
        Next
        For i As Integer = 0 To trimmed.Count - 1
            If Not trimmed(i).StartsWith("+CMGS:", System.StringComparison.OrdinalIgnoreCase) Then Continue For
            ' (Round 45) B-M1: a LIVE +CMGS: line's forward scan
            ' stops at the command's own info line - an OK after
            ' +CPIN: ... is the command's final OK, not that
            ' line's confirmation. A STALE line keeps the old
            ' scan (it still claims a live OK past an info line
            ' - the Round-41 e2 rule), so the dl2/c2 provisional
            ' cases are untouched.
            Dim cmgsIsLive As Boolean = starts(i) >= staleLength
            Dim infoPrefix As String = CommandInfoLinePrefix(command)
            For j As Integer = i + 1 To trimmed.Count - 1
                Dim u As String = trimmed(j)
                If u.Length = 0 Then Continue For
                If cmgsIsLive AndAlso infoPrefix.Length > 0 AndAlso u.StartsWith(infoPrefix, System.StringComparison.OrdinalIgnoreCase) Then Exit For
                If LineTextMatchesExpected(u, "OK") Then
                    claimedOk(j) = True
                    Exit For
                End If
                If u.StartsWith("ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMS ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CME ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMGS:", System.StringComparison.OrdinalIgnoreCase) Then Exit For
            Next
        Next
        For i As Integer = 0 To trimmed.Count - 1
            If starts(i) >= staleLength AndAlso starts(i) >= fromPos AndAlso Not claimedOk(i) AndAlso LineTextMatchesExpected(trimmed(i), "OK") Then Return starts(i)
        Next
        Return -1
    End Function

    ''' <summary>(Round 43) B-M1/B-L1/B-L2: the buffer offset just past
    ''' the line starting at lineStart (its terminator included, when
    ''' one has arrived) — the next scan floor after swallowing that
    ''' OK. An unterminated trailing line ends at text.Length.</summary>
    Private Shared Function OkLineEndAfter(text As String, lineStart As Integer) As Integer
        Dim p As Integer = lineStart
        While p < text.Length AndAlso text(p) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(p) <> Microsoft.VisualBasic.ChrW(10)
            p += 1
        End While
        If p < text.Length AndAlso text(p) = Microsoft.VisualBasic.ChrW(13) Then p += 1
        If p < text.Length AndAlso text(p) = Microsoft.VisualBasic.ChrW(10) Then p += 1
        Return p
    End Function

    ''' <summary>(Round 45) B-H1: text with the line starting at
    ''' lineStart blanked (characters replaced by spaces, the
    ''' terminator kept) - every offset in the text stays valid.
    ''' A provisionally-swallowed OK is consumed for match
    ''' purposes while it is pending, so JoinedMatchesExpected
    ''' can never re-match that same OK as the command's own
    ''' reply on a later chunk.</summary>
    Private Shared Function TextWithLineBlanked(text As String, lineStart As Integer) As String
        If String.IsNullOrEmpty(text) OrElse lineStart < 0 OrElse lineStart >= text.Length Then Return text
        Dim lineEnd As Integer = lineStart
        While lineEnd < text.Length AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(10)
            lineEnd += 1
        End While
        Dim sb As New System.Text.StringBuilder(text)
        For p As Integer = lineStart To lineEnd - 1
            sb(p) = " "c
        Next
        Return sb.ToString()
    End Function

    ''' <summary>(Round 45) B-H1: text with the lines starting
    ''' at the given offsets removed (line + terminator). The
    ''' offsets are into text; lines are removed highest-first
    ''' so earlier offsets stay valid.</summary>
    Private Shared Function RemoveLinesAt(text As String, lineStarts As System.Collections.Generic.List(Of Integer)) As String
        If String.IsNullOrEmpty(text) OrElse lineStarts Is Nothing OrElse lineStarts.Count = 0 Then Return text
        Dim sorted As New System.Collections.Generic.List(Of Integer)(lineStarts)
        sorted.Sort()
        sorted.Reverse()
        Dim result As String = text
        For Each s As Integer In sorted
            If s < 0 OrElse s >= result.Length Then Continue For
            Dim e As Integer = s
            While e < result.Length AndAlso result(e) <> Microsoft.VisualBasic.ChrW(13) AndAlso result(e) <> Microsoft.VisualBasic.ChrW(10)
                e += 1
            End While
            If e < result.Length AndAlso result(e) = Microsoft.VisualBasic.ChrW(13) Then e += 1
            If e < result.Length AndAlso result(e) = Microsoft.VisualBasic.ChrW(10) Then e += 1
            result = result.Substring(0, s) & result.Substring(e)
        Next
        Return result
    End Function

    ''' <summary>(Round 45) B-M1: how many LIVE +CMGS: lines
    ''' before the OK at okLineStart would have claimed that OK
    ''' under the old claim rule (forward scan past blank / URC /
    ''' other lines, including the command's own info line) but
    ''' do NOT claim it under the new rule (the scan stops at
    ''' the command's own info line). Each such line is a
    ''' previous segment's late confirmation whose OK never
    ''' arrived as its own, so the caller parks it unconfirmed
    ''' instead of letting it evaporate. Only LIVE lines count:
    ''' a stale line keeps the old scan, so a stale line that
    ''' would claim the OK does claim it, and the OK would not
    ''' be unclaimed at all.</summary>
    Private Shared Function CountDisclaimedLiveCmgsBeforeOk(text As String, staleLength As Integer, okLineStart As Integer, command As String) As Integer
        If String.IsNullOrEmpty(text) OrElse okLineStart < 0 Then Return 0
        Dim infoPrefix As String = CommandInfoLinePrefix(command)
        If infoPrefix.Length = 0 Then Return 0
        Dim starts As New System.Collections.Generic.List(Of Integer)()
        Dim trimmed As New System.Collections.Generic.List(Of String)()
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(10)
                lineEnd += 1
            End While
            starts.Add(lineStart)
            trimmed.Add(text.Substring(lineStart, lineEnd - lineStart).Trim())
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(10) Then nextPos += 1
            pos = nextPos
        End While
        Dim count As Integer = 0
        For i As Integer = 0 To trimmed.Count - 1
            If starts(i) < staleLength OrElse starts(i) >= okLineStart Then Continue For
            If Not trimmed(i).StartsWith("+CMGS:", System.StringComparison.OrdinalIgnoreCase) Then Continue For
            For j As Integer = i + 1 To trimmed.Count - 1
                Dim u As String = trimmed(j)
                If u.Length = 0 Then Continue For
                If LineTextMatchesExpected(u, "OK") Then
                    If starts(j) = okLineStart Then
                        Dim crossedInfo As Boolean = False
                        For k As Integer = i + 1 To j - 1
                            If trimmed(k).StartsWith(infoPrefix, System.StringComparison.OrdinalIgnoreCase) Then crossedInfo = True : Exit For
                        Next
                        If crossedInfo Then count += 1
                    End If
                    Exit For
                End If
                If u.StartsWith("ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMS ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CME ERROR", System.StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMGS:", System.StringComparison.OrdinalIgnoreCase) Then Exit For
            Next
        Next
        Return count
    End Function

    ''' <summary>(Round 44) B-M2: the info-line prefix of THIS
    ''' command's own reply ("+CPIN:" for "AT+CPIN?", "+CSQ:"
    ''' for "AT+CSQ"), derived from the AT command name — or
    ''' String.Empty when the command has no info line.
    ''' "+CMGS:" is NEVER an info line: on the AT+CMGS prompt
    ''' path a +CMGS: line is a late confirmation of a previous
    ''' segment, not this command's reply (its own reply is the
    ''' "&gt;" prompt), so the prompt path has no info lines at
    ''' all. An empty result means NO line counts as info for
    ''' this command, so any OK before the reply is swallowed
    ''' while the allowance lasts. This replaces the Round 43
    ''' IsUnsolicitedResultLine denylist (deleted): the URC set
    ''' is open-ended by nature, so denylisting URCs cannot
    ''' close it — allowlisting the command's own info lines
    ''' can, because any line that is not one of them (a URC
    ''' inside or outside any set, an unknown "+XXX:" line, the
    ''' echo, a +CMGS: line, a final-result line) is simply not
    ''' info.</summary>
    Private Shared Function CommandInfoLinePrefix(command As String) As String
        If String.IsNullOrWhiteSpace(command) Then Return String.Empty
        Dim c As String = command.Trim()
        If Not c.StartsWith("AT+", StringComparison.OrdinalIgnoreCase) Then Return String.Empty
        Dim i As Integer = 3
        While i < c.Length AndAlso ((c(i) >= "A"c AndAlso c(i) <= "Z"c) OrElse (c(i) >= "a"c AndAlso c(i) <= "z"c))
            i += 1
        End While
        If i = 3 Then Return String.Empty
        Dim name As String = c.Substring(3, i - 3).ToUpperInvariant()
        If name = "CMGS" Then Return String.Empty
        ' (Round 45) B-L1: these commands have no info line -
        ' CMGF/CSCS are set commands answering bare OK, and
        ' CGMI/CGMM/CGMR are TS 27.007 identification commands
        ' answering plain text, with no +CGMI: prefix. Deriving
        ' a +NAME: prefix for them invents one, so a future
        ' non-OK caller would misclassify an unrelated +XXX:
        ' line as its info line.
        If name = "CMGF" OrElse name = "CSCS" OrElse name = "CGMI" OrElse name = "CGMM" OrElse name = "CGMR" Then Return String.Empty
        Return "+" & name & ":"
    End Function

    ''' <summary>(Round 44) B-M2: True when one of THIS command's
    ''' own known info lines (see CommandInfoLinePrefix — "+CPIN:"
    ''' for AT+CPIN?, "+CSQ:" for AT+CSQ, none for the AT+CMGS
    ''' prompt path) precedes the live OK starting at okLineStart
    ''' (lines from staleLength onward only — stale lines never
    ''' count). The Round 43 version treated ANY non-URC line as
    ''' the command's info line, so a URC outside its 7-item set
    ''' (+CUSD:, +CDSI:, +CIEV:, +CRING:, ...) blocked the
    ''' swallow and the orphaned OK exited the loop early as if
    ''' it were the command's own. Now ONLY the command's own
    ''' info prefix counts, and every other line (a URC inside
    ''' or outside any set, an unknown "+XXX:" line, the echo,
    ''' a +CMGS: line, a final-result line, the ">" prompt, a
    ''' blank line) is simply not info. A line that is an info
    ''' line for ANOTHER command ("+CSQ:" during AT+CPIN?)
    ''' does not count either.</summary>
    Private Shared Function LiveOkPrecededByInfoLine(text As String, staleLength As Integer, okLineStart As Integer, command As String) As Boolean
        Dim prefix As String = CommandInfoLinePrefix(command)
        If prefix.Length = 0 Then Return False
        Dim pos As Integer = 0
        While pos < text.Length AndAlso pos < okLineStart
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(13) AndAlso text(lineEnd) <> Microsoft.VisualBasic.ChrW(10)
                lineEnd += 1
            End While
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = Microsoft.VisualBasic.ChrW(10) Then nextPos += 1
            pos = nextPos
            If lineStart < staleLength Then Continue While
            Dim t As String = text.Substring(lineStart, lineEnd - lineStart).Trim()
            If t.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return True
        End While
        Return False
    End Function

    ''' <summary>FIX B-L1: splits ANY trailing unterminated line off
    ''' the end of text (returned in fragment, removed from text) — the
    ''' same hold-back StripStaleFinalResultLines applies (B-B4), but
    ''' for the drain path. A trailing partial "O" (of OK), "ERR" (of
    ''' ERROR) or "+CM" (of +CMGS:) is not a complete line yet, so it
    ''' must not be returned as one: the caller re-preserves it in
    ''' _preservedText for the next read. After this call text always
    ''' ends at a line boundary (or is empty).</summary>
    Private Shared Sub SplitTrailingUnterminatedLine(ByRef text As String, ByRef fragment As String)
        fragment = String.Empty
        If String.IsNullOrEmpty(text) Then Return
        If text.EndsWith(vbCr) OrElse text.EndsWith(vbLf) Then Return
        Dim lastBreak As Integer = text.LastIndexOfAny(New Char() {ChrW(13), ChrW(10)})
        If lastBreak < 0 Then
            fragment = text
            text = String.Empty
        Else
            fragment = text.Substring(lastBreak + 1)
            text = text.Substring(0, lastBreak + 1)
        End If
    End Sub

    ''' <summary>FIX B-L3: True when text is a +CMGS:-related fragment:
    ''' a +CMGS: line ("+CMGS: 12") or a prefix of one ("+CM", "+CMGS:").
    ''' FIX D-L2: the single predicate - SplitTrailingCmgsFragment
    ''' and DropStraddlingFirstLineUnlessCmgs both call this now,
    ''' instead of inlining the same checks.</summary>
    Private Shared Function IsCmgsFragmentText(text As String) As Boolean
        Dim t As String = text.TrimStart()
        If t.Length = 0 Then Return False
        Return t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) _
            OrElse "+CMGS:".StartsWith(t, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>FIX B-D4: True only when text is already an
    ''' unambiguous +CMGS: fragment: it starts with "+CMGS:"
    ''' (colon included) — or is the bare "+CMGS" itself (FIX B-L2:
    ''' the colon has simply not arrived yet; "+CMGS" can only become
    ''' "+CMGS:", since no other modem reply line starts "+CMGS":
    ''' "+CMGL:" starts "+CMGL", "+CMS ERROR" starts "+CMS "). A
    ''' bare "+CM" / "+CMG" / "+C" prefix is equally a prefix of
    ''' "+CMS ERROR", "+CMTI:", "+CMT:" and "+CMGL:", so it must
    ''' never be flushed as a confirmation.
    ''' Used ONLY by FlushStrandedCmgsFragment; the hold-back
    ''' paths keep using IsCmgsFragmentText (a prefix that MIGHT
    ''' become +CMGS: must still be held back).</summary>
    Private Shared Function IsUnambiguousCmgsFragmentText(text As String) As Boolean
        Dim t As String = text.TrimStart()
        If t.Length = 0 Then Return False
        If t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Return True
        ' FIX B-L2: accept a fragment that is a prefix of "+CMGS:" of
        ' length >= 5 (i.e. "+CMGS" itself), while still rejecting
        ' "+CM", "+CMG" and "+C".
        Return t.Length >= 5 AndAlso "+CMGS:".StartsWith(t, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>FIX B-M1/B-L1: splits a trailing unterminated
    ''' +CMGS:-related fragment off the end of text (the fragment is
    ''' returned in fragment and removed from text). "Related" is a
    ''' +CMGS: line ("+CMGS: 12") or a prefix of one ("+CM", "+CMGS:").
    ''' Anything else — including a partial "O" — is left in place.
    ''' Callers re-preserve the fragment in _preservedText so the next
    ''' read can complete it, instead of counting it as a whole line.</summary>
    Private Shared Sub SplitTrailingCmgsFragment(ByRef text As String, ByRef fragment As String)
        fragment = String.Empty
        If String.IsNullOrEmpty(text) Then Return
        If text.EndsWith(vbCr) OrElse text.EndsWith(vbLf) Then Return
        Dim lastBreak As Integer = text.LastIndexOfAny(New Char() {ChrW(13), ChrW(10)})
        Dim trailing As String = If(lastBreak < 0, text, text.Substring(lastBreak + 1))
        ' FIX D-L2: one predicate - call IsCmgsFragmentText instead
        ' of inlining the same checks here (the two copies could
        ' drift). TrimStart handling is identical: both trim leading
        ' whitespace before the StartsWith tests.
        If IsCmgsFragmentText(trailing) Then
            fragment = trailing
            text = If(lastBreak < 0, String.Empty, text.Substring(0, lastBreak + 1))
        End If
    End Sub

    ''' <summary>FIX B-M1: drops the FIRST line of text when it straddles
    ''' a read boundary (a held fragment completed by live bytes) and is
    ''' NOT a +CMGS: line — it is a stale line (e.g. "O" + "K" into a
    ''' stale OK) and must not reach the caller. A +CMGS:-related first
    ''' line (a +CMGS: line, or a prefix of one) is kept: the caller
    ''' counts/parks it, and SplitTrailingCmgsFragment holds back an
    ''' unterminated one. Call only when a fragment was actually held.</summary>
    Private Shared Function DropStraddlingFirstLineUnlessCmgs(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        Dim firstBreak As Integer = text.IndexOfAny(New Char() {ChrW(13), ChrW(10)})
        Dim firstLine As String = If(firstBreak < 0, text, text.Substring(0, firstBreak))
        ' FIX D-L2: same single predicate as SplitTrailingCmgsFragment.
        If IsCmgsFragmentText(firstLine) Then Return text
        If firstBreak < 0 Then Return String.Empty
        Dim rest As Integer = firstBreak
        If rest < text.Length AndAlso text(rest) = ChrW(13) Then rest += 1
        If rest < text.Length AndAlso text(rest) = ChrW(10) Then rest += 1
        Return text.Substring(rest)
    End Function

    ''' <summary>FIX B-D3: repairs a reply whose first line is a held
    ''' fragment glued onto a live reply line (held "+CM" + live
    ''' "OK" + vbCrLf -> "+CMOK"), after
    ''' DropStraddlingFirstLineUnlessCmgs has dropped that glued
    ''' line. Two cases, in order: (a) the fragment's continuation
    ''' is the NEXT line in this reply - held "+CM" + next line
    ''' "GS: 12" is the delayed +CMGS: line, split by the
    ''' intervening live reply, so it is rejoined in place and the
    ''' caller can still park/count it; (b) no continuation is in
    ''' this reply yet - the fragment is re-preserved for the next
    ''' read (only +CMGS:-prefix fragments; a stale "O" stays
    ''' dropped as before), so a later "GS: 12" can still complete
    ''' it instead of the +CMGS: being lost. Returns the repaired
    ''' reply; when the first line was a kept +CMGS: line the reply
    ''' is returned unchanged. FIX B-H1: the dropped glued line's live
    ''' suffix is preserved as its own line when it is itself a
    ''' final-result line (a live OK glued onto the fragment), so the
    ''' probe's proof of command mode is not lost with the stale fragment.</summary>
    Private Function RejoinGluedFragmentReply(reply As String, heldFragment As String) As String
        If String.IsNullOrEmpty(heldFragment) OrElse String.IsNullOrEmpty(reply) Then Return reply
        Dim dropped As String = DropStraddlingFirstLineUnlessCmgs(reply)
        If dropped = reply Then Return reply
        ' FIX B-H1: the dropped glued line's LIVE suffix (the bytes
        ' after the held fragment) may itself be a genuine
        ' final-result line — held "+CM" + live "OK" glues into
        ' "+CMOK", and that live OK is the probe's proof of command
        ' mode. Preserve it as its own line, or the proof is dropped
        ' with the stale fragment and the B-B6 gate later fails a row
        ' whose segment was already counted. A suffix that merely
        ' completes the stale line (held "O" + live "K" -> "K") is
        ' never a final-result line, so it stays dropped. Sub-case
        ' (b) is not made worse: the fragment is still re-preserved
        ' below, and a live OK in this reply is now kept with it.
        Dim liveSuffixLine As String = String.Empty
        Dim replyFirstBreak As Integer = reply.IndexOfAny(New Char() {ChrW(13), ChrW(10)})
        Dim replyFirstLine As String = If(replyFirstBreak < 0, reply, reply.Substring(0, replyFirstBreak))
        If replyFirstLine.StartsWith(heldFragment, StringComparison.Ordinal) Then
            Dim liveSuffix As String = replyFirstLine.Substring(heldFragment.Length).Trim()
            If liveSuffix.Length > 0 AndAlso IsFinalResultLineText(liveSuffix) Then liveSuffixLine = liveSuffix & vbCrLf
        End If
        If dropped.Length > 0 Then
            Dim firstBreak As Integer = dropped.IndexOfAny(New Char() {ChrW(13), ChrW(10)})
            Dim firstLine As String = If(firstBreak < 0, dropped, dropped.Substring(0, firstBreak))
            Dim rejoined As String = heldFragment & firstLine
            ' Only a PROPER prefix may rejoin: a held fragment that
            ' already contains the colon ("+CMGS:") cannot gain a
            ' "GS: ..." continuation, so it is re-preserved below
            ' instead of forging a "+CMGS:GS: ..." line.
            If IsCmgsFragmentText(heldFragment) _
                AndAlso Not heldFragment.TrimStart().StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) _
                AndAlso rejoined.TrimStart().StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                Return liveSuffixLine & rejoined & If(firstBreak < 0, String.Empty, dropped.Substring(firstBreak))
            End If
        End If
        If IsCmgsFragmentText(heldFragment) Then PrependPreservedFragment(heldFragment)
        Return liveSuffixLine & dropped
    End Function

    ''' <summary>FIX B-M2: re-preserves a trailing fragment in
    ''' _preservedText, separated from anything already there by a
    ''' line break — a bare concatenation forges one bogus line
    ''' ("+CM" + "+CMGS: 9" -> "+CM+CMGS: 9"), losing the
    ''' confirmation permanently. _preservedText is empty in the
    ''' common case, so this is a no-op there. (Round 54)</summary>
    Private Sub AppendPreservedFragment(fragment As String)
        If String.IsNullOrEmpty(fragment) Then Return
        If _preservedText <> String.Empty Then _preservedText &= vbCrLf
        _preservedText &= fragment
    End Sub

    ''' <summary>
    ''' FIX B-M2 (Round 14 hardening): the PREPEND half of
    ''' AppendPreservedFragment — re-preserves a fragment BEFORE
    ''' anything already in _preservedText (chronological order),
    ''' separated by a line break when both are present, so a bare
    ''' concatenation can never glue two fragments into one bogus
    ''' line. Same pattern as DrainResult's inline prepend. No
    ''' duplicate guard is added: the two call sites (Rejoin and the
    ''' pre-probe hold-back) have none today, and _preservedText is
    ''' provably empty at both of them (each follows a
    ''' TakePreservedText with no _preservedText write between), so
    ''' routing them through this helper changes no behaviour today
    ''' — it only makes the separator structural for a future edit.
    ''' </summary>
    Private Sub PrependPreservedFragment(fragment As String)
        If String.IsNullOrEmpty(fragment) Then Return
        If _preservedText <> String.Empty Then _preservedText = fragment & vbCrLf & _preservedText Else _preservedText = fragment
    End Sub

    ''' <summary>FIX B-L1/B-L4: the drain's return path. Applies the
    ''' generalised trailing-fragment hold-back (ANY unterminated line,
    ''' as Strip does in B-B4 — not only +CMGS: fragments): the
    ''' fragment is re-preserved in _preservedText for the next read
    ''' instead of being returned (and counted) as a complete line, so
    ''' the returned text always ends at a line boundary. FIX B-L4:
    ''' the drain fragment is PREPENDED before any fragment already in
    ''' _preservedText (chronological order — appending it after the
    ''' probe reply's fragment inverted the order, e.g. "+CM+CMGS: 12"),
    ''' and a fragment already present is not preserved twice.</summary>
    Private Function DrainResult(drained As System.Text.StringBuilder) As String
        Dim text As String = drained.ToString()
        Dim fragment As String = String.Empty
        SplitTrailingUnterminatedLine(text, fragment)
        ' FIX B-M2 (audit): the same bare-concatenation forgery as the
        ' return paths — an unterminated drain fragment glued onto a
        ' fragment already in _preservedText forges one bogus line.
        ' The EndsWith duplicate guard and the chronological prepend
        ' order both still hold; only a line break is inserted between
        ' the two fragments when both are present.
        If fragment <> String.Empty AndAlso Not _preservedText.EndsWith(fragment, StringComparison.Ordinal) Then
            If _preservedText <> String.Empty Then _preservedText = fragment & vbCrLf & _preservedText Else _preservedText = fragment
        End If
        Return text
    End Function

    ''' <summary>Counts the +CMGS: lines in one portion of text and,
    ''' in confirmed, how many of them have their own trailing OK: an
    ''' OK line after the +CMGS: line, before any ERROR / +CMS ERROR /
    ''' +CME ERROR line or another +CMGS: line. The OK match is the
    ''' same matcher as LineMatchesExpected (not a strict Equals), so
    ''' e.g. a line ending in " OK" still confirms. An OK line before
    ''' the +CMGS: never counts.</summary>
    ''' <summary>B-B1/B-B5: one confirmation flag per +CMGS: line in
    ''' arrival order — True when that line has its own trailing OK
    ''' (an OK after it, before any ERROR / +CMS ERROR / +CME ERROR
    ''' / next +CMGS: line; other lines such as URCs are skipped,
    ''' exactly as CmgsLineConfirmations always did). The current segment's
    ''' confirmation is the line with the GREATEST message reference
    ''' (see CmgsLineReferences / SplitCmgsFlags - last-confirmed
    ''' fallback when refs are missing or shorter than flags, no
    ''' parseable mr at all, or a tie on the greatest mr - a window
    ''' spanning the 255-&gt;0 wrap is unwrapped first (FIX B-L2,
    ''' Round 24);
    ''' earlier lines belong to previous segments/rows and must be
    ''' parked, not counted.</summary>
    Private Shared Function CmgsLineConfirmations(text As String) As List(Of Boolean)
        Dim flags As New List(Of Boolean)()
        If String.IsNullOrEmpty(text) Then Return flags
        Dim lines As String() = SplitLines(text)
        For i As Integer = 0 To lines.Length - 1
            Dim t As String = lines(i).Trim()
            If t.Length = 0 Then Continue For
            If Not t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim confirmed As Boolean = False
            For j As Integer = i + 1 To lines.Length - 1
                Dim u As String = lines(j).Trim()
                If u.Length = 0 Then Continue For
                If LineMatchesExpected(u, "OK") Then
                    confirmed = True
                    Exit For
                End If
                If u.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) _
                    OrElse u.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Exit For
            Next
            flags.Add(confirmed)
        Next
        Return flags
    End Function

    ''' <summary>B-B1/B-B5: one message reference per +CMGS: line in
    ''' arrival order, parallel to CmgsLineConfirmations: the number
    ''' after "+CMGS:" (e.g. "+CMGS: 13" -&gt; 13), or -1 when the
    ''' line carries no parseable number. References increment by one
    ''' per segment (traces show 11 -&gt; 12 -&gt; 13), so the greatest
    ''' reference in a window is the current segment's line. The
    ''' reference is a byte (0-255) and wraps 255 -&gt; 0, so a window
    ''' spanning the wrap is unwrapped relative to the first
    ''' parseable ref before ordering (FIX B-L2, Round 24 - see
    ''' CountedCmgsIndex). Only SplitCmgsFlags consumes
    ''' this list, at the same three call sites as the flags.
    ''' CmgsLineConfirmations is untouched: that function and the
    ''' kept-flags use keep working on flags alone.</summary>
    Private Shared Function CmgsLineReferences(text As String) As List(Of Integer)
        Dim refs As New List(Of Integer)()
        If String.IsNullOrEmpty(text) Then Return refs
        Dim lines As String() = SplitLines(text)
        For Each line As String In lines
            Dim t As String = line.Trim()
            If t.Length = 0 Then Continue For
            If Not t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Continue For
            ' FIX B-M1b (Round 18, proven by simulation): TryParse sets
            ' its out parameter to 0 on failure, so parsing straight
            ' into mr would turn an unparseable line ("+CMGS: abc")
            ' into mr 0 - a REAL reference (the post-wrap value). Parse
            ' into a scratch variable and copy only on success, so mr
            ' stays -1 (unparseable) unless the digits really parsed.
            Dim mr As Integer = -1
            Dim rest As String = t.Substring(6).Trim()
            If rest.Length > 0 Then
                Dim parts As String() = rest.Split(New Char() {" "c, ChrW(9)}, StringSplitOptions.RemoveEmptyEntries)
                Dim parsedMr As Integer = 0
                If parts.Length > 0 AndAlso Integer.TryParse(parts(0), parsedMr) Then mr = parsedMr
            End If
            refs.Add(mr)
        Next
        Return refs
    End Function

    ''' <summary>FIX B-M1 (Round 22): the index of the COUNTED +CMGS:
    ''' line (the current segment's) among the flags - the one with
    ''' the GREATEST message reference. Moved verbatim out of
    ''' SplitCmgsFlags so the grace checks share the counting path's
    ''' definition of "current". Returns -1 when no greatest
    ''' reference can be picked (refs missing or shorter than flags,
    ''' no parseable mr at all, or a tie on the greatest mr) - the
    ''' caller then falls back to the last-confirmed rule, exactly
    ''' as SplitCmgsFlags did. A window spanning the 255 -&gt; 0 wrap
    ''' is unwrapped relative to the first parseable ref before
    ''' the greatest rule is applied (FIX B-L2, Round 24), so it
    ''' no longer falls back.</summary>
    Private Shared Function CountedCmgsIndex(flags As List(Of Boolean), refs As List(Of Integer)) As Integer
        If refs Is Nothing OrElse refs.Count <> flags.Count Then Return -1
        Dim bestMr As Integer = -1
        Dim bestIndex As Integer = -1
        Dim bestTied As Boolean = False
        Dim minMr As Integer = Integer.MaxValue
        Dim maxMr As Integer = Integer.MinValue
        Dim parseableCount As Integer = 0
        For k As Integer = 0 To refs.Count - 1
            If refs(k) < 0 Then Continue For
            parseableCount += 1
            If refs(k) < minMr Then minMr = refs(k)
            If refs(k) > maxMr Then maxMr = refs(k)
            If refs(k) > bestMr Then
                bestMr = refs(k)
                bestIndex = k
                bestTied = False
            ElseIf refs(k) = bestMr Then
                bestTied = True
            End If
        Next
        ' A window spanning the 255 -&gt; 0 wrap holds both a high mr
        ' and a low one (max-min far beyond any real window of
        ' consecutive references): the greatest RAW mr there is
        ' stale, so the raw ordering cannot be used.
        ' FIX B-L2 (Round 24): unwrap instead of falling back -
        ' refs arrive in generation order on a serial line and mr
        ' increments by 1 per segment, so unwrapping every ref
        ' relative to the FIRST parseable ref (eff = anchor +
        ' ((r - anchor) mod 256), anchor = first parseable ref)
        ' restores the true order and the SAME greatest/tie rule
        ' applies on the unwrapped values. Unparseable refs are
        ' skipped exactly as in the loop above. The double Mod
        ' keeps VB's signed Mod non-negative. A tie on the
        ' greatest unwrapped value still returns -1, so the
        ' last-confirmed fallback is preserved there.
        Dim spansWrap As Boolean = parseableCount > 1 AndAlso (maxMr - minMr) > 127
        If spansWrap Then
            Dim anchor As Integer = -1
            For k As Integer = 0 To refs.Count - 1
                If refs(k) >= 0 Then
                    anchor = refs(k)
                    Exit For
                End If
            Next
            Dim bestEff As Integer = -1
            Dim bestEffIndex As Integer = -1
            Dim bestEffTied As Boolean = False
            For k As Integer = 0 To refs.Count - 1
                If refs(k) < 0 Then Continue For
                Dim eff As Integer = anchor + (((refs(k) - anchor) Mod 256 + 256) Mod 256)
                If eff > bestEff Then
                    bestEff = eff
                    bestEffIndex = k
                    bestEffTied = False
                ElseIf eff = bestEff Then
                    bestEffTied = True
                End If
            Next
            If Not bestEffTied Then Return bestEffIndex
            Return -1
        End If
        If parseableCount > 0 AndAlso Not bestTied Then
            Return bestIndex
        End If
        Return -1
    End Function

    ''' <summary>FIX B-M1 (Round 22): the counted-line index with the
    ''' Round 17 last-confirmed fallback (last CONFIRMED line; the
    ''' last line when nothing is confirmed) applied when the
    ''' greatest-mr pick is unavailable - the same selection
    ''' SplitCmgsFlags makes, shared by the grace checks below.</summary>
    Private Shared Function CountedCmgsIndexOrFallback(flags As List(Of Boolean), refs As List(Of Integer)) As Integer
        Dim idx As Integer = CountedCmgsIndex(flags, refs)
        If idx >= 0 Then Return idx
        If flags.Count = 0 Then Return -1
        For k As Integer = flags.Count - 1 To 0 Step -1
            If flags(k) Then Return k
        Next
        Return flags.Count - 1
    End Function

    ''' <summary>FIX B-M1 (Round 22): True when the CURRENT +CMGS:
    ''' line (greatest mr; last-confirmed fallback) has its trailing
    ''' OK. The old grace checks asked whether ANY line had an OK, so
    ''' a very-late confirmation of a PREVIOUS segment skipped the
    ''' grace for the current line and the row was falsely reported
    ''' "(M-XM-(M-XM-/M-YM-^HM-YM-^F M-XM-*M-XM-#M-YM-^CM-YM-^JM-XM-/ OK)".</summary>
    Private Shared Function CurrentCmgsConfirmed(text As String) As Boolean
        Dim flags As List(Of Boolean) = CmgsLineConfirmations(text)
        If flags.Count = 0 Then Return False
        Dim idx As Integer = CountedCmgsIndexOrFallback(flags, CmgsLineReferences(text))
        If idx < 0 OrElse idx >= flags.Count Then Return False
        Return flags(idx)
    End Function

    ''' <summary>FIX B-M1 (Round 22): True when an ERROR / +CMS ERROR /
    ''' +CME ERROR line trails the CURRENT +CMGS: line - i.e. it is the
    ''' first non-blank line after it, before any OK and before the
    ''' next +CMGS: line (the same scan CmgsLineConfirmations makes).
    ''' An ERROR belonging only to an EARLIER line returns False, so
    ''' it no longer suppresses the current line's grace; an ERROR
    ''' trailing the current line still ends the grace at once.</summary>
    Private Shared Function CurrentCmgsHasTrailingError(text As String) As Boolean
        Dim flags As List(Of Boolean) = CmgsLineConfirmations(text)
        If flags.Count = 0 Then Return False
        Dim idx As Integer = CountedCmgsIndexOrFallback(flags, CmgsLineReferences(text))
        If idx < 0 Then Return False
        Dim lines As String() = SplitLines(text)
        Dim lineIdx As Integer = -1
        Dim seen As Integer = -1
        For i As Integer = 0 To lines.Length - 1
            If lines(i).Trim().StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then
                seen += 1
                If seen = idx Then
                    lineIdx = i
                    Exit For
                End If
            End If
        Next
        If lineIdx < 0 Then Return False
        For j As Integer = lineIdx + 1 To lines.Length - 1
            Dim u As String = lines(j).Trim()
            If u.Length = 0 Then Continue For
            If LineMatchesExpected(u, "OK") Then Return False
            If u.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
                OrElse u.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
                OrElse u.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) Then Return True
            If u.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Return False
        Next
        Return False
    End Function

    ''' <summary>B-B1/B-B5: splits a set of +CMGS: lines (in arrival
    ''' order) into the counted line and the parked lines. counted
    ''' is True when added &gt; 0 and the COUNTED line (the current
    ''' segment's - see below) has its OK; parkedWithoutOk is True
    ''' when any PARKED line lacks its OK. The caller parks
    ''' flags.Count - added lines.
    ''' FIX B-M1 (Round 18): the counted line is the one with the
    ''' GREATEST message reference (refs, from CmgsLineReferences on
    ''' the same text as the flags), not the last line and not the
    ''' last confirmed line. Flags alone cannot identify the current
    ''' line: the symmetric cases [True, False] (stale confirmed,
    ''' current unconfirmed) and [False, True] (stale unconfirmed,
    ''' current confirmed) are indistinguishable by flags, and the
    ''' Round 17 last-CONFIRMED rule miscounts the first - a stale
    ''' "+CMGS: 12" + OK arriving ahead of the current "+CMGS: 13"
    ''' (grace loop scooping "OK" + "+CMGS: 13" in one poll, or a
    ''' stale 12+OK sitting in the drain buffer) counts stale 12,
    ''' parks current 13 as pending-without-OK, and credits the
    ''' previous row with the current segment's confirmation. The
    ''' greatest reference (13 &gt; 12) counts the current line in
    ''' both symmetric cases and in the mixed ones ([T,T,F],
    ''' [T,F,T], [F,T,F]). Fallback to the Round 17 last-confirmed
    ''' rule (last line when nothing is confirmed, so a lone
    ''' unconfirmed current line is still counted and flagged
    ''' without-OK) when no greatest reference can be picked: refs
    ''' missing or shorter than flags, no parseable mr at all, or a
    ''' tie on the greatest mr. A window spanning the 255 -&gt; 0 wrap
    ''' is unwrapped relative to the first parseable ref before the
    ''' greatest rule is applied (FIX B-L2, Round 24), so the
    ''' post-wrap line is the greatest unwrapped value and is
    ''' counted. Reordering /
    ''' splitting the kept block per line cannot fix this either
    ''' (Round 15 prescription, still rejected - proven to create
    ''' the bug in the 11F,12T case). Parked lines are every line
    ''' except the counted one (all lines when added = 0); every
    ''' caller bounds added to 0 or 1, so that is exactly
    ''' flags.Count - added lines.</summary>
    Private Shared Sub SplitCmgsFlags(flags As List(Of Boolean), added As Integer, ByRef countedConfirmed As Boolean, ByRef parkedWithoutOk As Boolean, ByRef parkedWithoutOkCount As Integer, Optional refs As List(Of Integer) = Nothing)
        countedConfirmed = True
        parkedWithoutOk = False
        ' FIX B-L2: exact number of PARKED lines lacking their OK.
        parkedWithoutOkCount = 0
        If flags.Count = 0 Then Return
        ' FIX B-M1 (Round 18): the counted line is the one with the
        ' GREATEST message reference (see the summary above). Only
        ' when no greatest reference can be picked - refs missing or
        ' shorter than flags, no parseable mr at all, a tie on the
        ' greatest mr - fall back to the last-confirmed rule below.
        ' A window spanning the 255 -&gt; 0 wrap is unwrapped first
        ' (FIX B-L2, Round 24), so it no longer falls back.
        ' FIX B-M1 (Round 22): the greatest-mr pick is factored into
        ' CountedCmgsIndex (moved verbatim) so the grace checks use
        ' the exact same counted line as the counting path.
        Dim countedIndex As Integer = CountedCmgsIndex(flags, refs)
        If countedIndex < 0 Then
            ' Fallback (Round 17 last-confirmed rule): the counted line
            ' is the last CONFIRMED one; when no line at all is
            ' confirmed the old rule (last line) is kept, so a lone
            ' unconfirmed current line is still counted and flagged
            ' without-OK instead of parked.
            For k As Integer = flags.Count - 1 To 0 Step -1
                If flags(k) Then
                    countedIndex = k
                    Exit For
                End If
            Next
            If countedIndex < 0 Then
                countedIndex = flags.Count - 1
                countedConfirmed = False
            Else
                countedConfirmed = True
            End If
        Else
            countedConfirmed = flags(countedIndex)
        End If
        ' Parked lines are every line except the counted one (all
        ' lines when added = 0, i.e. nothing was counted).
        For k As Integer = 0 To flags.Count - 1
            If added > 0 AndAlso k = countedIndex Then Continue For
            If Not flags(k) Then
                parkedWithoutOk = True
                parkedWithoutOkCount += 1
            End If
        Next
    End Sub

    ''' <summary>Counts the very-late +CMGS: lines in a drain result
    ''' (drained text + AT probe reply). drainedOk / probeOk are True
    ''' when no counted +CMGS: line lacks its own trailing OK (True
    ''' also when there is none at all), so the caller can tell
    ''' whether ANY counted +CMGS: lacks its own trailing OK.
    ''' FIX B-L2: the count and the confirmation are evaluated on the
    ''' JOINED drained+probe text (JoinDrainAndProbe), not per portion
    ''' — a drained +CMGS: whose OK arrives in the probe window IS
    ''' confirmed by that OK, so it must not be flagged unconfirmed.
    ''' The old per-portion rule ("the probe's own AT-OK can never
    ''' confirm a +CMGS: from the drain") is dropped deliberately.</summary>
    Private Shared Function DrainHasCmgs(drainedText As String, probeText As String, ByRef drainedOk As Boolean, ByRef probeOk As Boolean, ByRef withoutOkCount As Integer) As Integer
        Dim flags As List(Of Boolean) = CmgsLineConfirmations(JoinDrainAndProbe(drainedText, probeText))
        Dim confirmed As Integer = 0
        For Each f As Boolean In flags
            If f Then confirmed += 1
        Next
        drainedOk = (confirmed = flags.Count)
        probeOk = drainedOk
        ' FIX B-L2: exact number of counted lines lacking their OK.
        withoutOkCount = flags.Count - confirmed
        Return flags.Count
    End Function

    ''' <summary>FIX B-L2: joins a drain/probe pair into one text for
    ''' +CMGS: evaluation, so an OK in the probe window can confirm a
    ''' +CMGS: from the drain. A line break is inserted when the drain
    ''' does not end with one, so a drained "+CMGS: 12" and a probe
    ''' "OK" stay two lines instead of gluing into "+CMGS: 12OK".
    ''' (FIX B-L1: the drain's hold-back is generalised to ANY trailing
    ''' unterminated line — DrainResult and the pre-probe hold-back in
    ''' SendEscapeAndDrainCore — so drained always ends at a line
    ''' boundary and the inserted break provably never splits a real
    ''' line in two, e.g. a partial "O" of OK or "ERR" of ERROR. A
    ''' +CMGS: line split across the boundary is rejoined before this
    ''' point: the drain's fragment is preserved in _preservedText and
    ''' the probe's SendCommand seeds it — FIX B-M1.)</summary>
    Private Shared Function JoinDrainAndProbe(drainedText As String, probeText As String) As String
        If String.IsNullOrEmpty(drainedText) Then Return probeText
        If String.IsNullOrEmpty(probeText) Then Return drainedText
        If drainedText.EndsWith(vbCr) OrElse drainedText.EndsWith(vbLf) Then Return drainedText & probeText
        Return drainedText & vbCrLf & probeText
    End Function

    ''' <summary>Bounds a counted +CMGS: total so SegmentsSent can
    ''' never pass SegmentsTotal. Returns 0 when there is nothing to
    ''' add (no total yet, or every segment already counted).</summary>
    Private Shared Function BoundedCmgsAdd(result As SmsResult, count As Integer) As Integer
        If count <= 0 Then Return 0
        If result.SegmentsTotal <= 0 Then Return 0
        Dim remaining As Integer = result.SegmentsTotal - result.SegmentsSent
        If remaining <= 0 Then Return 0
        Return Math.Min(count, remaining)
    End Function

    ''' <summary>Counts +CMGS: confirmations found in a drain/probe
    ''' pair, but ONLY when a Ctrl-Z was actually sent for the current
    ''' segment: a +CMGS: found before any Ctrl-Z is stale (a previous
    ''' segment or a previous send) and must never be counted FOR the
    ''' current segment. D1: it is NOT discarded — it is parked in
    ''' the pending late-confirmation store (see
    ''' RecordPendingLateCmgs) for FrmMain to credit to the previous
    ''' send. The added count is bounded (see BoundedCmgsAdd) and
    ''' capped at 1 (B-B1: only the line with the GREATEST message
    ''' reference is the current segment's - greatest +CMGS: mr,
    ''' falling back to the last confirmed line when refs are
    ''' missing or shorter than flags, no parseable mr at all, a
    ''' tie on the greatest mr, or a window spanning the 255-&gt;0 wrap);
    ''' when that counted line lacks its own trailing OK,
    ''' sawCmgsWithoutOk is set, and the parked lines' status goes
    ''' to the pending store (B-B5). Returns the number of +CMGS: lines
    ''' found (0 when the guard refuses the count).</summary>
    Private Function CountGuardedCmgs(drainedText As String, probeText As String, ctrlZSent As Boolean, result As SmsResult, ByRef sawCmgsWithoutOk As Boolean, ByRef errorTrailedRefused As Boolean) As Integer
        errorTrailedRefused = False
        If Not ctrlZSent Then
            RecordPendingLateCmgs(drainedText, probeText)
            Return 0
        End If
        ' B-B2: a sighting here is real even though nothing may be
        ' added for the current segment — park it in the D1 store
        ' instead of dropping it (the old early return lost it).
        If result.SegmentsTotal <= 0 OrElse result.SegmentsSent >= result.SegmentsTotal Then
            RecordPendingLateCmgs(drainedText, probeText)
            Return 0
        End If
        ' FIX B-L2: evaluate confirmations on the JOINED drained+probe
        ' text (see JoinDrainAndProbe), not per portion — a drained
        ' +CMGS: whose OK arrives in the probe window is confirmed,
        ' not flagged without OK.
        Dim joinedText As String = JoinDrainAndProbe(drainedText, probeText)
        Dim flags As List(Of Boolean) = CmgsLineConfirmations(joinedText)
        Dim found As Integer = flags.Count
        If found <= 0 Then Return 0
        ' (Round 45) B-H2: refuse to count a +CMGS: line whose
        ' trailing result is an ERROR - one submission yields
        ' either +CMGS:+OK or +CMS ERROR, never both, so a
        ' buffer holding both holds two submissions' output:
        ' the +CMGS: is a previous segment's late confirmation
        ' and the ERROR is this segment's failure. Park all
        ' lines found (the error-trailed line included) and
        ' return 0 - the caller reports the segment failed,
        ' never sent.
        If CurrentCmgsHasTrailingError(joinedText) Then
            Dim errCountedConfirmed As Boolean = True
            Dim errParkedWithoutOk As Boolean = False
            Dim errParkedWithoutOkCount As Integer = 0
            SplitCmgsFlags(flags, 0, errCountedConfirmed, errParkedWithoutOk, errParkedWithoutOkCount, CmgsLineReferences(joinedText))
            ParkPendingLateCmgs(found, errParkedWithoutOkCount)
            errorTrailedRefused = True
            Return 0
        End If
        ' B-B1: one AT+CMGS + one Ctrl-Z yields exactly one +CMGS:,
        ' so at most ONE line in this window belongs to the current
        ' segment — the line with the GREATEST message reference;
        ' earlier lines are previous segments/rows and are parked,
        ' not counted.
        Dim added As Integer = Math.Min(BoundedCmgsAdd(result, found), 1)
        result.SegmentsSent += added
        Dim countedConfirmed As Boolean = True
        Dim parkedWithoutOk As Boolean = False
        Dim parkedWithoutOkCount As Integer = 0
        SplitCmgsFlags(flags, added, countedConfirmed, parkedWithoutOk, parkedWithoutOkCount, CmgsLineReferences(joinedText))
        ' B-B5: the current segment's OK-status comes from the
        ' COUNTED line (greatest mr; last-confirmed fallback) only;
        ' the parked lines' status goes to the pending store's
        ' withoutOk flag.
        If added > 0 AndAlso Not countedConfirmed Then sawCmgsWithoutOk = True
        ' B-B2: park the excess (found - added) instead of dropping
        ' it, contradicting the old "counted but discarded" path.
        ParkPendingLateCmgs(found - added, parkedWithoutOkCount)
        Return found
    End Function

    ''' <summary>True when a reply holds a final-result line: OK, ERROR,
    ''' +CMS ERROR, +CME ERROR, +CMGS: (OK optional), NO CARRIER,
    ''' NO DIALTONE, BUSY or NO ANSWER. Echo lines (starting with AT)
    ''' and other unsolicited "+..." lines never count.</summary>
    ''' <summary>FIX B-D3: the single-line half of HasFinalResultLine
    ''' (same rule, factored out so JoinedHasFinalResultLine below can
    ''' apply it to a straddling line's live suffix too). Behaviour
    ''' unchanged: echo lines (starting with AT) never count.</summary>
    Private Shared Function IsFinalResultLineText(t As String) As Boolean
        If t.Length = 0 Then Return False
        If t.StartsWith("AT", StringComparison.OrdinalIgnoreCase) Then Return False
        Return t.Equals("OK", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("+CMS ERROR", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("+CME ERROR", StringComparison.OrdinalIgnoreCase) _
            OrElse t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) _
            OrElse t.Equals("NO CARRIER", StringComparison.OrdinalIgnoreCase) _
            OrElse t.Equals("NO DIALTONE", StringComparison.OrdinalIgnoreCase) _
            OrElse t.Equals("BUSY", StringComparison.OrdinalIgnoreCase) _
            OrElse t.Equals("NO ANSWER", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function HasFinalResultLine(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each line As String In SplitLines(text)
            Dim t As String = line.Trim()
            If IsFinalResultLineText(t) Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 40) B-L2: True when a reply holds a
    ''' final-result line OTHER than a bare +CMGS: line (OK, ERROR,
    ''' +CMS/+CME ERROR, NO CARRIER, ...). A reply whose only
    ''' final-result evidence is a +CMGS: line is a late
    ''' confirmation of a previous segment, not proof that this
    ''' command completed in command mode - it must never set
    ''' _commandModeKnown on its own.</summary>
    Private Shared Function HasFinalResultLineOtherThanCmgs(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each line As String In SplitLines(text)
            Dim t As String = line.Trim()
            If t.Length = 0 Then Continue For
            If t.StartsWith("+CMGS:", System.StringComparison.OrdinalIgnoreCase) Then Continue For
            If IsFinalResultLineText(t) Then Return True
        Next
        Return False
    End Function

    ''' <summary>(Round 40) B-L2: True when this SendCommand reply
    ''' proves the modem is in command mode. The loop must have
    ''' exited on a match of expected or an error line - never on a
    ''' timeout (a timed-out reply proves nothing about the modem's
    ''' mode) and never on a stop; a ">" prompt reply never proves
    ''' it either ('>' is the text prompt, not command mode); and
    ''' the reply must hold final-result evidence beyond a bare
    ''' +CMGS: line alone (a late confirmation of a previous
    ''' segment). Traces: prompt "+CMGS: 12" (timeout) -> False;
    ''' prompt "+CMGS: 12 + OK" (timeout) -> False; prompt with '>'
    ''' seen -> False; AT "OK" -> True; CPIN "+CPIN: READY + OK"
    ''' -> True; ERROR reply -> True.</summary>
    Private Shared Function ReplyProvesCommandMode(expected As String, reply As String, exitedOnExpected As Boolean, exitedOnError As Boolean, exitedOnDefinitiveOk As Boolean) As Boolean
        ' (Round 41) D-Low-2: the definitive-OK exit (B-L3) proves the
        ' command completed too - for READY/CPIN the definitive OK is
        ' the command's own final result, so the flag counts exactly
        ' like exitedOnExpected/exitedOnError. The expected = ">"
        ' exclusion below still applies: on the prompt path a
        ' definitive OK must never prove command mode, since '>' may
        ' still arrive after it (the B-L2 hazard).
        If Not (exitedOnExpected OrElse exitedOnError OrElse exitedOnDefinitiveOk) Then Return False
        If String.Equals(expected, ">", System.StringComparison.Ordinal) Then Return False
        Return HasFinalResultLineOtherThanCmgs(reply)
    End Function

    ''' <summary>FIX B-D3: HasFinalResultLine on a buffer whose first
    ''' line may be a held fragment glued onto a live reply (held
    ''' "+CM" + live "OK" + vbCrLf -> "+CMOK"). boundary is the held
    ''' fragment's length (the fragment is the buffer's prefix). A
    ''' line fully in the live region is checked as usual; a
    ''' straddling line's LIVE suffix is checked too, so a live OK
    ''' glued onto the fragment still counts as a final result.
    ''' FIX B-L1: the JOINED line itself counts ONLY as a +CMGS:
    ''' completion (the one line whose completion is a live
    ''' confirmation); OK / ERROR / BUSY / etc. count only via the
    ''' live suffix or a fully-live line, exactly as
    ''' JoinedMatchesExpected / JoinedHasErrorLine do — held "O" +
    ''' live "K" must never end the wait.</summary>
    Private Shared Function JoinedHasFinalResultLine(text As String, boundary As Integer) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        Dim pos As Integer = 0
        While pos < text.Length
            Dim lineStart As Integer = pos
            Dim lineEnd As Integer = pos
            While lineEnd < text.Length AndAlso text(lineEnd) <> ChrW(13) AndAlso text(lineEnd) <> ChrW(10)
                lineEnd += 1
            End While
            Dim line As String = text.Substring(lineStart, lineEnd - lineStart)
            Dim nextPos As Integer = lineEnd
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(13) Then nextPos += 1
            If nextPos < text.Length AndAlso text(nextPos) = ChrW(10) Then nextPos += 1
            Dim t As String = line.Trim()
            If t.Length > 0 Then
                If lineStart >= boundary Then
                    ' Fully-live line: checked as usual.
                    If IsFinalResultLineText(t) Then Return True
                ElseIf lineEnd > boundary Then
                    ' FIX B-L1: straddling line — the JOINED line counts
                    ' ONLY as a +CMGS: completion (the one line whose
                    ' completion is itself a live confirmation). OK /
                    ' ERROR / BUSY / etc. count only via the live suffix
                    ' (or a fully-live line), exactly as
                    ' JoinedMatchesExpected / JoinedHasErrorLine do:
                    ' held "O" + live "K" must never end the wait.
                    If t.StartsWith("+CMGS:", StringComparison.OrdinalIgnoreCase) Then Return True
                    Dim liveSuffix As String = text.Substring(boundary, lineEnd - boundary).Trim()
                    If liveSuffix.Length > 0 AndAlso IsFinalResultLineText(liveSuffix) Then Return True
                End If
                ' Fully-stale line (lineEnd <= boundary): never counts.
            End If
            pos = nextPos
        End While
        Return False
    End Function

    ''' <summary>GSM 03.38 default alphabet: character -> septet value.</summary>
    Private Shared Function GsmDefaultMap() As Dictionary(Of Char, Integer)
        Return _gsmDefaultMapLazy.Value
    End Function

    Private Shared Function BuildGsmDefaultMap() As Dictionary(Of Char, Integer)
                    Dim built As New Dictionary(Of Char, Integer)()
            built.Add("@"c, &H0)
            built.Add(ChrW(&HA3), &H1)   ' £
            built.Add("$"c, &H2)
            built.Add(ChrW(&HA5), &H3)   ' ¥
            built.Add(ChrW(&HE8), &H4)   ' è
            built.Add(ChrW(&HE9), &H5)   ' é
            built.Add(ChrW(&HF9), &H6)   ' ù
            built.Add(ChrW(&HEC), &H7)   ' ì
            built.Add(ChrW(&HF2), &H8)   ' ò
            built.Add(ChrW(&HC7), &H9)   ' Ç
            built.Add(ChrW(&HA), &HA)    ' LF
            built.Add(ChrW(&HD8), &HB)   ' Ø
            built.Add(ChrW(&HF8), &HC)   ' ø
            built.Add(ChrW(&HD), &HD)    ' CR
            built.Add(ChrW(&HC5), &HE)   ' Å
            built.Add(ChrW(&HE5), &HF)   ' å
            built.Add(ChrW(&H394), &H10) ' Δ
            built.Add("_"c, &H11)
            built.Add(ChrW(&H3A6), &H12) ' Φ
            built.Add(ChrW(&H393), &H13) ' Γ
            built.Add(ChrW(&H39B), &H14) ' Λ
            built.Add(ChrW(&H3A9), &H15) ' Ω
            built.Add(ChrW(&H3A0), &H16) ' Π
            built.Add(ChrW(&H3A8), &H17) ' Ψ
            built.Add(ChrW(&H3A3), &H18) ' Σ
            built.Add(ChrW(&H398), &H19) ' Θ
            built.Add(ChrW(&HC6), &H1C)  ' Æ
            built.Add(ChrW(&HE6), &H1D)  ' æ
            built.Add(ChrW(&HDF), &H1E)  ' ß
            built.Add(ChrW(&HC9), &H1F)  ' É
            built.Add(" "c, &H20)
            Dim code As Integer
            For code = AscW("!"c) To AscW("?"c)
                If code = AscW("$"c) Then Continue For
                built.Add(ChrW(code), code)
            Next
            built.Add(ChrW(&HA4), &H24)  ' ¤
            built.Add(ChrW(&HA1), &H40)  ' ¡
            For code = AscW("A"c) To AscW("Z"c)
                built.Add(ChrW(code), code)
            Next
            built.Add(ChrW(&HC4), &H5B)  ' Ä
            built.Add(ChrW(&HD6), &H5C)  ' Ö
            built.Add(ChrW(&HD1), &H5D)  ' Ñ
            built.Add(ChrW(&HDC), &H5E)  ' Ü
            built.Add(ChrW(&HA7), &H5F)  ' §
            built.Add(ChrW(&HBF), &H60)  ' ¿
            For code = AscW("a"c) To AscW("z"c)
                built.Add(ChrW(code), code)
            Next
            built.Add(ChrW(&HE4), &H7B)  ' ä
            built.Add(ChrW(&HF6), &H7C)  ' ö
            built.Add(ChrW(&HF1), &H7D)  ' ñ
            built.Add(ChrW(&HFC), &H7E)  ' ü
            built.Add(ChrW(&HE0), &H7F)  ' à
        Return built
    End Function

    ''' <summary>GSM 03.38 extension alphabet (sent after a 0x1B escape):
    ''' character -> septet value.</summary>
    Private Shared Function GsmExtensionMap() As Dictionary(Of Char, Integer)
        Return _gsmExtMapLazy.Value
    End Function

    Private Shared Function BuildGsmExtensionMap() As Dictionary(Of Char, Integer)
                    Dim builtExt As New Dictionary(Of Char, Integer)()
            builtExt.Add(ChrW(&HC), &HA)    ' form feed
            builtExt.Add("^"c, &H14)
            builtExt.Add("{"c, &H28)
            builtExt.Add("}"c, &H29)
            builtExt.Add("\"c, &H2F)
            builtExt.Add("["c, &H3C)
            builtExt.Add("~"c, &H3D)
            builtExt.Add("]"c, &H3E)
            builtExt.Add("|"c, &H40)
            builtExt.Add(ChrW(&H20AC), &H65) ' €
        Return builtExt
    End Function

    ''' <summary>Encodes a GSM segment as raw septet bytes (GSM 03.38
    ''' default table). Returns Nothing when a character cannot be sent
    ''' in GSM text mode: extension-table characters need a 0x1B escape
    ''' and 0x1A/0x1B are control bytes (0x1B aborts text mode, 0x1A ends
    ''' the message), so neither may ever be emitted as data. A GSM
    ''' segment is NEVER sent via Write(String): the port's ASCII default
    ''' encoding would corrupt accented Latin (é/ü/ñ/£) to '?'.
    ''' On failure the offending character and its (0-based) index are
    ''' returned in unsupportedChar / unsupportedIndex. Unmapped
    ''' characters also fail (Nothing) instead of silently becoming '?'.</summary>
    ''' <summary>Next concatenation reference for a multi-segment
    ''' message: an incrementing byte, 1..255, wrapping to 1 (never
    ''' 0). One value per message — the caller passes the same
    ''' reference to every part, so the phone joins them.
    ''' (Round 57) Shared across instances: see _concatRef.</summary>
    Public Function NextConcatRef() As Integer
        SyncLock _concatRefLock
            _concatRef += 1
            If _concatRef > 255 Then _concatRef = 1
            Return _concatRef
        End SyncLock
    End Function

    ''' <summary>
    ''' (Round 56) Builds a CONCATENATED SMS-SUBMIT PDU (hex) for a
    ''' GSM 7-bit part in PDU mode (AT+CMGF=0): the text-mode path
    ''' cannot carry a UDH, so a long Latin message (>160 chars) is
    ''' sent as PDU parts, each with the 6-octet concatenation
    ''' header 05 00 03 &lt;ref&gt; &lt;total&gt; &lt;seq&gt; (8-bit reference),
    ''' first octet 41 (UDHI set), DCS 00 (GSM 7-bit). The UDH
    ''' occupies 6 octets = 48 bits, followed by ONE fill bit
    ''' (bit 48 = 0) so the text starts on a septet boundary
    ''' at bit 49 (3GPP TS 23.038): UDL (in septets) is 7 + the
    ''' message septets, packed LSB-first from bit 49.
    ''' (Round 57) The Round-56 code packed from bit 48, so a
    ''' real phone decoded every long Latin message garbled
    ''' ("Hi" arrived as "d4"). Only default-alphabet
    ''' characters are accepted (an extension character is routed
    ''' to UCS2 by NeedsUcs2 before this runs); any other
    ''' character fails loudly with String.Empty. Returns
    ''' String.Empty when the number is not a valid +20 number
    ''' (see PhoneNormalizer.BuildDestinationSemiOctets), the
    ''' part is empty, or a UDH parameter is out of range.
    ''' </summary>
    Public Shared Function BuildGsmSubmitPdu(normalizedNumber As String, segment As String, udhRef As Integer, udhTotal As Integer, udhSeq As Integer) As String
        If String.IsNullOrEmpty(segment) OrElse String.IsNullOrEmpty(normalizedNumber) Then Return String.Empty
        If udhRef < 1 OrElse udhRef > 255 OrElse udhTotal < 2 OrElse udhTotal > 255 OrElse udhSeq < 1 OrElse udhSeq > udhTotal Then Return String.Empty
        Dim daField As String = PhoneNormalizer.BuildDestinationSemiOctets(normalizedNumber)
        If daField = String.Empty Then Return String.Empty
        Dim defaultMap As Dictionary(Of Char, Integer) = GsmDefaultMap()
        Dim septets As New List(Of Integer)()
        For Each ch As Char In segment
            Dim septet As Integer = 0
            If Not defaultMap.TryGetValue(ch, septet) Then Return String.Empty
            septets.Add(septet)
        Next
        If septets.Count = 0 Then Return String.Empty
        Dim udlSeptets As Integer = 7 + septets.Count
        If udlSeptets > 255 Then Return String.Empty
        Dim udh As Byte() = {CByte(&H5), CByte(&H0), CByte(&H3), CByte(udhRef), CByte(udhTotal), CByte(udhSeq)}
        ' (Round 57) D-55-H1: the 6-octet UDH is 48 bits; GSM 7-bit
        ' text must start on a septet boundary, so ONE fill bit
        ' (bit 48 = 0 — the ud array is zero-initialised) precedes
        ' the first message septet at bit 49 (3GPP TS 23.038).
        ' The Round-56 code packed from bit 48: UDL was right
        ' (7 + N) but every long Latin message decoded garbled on
        ' a real phone ("Hi" arrived as "d4"). totalBits counts
        ' from bit 49 too: for N = 0 mod 8 the UD grows by one
        ' octet (e.g. an 8-char part: 14 octets, not 13).
        Dim totalBits As Integer = 49 + 7 * septets.Count
        Dim udOctetCount As Integer = (totalBits + 7) \ 8
        Dim ud As Byte() = New Byte(udOctetCount - 1) {}
        Array.Copy(udh, ud, 6)
        Dim bitPos As Integer = 49
        For Each septet As Integer In septets
            Dim byteIndex As Integer = bitPos \ 8
            Dim bitOffset As Integer = bitPos Mod 8
            ud(byteIndex) = CByte(ud(byteIndex) Or ((septet << bitOffset) And &HFF))
            If bitOffset + 7 > 8 Then
                ud(byteIndex + 1) = CByte(ud(byteIndex + 1) Or (septet >> (8 - bitOffset)))
            End If
            bitPos += 7
        Next
        Dim sb As New System.Text.StringBuilder()
        sb.Append("00") ' SMSC length 0: use the SIM's default SMSC.
        sb.Append("41") ' SMS-SUBMIT with UDHI (concatenated part).
        sb.Append("00") ' Message reference (the modem assigns one).
        sb.Append(daField)
        sb.Append("00") ' Protocol identifier.
        sb.Append("00") ' Data coding scheme: GSM 7-bit.
        sb.Append(udlSeptets.ToString("X2")) ' User data length, in septets.
        For Each b As Byte In ud
            sb.Append(b.ToString("X2"))
        Next
        Return sb.ToString()
    End Function

    Private Shared Function GsmEncodeSegment(segment As String, ByRef unsupportedChar As Char, ByRef unsupportedIndex As Integer) As Byte()
        unsupportedChar = ChrW(0)
        unsupportedIndex = -1
        Dim bytes As New List(Of Byte)()
        Dim defaultMap As Dictionary(Of Char, Integer) = GsmDefaultMap()
        Dim extMap As Dictionary(Of Char, Integer) = GsmExtensionMap()
        For i As Integer = 0 To segment.Length - 1
            Dim ch As Char = segment(i)
            Dim septet As Integer = 0
            If defaultMap.TryGetValue(ch, septet) Then
                ' NOTE: no character in the GSM default map has
                ' septet &H1A (Ctrl-Z) or &H1B (ESC), so there is
                ' nothing to guard here; the real protection is
                ' PhoneNormalizer.NeedsUcs2, which routes any
                ' character outside GSM (including every character
                ' that would need the 0x1B escape) to UCS2 before
                ' this encoder ever runs.
                bytes.Add(CByte(septet))
            ElseIf extMap.TryGetValue(ch, septet) Then
                ' Extension characters need a 0x1B escape byte, which
                ' is the text-mode ABORT byte: they cannot be sent in
                ' GSM text mode at all (routed to UCS2 instead).
                unsupportedChar = ch
                unsupportedIndex = i
                Return Nothing
            Else
                ' Unmapped character: fail loudly instead of silently
                ' sending '?' in place of the real character.
                unsupportedChar = ch
                unsupportedIndex = i
                Return Nothing
            End If
        Next
        Return bytes.ToArray()
    End Function

    ''' <summary>True when the text holds an unpaired surrogate: a
    ''' high surrogate not followed by a low one, or a lone low
    ''' surrogate. Such a character silently becomes U+FFFD in
    ''' ToUcs2Hex (PhoneNormalizer) (Round 51: not '?'), so the
    ''' caller must fail loudly instead of sending it.</summary>
    Private Shared Function HasUnpairedSurrogate(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        Dim i As Integer = 0
        While i < text.Length
            If Char.IsSurrogate(text(i)) Then
                If Not Char.IsSurrogatePair(text, i) Then Return True
                i += 2
            Else
                i += 1
            End If
        End While
        Return False
    End Function

    Private Sub MarkCancelled(result As SmsResult, Optional withoutOk As Boolean = False)
        result.Success = False
        result.Cancelled = True
        Dim withoutOkSuffix As String = If(withoutOk, " (بدون تأكيد OK)", String.Empty)
        ' (Round 32) D-L1: key the counted branch on SegmentsSent > 0, like every render site — a stop with 0 sent must not warn of re-send duplication that cannot happen.
        If result.SegmentsTotal > 0 AndAlso result.SegmentsSent > 0 Then
            result.Message = "تم الإيقاف — تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " قبل الإيقاف" & " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها" & withoutOkSuffix  ' FIX B-R29-L1 (Round 31): stopped-partial re-send warning
        Else
            result.Message = "تم الإيقاف بواسطة المستخدم" & withoutOkSuffix
        End If
    End Sub

    ''' <summary>
    ''' Sends one AT command and reads the reply until the expected text
    ''' appears or the timeout elapses. Returns the full reply.
    ''' Exits immediately on ERROR / +CMS ERROR / +CME ERROR instead of
    ''' burning the whole timeout. A failed write returns an empty reply.
    ''' When stopPredicate is supplied it is checked every poll (~50ms)
    ''' so a stop aborts mid-command. Error checks are line-anchored.
    ''' discardBeforeWrite is retained for call compatibility only
    ''' and no longer discards anything: pending bytes are always
    ''' read into the preserved buffer and prepended to the reply
    ''' (see _preservedText), so a reply arriving between a drain
    ''' and this write (e.g. a late +CMGS:) is never destroyed.
    ''' </summary>
    Public Function SendCommand(command As String, expected As String, timeoutMs As Integer, Optional stopPredicate As Func(Of Boolean) = Nothing, Optional discardBeforeWrite As Boolean = True, Optional parkStaleCmgs As Boolean = True) As String
        If _port Is Nothing OrElse Not _port.IsOpen Then Return String.Empty
        If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Return String.Empty
        ' Read-and-preserve: the input buffer is NEVER discarded
        ' here. Bytes already waiting (e.g. a very-late +CMGS: from
        ' a previous segment) are read into the preserved buffer and
        ' prepended to this reply, so the caller can examine them
        ' instead of losing them to DiscardInBuffer.
        ' discardBeforeWrite is retained for call compatibility
        ' only: it no longer discards — pending bytes are always
        ' preserved, whatever its value.
        Try
            Dim pending As String = _port.ReadExisting()
            If Not String.IsNullOrEmpty(pending) Then _preservedText &= pending
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Preserve read failed: " & ex.Message)
        End Try
        Try
            _port.Write(command & vbCrLf)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Write failed: " & ex.Message)
            Return String.Empty
        End Try
        Dim watch As New System.Diagnostics.Stopwatch()
        watch.Start()
        Dim buffer As New System.Text.StringBuilder()
        ' D5: the preserved prefix was read BEFORE this command was
        ' written, so a stale final-result line in it (OK / ERROR /
        ' +CMS ERROR / +CME ERROR) must never satisfy this command's
        ' match — strip it. A +CMGS: line in it is never dropped:
        ' StripStaleFinalResultLines routes it to the D1 pending
        ' late-confirmation store instead.
        ' FIX B-D1: parkStaleCmgs = False is the in-send recovery
        ' drain's probe (see SendEscapeAndDrainCore) - a rejoined
        ' +CMGS: stays in the reply for CountGuardedCmgs.
        Dim keptStaleCmgs As String = String.Empty
        buffer.Append(StripStaleFinalResultLines(TakePreservedText(), parkStaleCmgs, keptStaleCmgs))
        ' FIX B-M1: the kept stale prefix ends here — record it, then
        ' REJOIN the fragment Strip held back (a trailing unterminated
        ' line, e.g. "+CM") with the bytes arriving during this live
        ' read, BEFORE any +CMGS:/OK attribution: the fragment must not
        ' be stranded across the read. staleLength is the stale/live
        ' boundary — every byte before it arrived before this command
        ' was written. The live loop matches expected/ERROR only on
        ' fully-live lines (JoinedMatchesExpected / JoinedHasErrorLine),
        ' and the +CMGS: pair's own OK never satisfies the command.
        Dim staleKeptLength As Integer = buffer.Length
        Dim heldFragment As String = TakePreservedText()
        buffer.Append(heldFragment)
        Dim staleLength As Integer = buffer.Length
        ' (Round 40) B-L2/B-L3: record HOW the live loop exits - the
        ' tail may only assert command mode on a match of expected or
        ' an error line, never on a timeout (see
        ' ReplyProvesCommandMode below); and a definitive final OK
        ' (B-L3) ends the wait promptly for commands whose expected
        ' token is not "OK".
        Dim exitedOnError As Boolean = False
        Dim exitedOnExpected As Boolean = False
        Dim exitedOnDefinitiveOk As Boolean = False
        ' (Round 43) B-M1/B-L1/B-L2: one coherent orphaned-OK
        ' mechanism (defects 2, 6, 7 — the Round 42 single-flag
        ' swallow split three ways). A parked UNCONFIRMED +CMGS:
        ' line's OK may still be in flight (its line was parked by
        ' this call's Strip — or by an earlier call — without its OK,
        ' so the forward scan in JoinedHasDefinitiveFinalOk cannot
        ' claim the live OK for it). A live OK is therefore presumed
        ' to be that parked line's confirmation and must NOT trigger
        ' the definitive-OK exit — the loop keeps waiting for the
        ' command's own reply (">" on the prompt path, "+CPIN: READY"
        ' on CheckSim) — subject to the three rules below:
        ' - allowance (D2): the store is a COUNT, so up to that many
        '   live definitive OKs may be swallowed, one per parked
        '   unconfirmed line — never just the first.
        ' - guard (D6): an OK is swallowed only when no info line
        '   of THIS command preceded it in this command's live
        '   reply (see LiveOkPrecededByInfoLine — Round 44: the
        '   allowlist is the command's own info prefix only, so a
        '   URC of any kind never counts as info) — the command's
        '   own final OK is preceded by its info line (e.g.
        '   "+CPIN: SIM PIN"), while an orphaned OK arrives bare
        '   or after URCs only.
        ' - ambiguity (D-L2, Round 44): an info-preceded OK while
        '   the allowance lasts is ambiguous (parked line's OK
        '   after our info line, or our own final OK) — it is
        '   swallowed PROVISIONALLY and a 500 ms grace decides:
        '   a second final result confirms it as the parked
        '   line's, the grace expiring proves it was our own.
        '   The provisional is counted only on confirmation, so
        '   the Round 43 (c2) case (PIN-locked SIM, parked line,
        '   own OK only) still leaves the store untouched.
        ' - reconciliation (D7): the store is decremented per
        '   swallowed OK only after the loop, and only when the loop
        '   exited on the command's own reply (expected, definitive
        '   OK, or error) — proof the swallowed OK was the parked
        '   line's. A timeout after a swallow leaves the flag set.
        ' nextOkScanFrom is the re-arm floor: the buffer offset just
        ' past the last swallowed OK, so a swallowed OK can never
        ' re-trigger the exit on a later chunk (e.g. a URC arriving
        ' before ">"). When nothing is parked unconfirmed the
        ' allowance is 0 and the exit fires promptly, as before.
        ' (Round 45) B-H1: the same machinery now also covers
        ' expected=OK (AT / ATE0 / AT+CMGF=1 / AT+CSCS / AT+CSQ):
        ' a live OK with an allowance outstanding is swallowed
        ' provisionally, never exited on - a bare OK for a
        ' command with NO info line is ambiguous (the command's
        ' own OK is bare too), so it is provisional, not a
        ' definitive swallow. A second final result (a second
        ' OK, or an error) confirms it as the parked line's.
        ' (Round 46) B-M1: a grace expiry now exits on the
        ' provisional as the command's own, for expected=OK
        ' too (see the grace-expiry branch below). The two
        ' cases are indistinguishable by bytes alone - one
        ' live OK with a parked line outstanding is either
        ' the command's own answer or the parked line's
        ' confirmation - but they are not equally likely: a
        ' modem alive enough to emit that OK will answer AT,
        ' while the at_falsepos fiction (the modem never
        ' answers AT, only the parked line's OK arrives)
        ' needs bytes already sitting in the UART buffer at
        ' the instant of unplug. The normal case fires in
        ' routine operation (a late +CMGS: in any setup
        ' window, or any AT/CMGF/CSCS run after a disclaimed
        ' park), so the grace resolves the ambiguity in
        ' favour of the command's own OK. KNOWN CONSEQUENCE:
        ' at_falsepos now returns True (~0.8 s) - the b46
        ' harness expectation is updated to say so.
        ' Swallowed OKs are excised from the reply, so
        ' ReplyMatchesExpected cannot spend a parked line's
        ' OK as the command's own.
        ' (Round 45) B-M1: a live +CMGS: line's claim scan stops
        ' at the command's own info line, and a line disclaimed
        ' that way is parked unconfirmed on a definitive-OK exit.
        Dim orphanedOkAllowance As Integer = _pendingLateCmgsWithoutOkCount
        Dim orphanedOkSwallowed As Integer = 0
        Dim nextOkScanFrom As Integer = staleLength
        ' (Round 44) D-L2: a provisionally-swallowed OK — an OK
        ' preceded by the command's own info line while the
        ' allowance lasts is AMBIGUOUS: it may be the parked
        ' line's OK arriving after our info line, or our own
        ' final OK. It is swallowed provisionally (one allowance
        ' unit reserved, line start + grace deadline recorded)
        ' and the loop keeps waiting a short grace (500 ms) for
        ' a second final result: if one arrives, the provisional
        ' OK was the parked line's (it is counted swallowed);
        ' if the grace expires, it was our own (the loop exits
        ' on it as a definitive OK, nothing counted). The grace
        ' is bounded and far below the 5 s timeout, so the
        ' Round-41 B-L1 fast exit is preserved: a PIN-locked SIM
        ' with a parked line costs at most ~0.5 s extra, never
        ' the full timeout. -1 = no provisional OK pending.
        Dim provisionalOkLineStart As Integer = -1
        Dim provisionalOkDeadlineMs As Long = -1
        ' (Round 46) B-M2: a +CMS ERROR skipped on the prompt
        ' path starts a short grace (see below): '>' arriving
        ' means the error was late; grace expiry with no '>'
        ' means it was the command's own. -1 = no grace pending.
        Dim cmsErrorGraceLineStart As Integer = -1
        Dim cmsErrorGraceDeadlineMs As Long = -1
        ' (Round 45) B-H1: line starts (into buffer) of OKs
        ' confirmed as a parked line's confirmation - excised
        ' from the returned reply below, so a swallowed
        ' orphaned OK can never make a caller report success
        ' for a command that failed or never answered.
        ' (Round 45) B-M1: exitOkLineStart is the OK the loop
        ' exited on (definitive-OK exit only), for the
        ' disclaimed-live-+CMGS: parking below.
        Dim swallowedOkLineStarts As New System.Collections.Generic.List(Of Integer)()
        Dim exitOkLineStart As Integer = -1
        While watch.ElapsedMilliseconds < timeoutMs
            Dim chunk As String = _port.ReadExisting()
            If Not String.IsNullOrEmpty(chunk) Then
                buffer.Append(chunk)
                Dim text As String = buffer.ToString()
                ' (Round 45) B-H1: the provisional OK is consumed
                ' for match purposes (blanked in matchText), so a
                ' second final result - not the provisional OK
                ' itself - is what the matches below see.
                Dim matchText As String = text
                If provisionalOkLineStart >= 0 Then matchText = TextWithLineBlanked(text, provisionalOkLineStart)
                If JoinedHasErrorLine(matchText, staleLength, expected, orphanedOkAllowance, provisionalOkLineStart) Then
                    ' (Round 44) D-L2: the command's own reply
                    ' arrived separately, proving the provisional
                    ' OK was the parked line's, not this command's
                    ' (one AT command yields exactly one final
                    ' result code — the ERROR-exit reconciliation
                    ' principle, applied here at once).
                    If provisionalOkLineStart >= 0 Then
                        orphanedOkSwallowed += 1
                        swallowedOkLineStarts.Add(provisionalOkLineStart)
                        ' (Round 46) B-L1: the provisional just
                        ' confirmed may have disclaimed a live
                        ' +CMGS: line before it (its OK was this
                        ' command's own, not the line's
                        ' confirmation) - anchor the post-loop
                        ' parking on it, so the line is parked
                        ' unconfirmed instead of evaporating.
                        exitOkLineStart = provisionalOkLineStart
                        provisionalOkLineStart = -1
                        provisionalOkDeadlineMs = -1
                    End If
                    exitedOnError = True
                    Exit While
                End If
                ' (Round 46) B-M2: on the prompt path only, a
                ' +CMS ERROR skipped above (a parked line is
                ' outstanding, so it may be a previous segment's
                ' late result) starts a short grace: '>' arriving
                ' within ~1 s means the error was late; if no '>'
                ' arrives the error was this command's own (the
                ' modem rejected the AT+CMGS prompt) and the
                ' grace-expiry check below exits on it. Non-prompt
                ' commands never reach here as their own error:
                ' +CMS ERROR can only come from an SMS submission.
                If String.Equals(expected, ">", System.StringComparison.Ordinal) AndAlso cmsErrorGraceLineStart < 0 Then
                    Dim skippedErrStart As Integer = FirstLiveSkippedCmsErrorLineStart(matchText, staleLength, expected, orphanedOkAllowance, provisionalOkLineStart)
                    If skippedErrStart >= 0 Then
                        cmsErrorGraceLineStart = skippedErrStart
                        cmsErrorGraceDeadlineMs = watch.ElapsedMilliseconds + 1000
                    End If
                End If
                Dim matchedViaStaleWindow As Boolean = False
                Dim matchesExpected As Boolean = JoinedMatchesExpected(matchText, staleLength, expected, matchedViaStaleWindow)
                If matchesExpected Then
                    ' (Round 44) D-L2: same proof as the error
                    ' exit above — the expected token arrived
                    ' separately, so the provisional OK was not
                    ' this command's final OK.
                    If matchedViaStaleWindow Then
                        ' (Round 17) B-L1: a live OK after a
                        ' FULLY-STALE +CMGS: line IS the probe's
                        ' own reply — exit on it at once, never
                        ' swallow it as an orphan.
                        If provisionalOkLineStart >= 0 Then
                            orphanedOkSwallowed += 1
                            swallowedOkLineStarts.Add(provisionalOkLineStart)
                            ' (Round 46) B-L1: anchor the post-loop
                            ' parking on the confirmed provisional,
                            ' as on the error exit above.
                            exitOkLineStart = provisionalOkLineStart
                            provisionalOkLineStart = -1
                            provisionalOkDeadlineMs = -1
                        End If
                        exitedOnExpected = True
                        Exit While
                    End If
                    If String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase) AndAlso orphanedOkAllowance > 0 Then
                        ' (Round 45) B-H1: this OK may be a parked
                        ' line's orphan — do NOT exit on it; fall
                        ' through to the classification below, which
                        ' provisionally swallows it (or confirms it).
                    Else
                        If provisionalOkLineStart >= 0 Then
                            orphanedOkSwallowed += 1
                            swallowedOkLineStarts.Add(provisionalOkLineStart)
                            ' (Round 46) B-L1: anchor the post-loop
                            ' parking on the confirmed provisional,
                            ' as on the error exit above.
                            exitOkLineStart = provisionalOkLineStart
                            provisionalOkLineStart = -1
                            provisionalOkDeadlineMs = -1
                        End If
                        exitedOnExpected = True
                        Exit While
                    End If
                End If
                ' (Round 40) B-L3: a fully-live OK line that is a
                ' definitive final answer - its immediately-preceding
                ' non-blank line is not a +CMGS: line, so it is not
                ' that line's own confirmation - ends the wait at
                ' once instead of burning the whole timeout (e.g.
                ' AT+CPIN? answering "+CPIN: SIM PIN" + OK for a
                ' PIN-locked SIM). Never on a +CMGS: line's own
                ' confirmation OK (the prompt path must keep waiting
                ' for '>' - see JoinedHasDefinitiveFinalOk).
                ' (Round 45) B-H1: the block also runs for
                ' expected=OK when the match above fell through
                ' (a live OK with an allowance outstanding) - the
                ' classification below provisionally swallows that
                ' OK instead of exiting on it. With no allowance
                ' outstanding the match exits at once, as before.
                Dim runDefinitive As Boolean = False
                If Not String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase) Then
                    runDefinitive = JoinedHasDefinitiveFinalOk(text, staleLength, command)
                ElseIf matchesExpected AndAlso orphanedOkAllowance > 0 Then
                    runDefinitive = True
                End If
                If runDefinitive Then
                    ' (Round 43) B-M1/B-L1/B-L2: classify each live
                    ' unclaimed OK in order, in THIS chunk's text: a
                    ' bare/URC-preceded OK is swallowed while the
                    ' allowance lasts (one per parked line — a second
                    ' orphaned OK in a later poll is swallowed too,
                    ' and one sharing the chunk with the command's own
                    ' OK is swallowed before that OK exits the loop);
                    ' an OK preceded by the command's own info line,
                    ' or any OK once the allowance is spent, IS the
                    ' command's own final OK and exits the loop.
                    Dim definitiveExit As Boolean = False
                    While True
                        Dim okLineStart As Integer = FirstLiveUnclaimedOkLineStart(text, staleLength, nextOkScanFrom, command)
                        Dim okInfoPreceded As Boolean = LiveOkPrecededByInfoLine(text, staleLength, okLineStart, command)
                        Dim okInfoPrefix As String = CommandInfoLinePrefix(command)
                        If okLineStart < 0 Then Exit While
                        If orphanedOkAllowance > 0 AndAlso Not okInfoPreceded AndAlso okInfoPrefix.Length > 0 Then
                            ' Bare / non-info-preceded OK for a command
                            ' that HAS an info line: unambiguously
                            ' orphaned (the command's own OK is always
                            ' info-preceded). A pending provisional OK is
                            ' confirmed by this second final result
                            ' arriving: it was the parked line's, not the
                            ' command's.
                            If provisionalOkLineStart >= 0 Then
                                orphanedOkSwallowed += 1
                                swallowedOkLineStarts.Add(provisionalOkLineStart)
                                provisionalOkLineStart = -1
                                provisionalOkDeadlineMs = -1
                            End If
                            orphanedOkAllowance -= 1
                            orphanedOkSwallowed += 1
                            swallowedOkLineStarts.Add(okLineStart)
                            nextOkScanFrom = OkLineEndAfter(text, okLineStart)
                            Continue While
                        End If
                        If orphanedOkAllowance > 0 AndAlso (okInfoPreceded OrElse okInfoPrefix.Length = 0) Then
                            ' (Round 44) D-L2: AMBIGUOUS OK — preceded
                            ' by the command's own info line, so it may
                            ' be the parked line's OK arriving after our
                            ' info line, or our own final OK. (Round 45)
                            ' B-H1: the same ambiguity holds for a bare
                            ' OK when the command has NO info line at
                            ' all (AT / ATE0 / AT+CMGF=1 / AT+CSCS) —
                            ' the command's own OK is bare too, so a
                            ' bare OK proves nothing either way.
                            ' Confirm a pending provisional first (this
                            ' OK proves it was the parked line's), then
                            ' swallow this one provisionally and keep
                            ' waiting the grace for a second final result.
                            If provisionalOkLineStart >= 0 Then
                                orphanedOkSwallowed += 1
                                swallowedOkLineStarts.Add(provisionalOkLineStart)
                            End If
                            orphanedOkAllowance -= 1
                            provisionalOkLineStart = okLineStart
                            provisionalOkDeadlineMs = watch.ElapsedMilliseconds + 500
                            nextOkScanFrom = OkLineEndAfter(text, okLineStart)
                            Continue While
                        End If
                        ' No allowance left (or none was ever parked):
                        ' this OK is the command's own final OK. A
                        ' pending provisional is confirmed by its
                        ' arrival — it was the parked line's.
                        If provisionalOkLineStart >= 0 Then
                            orphanedOkSwallowed += 1
                            swallowedOkLineStarts.Add(provisionalOkLineStart)
                            provisionalOkLineStart = -1
                            provisionalOkDeadlineMs = -1
                        End If
                        definitiveExit = True
                        exitOkLineStart = okLineStart
                        Exit While
                    End While
                    If definitiveExit Then
                        ' (Round 41) D-Low-2: record that the loop exited
                        ' on the definitive OK, so ReplyProvesCommandMode
                        ' can credit it - except on the prompt path
                        ' (expected = ">"), where a definitive OK must
                        ' still not prove command mode (see that function).
                        ' (Round 45) B-H1: for expected=OK the
                        ' definitive OK IS the expected token, so the
                        ' exit counts as exitedOnExpected.
                        If String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase) Then
                            exitedOnExpected = True
                        ElseIf Not String.Equals(expected, ">", System.StringComparison.Ordinal) Then
                            exitedOnDefinitiveOk = True
                        End If
                        Exit While
                    End If
                End If
            End If
            ' (Round 44) D-L2: the provisional OK's grace expired
            ' with no second final result — it was the command's
            ' own final OK. Exit on it as a definitive OK, WITHOUT
            ' counting it as swallowed: the parked line is still
            ' unconfirmed, so reconciliation must not decrement
            ' the store for it (Round 43 (c2) behaviour, at most
            ' ~0.5 s slower, never the full 5 s timeout).
            If provisionalOkLineStart >= 0 AndAlso watch.ElapsedMilliseconds >= provisionalOkDeadlineMs Then
                If String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase) Then
                    ' (Round 46) B-M1: for expected=OK the grace
                    ' expiry DOES exit on the provisional as the
                    ' command's own final OK (reversing the Round
                    ' 45 B-H1 carve-out). A single live OK with a
                    ' parked line outstanding is the command's own
                    ' answer in routine operation; the at_falsepos
                    ' alternative (the modem never answered at all)
                    ' is near-impossible on a real modem and is the
                    ' accepted known consequence (see the B-H1
                    ' comment above and the b46 at_falsepos trace).
                    ' The provisional is NOT counted as swallowed:
                    ' the parked line stays unconfirmed, so
                    ' reconciliation must not decrement the store
                    ' for it. It is NOT excised either: it is the
                    ' command's own OK, so the reply keeps it and
                    ' ReplyMatchesExpected sees it.
                    exitOkLineStart = provisionalOkLineStart
                    exitedOnExpected = True
                    Exit While
                Else
                    exitOkLineStart = provisionalOkLineStart
                    If Not String.Equals(expected, ">", System.StringComparison.Ordinal) Then exitedOnDefinitiveOk = True
                    Exit While
                End If
            End If
            ' (Round 46) B-M2: the skipped +CMS ERROR's grace
            ' expired with no '>' arriving - the error was this
            ' command's own (the modem rejected the AT+CMGS
            ' prompt), not a late result: exit on it as an error
            ' exit, so the prompt fails promptly (~1 s) with the
            ' error detail in the reply instead of burning the
            ' full timeout. A pending provisional OK is confirmed
            ' by this second final result, exactly as on the
            ' error exit above (and the disclaimed-line parking
            ' is anchored on it, as there).
            If cmsErrorGraceLineStart >= 0 AndAlso watch.ElapsedMilliseconds >= cmsErrorGraceDeadlineMs Then
                If provisionalOkLineStart >= 0 Then
                    orphanedOkSwallowed += 1
                    swallowedOkLineStarts.Add(provisionalOkLineStart)
                    exitOkLineStart = provisionalOkLineStart
                    provisionalOkLineStart = -1
                    provisionalOkDeadlineMs = -1
                End If
                exitedOnError = True
                Exit While
            End If
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
            Threading.Thread.Sleep(50)
        End While
        watch.Stop()
        ' (Round 45) B-M1: a live +CMGS: line whose OK was
        ' disclaimed by the info-line stop (its OK was the
        ' command's own, not its confirmation) is parked
        ' unconfirmed, not evaporated, when the loop exited
        ' on that definitive OK.
        ' (Round 46) B-L1: the same parking now also runs on
        ' an error/expected exit that just confirmed a
        ' provisional OK (exitOkLineStart is anchored on the
        ' provisional there): the READY case no longer
        ' evaporates the disclaimed line. Exits with no
        ' provisional confirmed leave exitOkLineStart at -1
        ' and park nothing, as before.
        If exitOkLineStart >= 0 Then
            Dim disclaimedCmgs As Integer = CountDisclaimedLiveCmgsBeforeOk(buffer.ToString(), staleLength, exitOkLineStart, command)
            If disclaimedCmgs > 0 Then ParkPendingLateCmgs(disclaimedCmgs, disclaimedCmgs)
        End If
        ' (Round 43) B-L2: post-loop reconciliation — each OK
        ' swallowed above was presumed to be a parked line's
        ' confirmation; when the loop then exited on the command's
        ' own reply (expected, definitive OK, or error), that reply
        ' arrived separately, proving the presumption — so the
        ' store's without-OK count drops by one per swallowed OK.
        ' On a timeout (or a stop) there is no such proof — the
        ' swallowed OK may have been the command's own final OK —
        ' so the count is left exactly as it was.
        If orphanedOkSwallowed > 0 AndAlso (exitedOnExpected OrElse exitedOnError OrElse exitedOnDefinitiveOk) Then
            _pendingLateCmgsWithoutOkCount = System.Math.Max(0, _pendingLateCmgsWithoutOkCount - orphanedOkSwallowed)
        End If
        ' FIX B-M1 (returned reply): the kept stale prefix is not part
        ' of this command's reply, so it can never satisfy a caller-side
        ' check (a stale "+CPIN: READY" or "&gt;" line). A first line
        ' that straddles the write is kept only when it is a +CMGS:
        ' line — any other straddling line is a stale line completed by
        ' live bytes (e.g. "O" + "K" into a stale OK) and is dropped.
        ' A trailing unterminated +CMGS: fragment is held back in
        ' _preservedText for the next read. Callers check the reply with
        ' ReplyMatchesExpected, which skips a +CMGS: line and its own OK.
        Dim reply As String = buffer.ToString()
        If staleKeptLength > 0 Then reply = reply.Substring(staleKeptLength)
        ' (Round 45) B-H1: excise the OK lines swallowed above
        ' as a parked line's confirmation - a swallowed orphaned
        ' OK left in the reply would make ReplyMatchesExpected
        ' report success for a command that failed
        ' (cmgf_fail/cscs_fail). A provisional OK left pending
        ' by a timeout/stop for expected=OK is excised too (the
        ' command never proved its own OK); a provisional that
        ' expired as the command's own (grace expiry, any
        ' expected - Round 46 B-M1 restores the exit for
        ' expected=OK) is never excised. Offsets are
        ' buffer-relative, so they shift by staleKeptLength
        ' here; lines go highest-first.
        Dim exciseStarts As New System.Collections.Generic.List(Of Integer)()
        For Each okStart As Integer In swallowedOkLineStarts
            exciseStarts.Add(okStart - staleKeptLength)
        Next
        If provisionalOkLineStart >= 0 AndAlso String.Equals(expected, "OK", System.StringComparison.OrdinalIgnoreCase) AndAlso Not (exitedOnExpected OrElse exitedOnError OrElse exitedOnDefinitiveOk) Then
            exciseStarts.Add(provisionalOkLineStart - staleKeptLength)
        End If
        If exciseStarts.Count > 0 Then reply = RemoveLinesAt(reply, exciseStarts)
        If heldFragment <> String.Empty Then reply = RejoinGluedFragmentReply(reply, heldFragment)
        ' FIX B-M1 (Round 14): the +CMGS: lines (and confirming OKs)
        ' Strip kept (parkStaleCmgs = False) were inside the stale
        ' prefix, so the cut above just discarded them — re-append
        ' them AFTER the rejoin, so the held fragment's glued line is
        ' already repaired. A kept block whose every +CMGS: line has
        ' its confirming OK inside the block (the complete pair) is
        ' PREPENDED as before: its OK is its own confirmation and the
        ' live OK still satisfies the probe. But an UNCONFIRMED kept
        ' block (a +CMGS: line with no OK of its own) must NOT be
        ' prepended: the live loop already consumed the modem's only
        ' live OK as that line's confirmation, so prepending lets
        ' ReplyMatchesExpected spend the same single OK on the kept
        ' line and report the probe unproven, while CountGuardedCmgs
        ' counts the segment — one OK attributed to two consumers.
        ' An unconfirmed block is therefore APPENDED after the live
        ' reply: the live OK is seen first (the probe is proven) and
        ' CmgsLineConfirmations marks the kept line unconfirmed
        ' ([False]). Before appending, a trailing +CMGS: fragment of
        ' the live reply is held back as before, and a line break is
        ' inserted when the live reply does not end on one, so the
        ' kept block can never glue onto a trailing partial line.
        If keptStaleCmgs <> String.Empty Then
            Dim keptFlags As List(Of Boolean) = CmgsLineConfirmations(keptStaleCmgs)
            Dim keptAllConfirmed As Boolean = keptFlags.Count > 0
            For Each keptFlag As Boolean In keptFlags
                If Not keptFlag Then keptAllConfirmed = False
            Next
            If keptAllConfirmed Then
                reply = keptStaleCmgs & reply
            Else
                Dim liveTrailing As String = String.Empty
                SplitTrailingCmgsFragment(reply, liveTrailing)
                If liveTrailing <> String.Empty Then AppendPreservedFragment(liveTrailing)
                If reply <> String.Empty AndAlso Not reply.EndsWith(vbCr) AndAlso Not reply.EndsWith(vbLf) Then reply &= vbCrLf
                reply &= keptStaleCmgs
            End If
        End If
        Dim trailingFragment As String = String.Empty
        SplitTrailingCmgsFragment(reply, trailingFragment)
        If trailingFragment <> String.Empty Then AppendPreservedFragment(trailingFragment)
        ' (Round 40) B-L2: the flag must NOT be set from a prompt reply
        ' or a timed-out reply - when expected is ">" the modem is
        ' either at the text prompt ('>' seen) or in an unknown state
        ' (timed out), never proven in command mode; and a bare
        ' +CMGS: in the reply is a late confirmation of a previous
        ' segment, not this command's completion. ReplyProvesCommandMode
        ' ties the flag to the command actually completing (loop exited
        ' on a match of expected or an error line) with final-result
        ' evidence beyond a +CMGS: line alone. Without this, a very-late
        ' +CMGS: arriving during the AT+CMGS prompt wait set the flag
        ' back to True after it had deliberately been cleared, so the
        ' recovery drain early-returned with no ESC, no drain and no
        ' probe, and the '>' issued just after the wait sat unconsumed
        ' until the next command was swallowed as SMS body text.
        If ReplyProvesCommandMode(expected, reply, exitedOnExpected, exitedOnError, exitedOnDefinitiveOk) Then _commandModeKnown = True
        Return reply
    End Function

    ''' <summary>Basic check: sends AT and expects OK.</summary>
    Public Function TestConnection(Optional stopPredicate As Func(Of Boolean) = Nothing) As Boolean
        Dim response As String = SendCommand("AT", "OK", 5000, stopPredicate)
        Return ReplyMatchesExpected(response, "OK")
    End Function

    ''' <summary>SIM check: AT+CPIN? expects READY.</summary>
    Public Function CheckSim(Optional stopPredicate As Func(Of Boolean) = Nothing) As String
        Dim response As String = SendCommand("AT+CPIN?", "READY", 5000, stopPredicate)
        If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Return "تم الإيقاف بواسطة المستخدم"
        If String.IsNullOrWhiteSpace(response) Then
            Return "خطأ: انتهت المهلة — لا رد من المودم (AT+CPIN?)"
        End If
        If LineMatchesExpected(response, "READY") Then Return "READY"
        ' (Round 44) B-M1: "+CPIN: NOT READY" is a real reply
        ' (SIMCom/Quectel: SIM absent or not ready) and must
        ' never read as READY (see LineTextMatchesExpected).
        ' (Round 45) D-D2: it is now a DISTINCT state, not the
        ' no-SIM / error state - the SIM may be present but not
        ' ready yet, and labelling it 'no SIM' misdiagnoses it.
        ' The SIM is still not usable, so the pre-send check
        ' must not green-light it, and FrmMain's READY branch
        ' (Ready=True) stays unreachable for it.
        If LineMatchesExpected(response, "NOT READY") Then Return "NOT READY - الشريحة غير جاهزة — انتظر قليلاً وأعد الاختبار"
        If LineMatchesExpected(response, "SIM PIN2") Then Return "SIM PIN2 - أدخل رمز PIN2 للشريحة"
        If LineMatchesExpected(response, "SIM PUK2") Then Return "SIM PUK2 - الشريحة مقفلة (PUK2)"
        If LineMatchesExpected(response, "SIM PIN") Then Return "SIM PIN - أدخل رمز PIN للشريحة"
        If LineMatchesExpected(response, "SIM PUK") Then Return "SIM PUK - الشريحة مقفلة"
        If LineStartsWith(response, "ERROR") _
            OrElse LineStartsWith(response, "+CME ERROR") _
            OrElse LineStartsWith(response, "+CMS ERROR") Then Return "ERROR - لا توجد شريحة أو خطأ"
        Return "غير معروف: " & response.Trim()
    End Function

    ''' <summary>
    ''' Signal strength: AT+CSQ. Returns e.g. "الإشارة: 18/31 (-77 dBm) - جيدة".
    ''' Waits for the full reply (OK) so the +CSQ digits have arrived
    ''' before the response is parsed. A modem error is returned as
    ''' error text instead of being hidden as "غير متوفرة".
    ''' </summary>
    Public Function GetSignal(Optional stopPredicate As Func(Of Boolean) = Nothing) As String
        Dim response As String = SendCommand("AT+CSQ", "OK", 5000, stopPredicate)
        If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Return "تم الإيقاف بواسطة المستخدم"
        If String.IsNullOrWhiteSpace(response) Then
            Return "خطأ: انتهت المهلة — لا رد من المودم (AT+CSQ)"
        End If
        Dim signalText As String = "غير متوفرة"
        Dim m As Match = Regex.Match(response, "\+CSQ:\s*(\d+)\s*,\s*(\d+)")
        If m.Success Then
            Dim rssi As Integer = 0
            Integer.TryParse(m.Groups(1).Value, rssi)
            If rssi = 99 Then
                signalText = "لا توجد إشارة (99)"
            Else
                Dim dbm As Integer = -113 + (2 * rssi)
                Dim quality As String = "ضعيفة"
                If rssi >= 20 Then
                    quality = "ممتازة"
                ElseIf rssi >= 15 Then
                    quality = "جيدة"
                ElseIf rssi >= 10 Then
                    quality = "مقبولة"
                End If
                signalText = rssi.ToString() & "/31 (" & dbm.ToString() & " dBm) - " & quality
            End If
        Else
            ' (Round 47) B-obs csq_lateerr: the error check moved here
            ' (it used to lead the function) - a LATE +CMS ERROR can
            ' lead the reply while the command's own +CSQ: line is
            ' present (the B-H3 skip lets the read continue past the
            ' late error), and the old order misreported the signal
            ' as failed although the digits were in the reply. A
            ' +CSQ: line is this command's own answer, so it wins;
            ' the error check only fires when no +CSQ: line exists
            ' (the modem's own error, or a late error with no
            ' signal at all).
            If LineStartsWith(response, "ERROR") _
                OrElse LineStartsWith(response, "+CME ERROR") _
                OrElse LineStartsWith(response, "+CMS ERROR") Then
                Return "خطأ: " & response.Trim()
            End If
            ' Non-empty reply with no +CSQ: and no ERROR above: the
            ' modem answered with something unexpected - say so
            ' instead of hiding it as "غير متوفرة".
            Return "خطأ: رد غير متوقع: " & response.Trim()
        End If
        Return signalText
    End Function

    ' ------------------------------------------------------------------
    ' Inbox (R75W): read received SMS from the SIM (AT+CMGL=4, PDU mode).
    ' Read-only v1: no delete / mark-read. The PDU decode below is pure
    ' logic (Shared, no SerialPort) so the sim pins it exactly; only
    ' ReadInboxSms touches the port, via the same SendCommand path as
    ' TestConnection/CheckSim/GetSignal.
    ' ------------------------------------------------------------------

    ''' <summary>One received SMS as listed by AT+CMGL=4.</summary>
    Public Class InboxMessage
        Public Property Index As Integer
        Public Property Status As String = String.Empty
        Public Property FromNumber As String = String.Empty
        Public Property TimeText As String = String.Empty
        Public Property Text As String = String.Empty
        ' (Fixer F) Which message storage the row was read from
        ' (SM / ME / MT). Empty when the read ran on the modem
        ' default storage (pre-CPMS fallback).
        Public Property Storage As String = String.Empty
    End Class

    Private Shared Function InboxStatusLabel(stat As Integer) As String
        Select Case stat
            Case 0
                Return "غير مقروءة"
            Case 1
                Return "مقروءة"
            Case 2
                Return "غير مرسلة"
            Case 3
                Return "مرسلة"
            Case Else
                Return "؟"
        End Select
    End Function

    ''' <summary>Hex string to bytes. Returns Nothing when the text is not even-length hex.</summary>
    Private Shared Function HexToBytes(hex As String) As Byte()
        If String.IsNullOrEmpty(hex) OrElse (hex.Length Mod 2) <> 0 Then Return Nothing
        Dim bytes(hex.Length \ 2 - 1) As Byte
        Try
            For i As Integer = 0 To bytes.Length - 1
                Dim hi As Integer = Convert.ToInt32(hex.Substring(i * 2, 1), 16)
                Dim lo As Integer = Convert.ToInt32(hex.Substring(i * 2 + 1, 1), 16)
                bytes(i) = CByte((hi << 4) Or lo)
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("HexToBytes failed: " & ex.Message)
            Return Nothing
        End Try
        Return bytes
    End Function

    ''' <summary>One SCTS semi-octet byte as a BCD value (TS 23.040 9.2.3.11: low nibble = tens, high nibble = units).
    ''' -1 when a nibble is not a digit.</summary>
    Private Shared Function BcdSemiOctet(b As Integer) As Integer
        Dim lo As Integer = b And &HF
        Dim hi As Integer = (b >> 4) And &HF
        If lo > 9 OrElse hi > 9 Then Return -1
        Return lo * 10 + hi
    End Function

    ''' <summary>Swapped-BCD semi-octets to digits. A trailing &HF filler is dropped.</summary>
    Private Shared Function SemiOctetsToDigits(data As Byte(), offset As Integer, digitCount As Integer) As String
        Dim sb As New System.Text.StringBuilder()
        Dim i As Integer = offset
        While sb.Length < digitCount AndAlso i < data.Length
            Dim lo As Integer = data(i) And &HF
            sb.Append(ChrW(AscW("0"c) + lo))
            If sb.Length >= digitCount Then Exit While
            Dim hi As Integer = (data(i) >> 4) And &HF
            If hi <> &HF Then sb.Append(ChrW(AscW("0"c) + hi))
            i += 1
        End While
        Return sb.ToString()
    End Function

    ''' <summary>Unpacks GSM 7-bit septets from packed bytes starting at startBit, for septetCount septets.
    ''' An unknown septet becomes U+FFFD (visible, never a silent '?'); 0x1B consumes the next septet from the
    ''' extension table (an unknown extension septet also becomes U+FFFD).</summary>
    Private Shared Function UnpackGsm7Bit(data As Byte(), startBit As Integer, septetCount As Integer) As String
        Dim revDef As Dictionary(Of Integer, Char) = _gsmDefaultRevLazy.Value
        Dim revExt As Dictionary(Of Integer, Char) = _gsmExtRevLazy.Value
        Dim sb As New System.Text.StringBuilder()
        Dim bitPos As Integer = startBit
        Dim s As Integer = 0
        While s < septetCount
            Dim byteIndex As Integer = bitPos \ 8
            If byteIndex >= data.Length Then Exit While
            Dim bitOffset As Integer = bitPos Mod 8
            Dim septet As Integer = (data(byteIndex) >> bitOffset) And &H7F
            If bitOffset + 7 > 8 AndAlso byteIndex + 1 < data.Length Then
                septet = septet Or ((data(byteIndex + 1) And ((1 << (bitOffset + 7 - 8)) - 1)) << (8 - bitOffset))
            End If
            bitPos += 7
            s += 1
            If septet = &H1B Then
                ' Extension escape: the NEXT septet selects the character.
                If s >= septetCount Then
                    sb.Append(ChrW(&HFFFD))
                    Exit While
                End If
                Dim b2 As Integer = bitPos \ 8
                If b2 >= data.Length Then
                    sb.Append(ChrW(&HFFFD))
                    Exit While
                End If
                Dim o2 As Integer = bitPos Mod 8
                Dim ext As Integer = (data(b2) >> o2) And &H7F
                If o2 + 7 > 8 AndAlso b2 + 1 < data.Length Then
                    ext = ext Or ((data(b2 + 1) And ((1 << (o2 + 7 - 8)) - 1)) << (8 - o2))
                End If
                bitPos += 7
                s += 1
                Dim extCh As Char = ChrW(0)
                If revExt.TryGetValue(ext, extCh) Then sb.Append(extCh) Else sb.Append(ChrW(&HFFFD))
            Else
                Dim ch As Char = ChrW(0)
                If revDef.TryGetValue(septet, ch) Then sb.Append(ch) Else sb.Append(ChrW(&HFFFD))
            End If
        End While
        Return sb.ToString()
    End Function

    ''' <summary>Decodes one SMS-DELIVER PDU (hex, as listed by AT+CMGL=4) into an InboxMessage (without Index/Status,
    ''' which come from the +CMGL: header line). Sender: semi-octet digits, '+' when the type is international (0x91),
    ''' alphanumeric (type 0xD0) via GSM 7-bit. Timestamp: SCTS as "YYYY-MM-DD HH:mm:ss +HH:MM" (SCTS is the SMSC's
    ''' local time; the timezone octet's bit 3 is the sign, per TS 23.040 9.2.3.11). Text: alphabet by DCS group (TS 23.038 4) -
    ''' UCS2 when DCS bits 3-2 are 10 in groups 00xx/01xx, or group 1110; 8-bit data as Latin-1 when 01 in groups 00xx/01xx,
    ''' or bit 2 set in group 1111; GSM 7-bit otherwise (incl. groups 1100/1101 and reserved 1000-1011); a UDH (first-octet
    ''' bit 6) is skipped - octets for UCS2/8-bit, septets (fill bit included) for GSM 7-bit. Returns False (msg left empty)
    ''' on any malformed/truncated PDU instead of guessing.</summary>
    Public Shared Function DecodeDeliverPdu(pduHex As String, ByRef msg As InboxMessage) As Boolean
        msg = New InboxMessage()
        Dim data As Byte() = HexToBytes(If(pduHex, String.Empty).Trim())
        If data Is Nothing OrElse data.Length < 1 Then Return False
        Dim pos As Integer = 0
        Dim smscLen As Integer = data(pos)
        pos = 1 + smscLen
        If pos + 1 >= data.Length Then Return False
        Dim firstOctet As Integer = data(pos)
        Dim udhi As Boolean = (firstOctet And &H40) <> 0
        ' (Round 76W) A2: MTI bits 1-0 must be 00 (SMS-DELIVER). A stored
        ' STATUS-REPORT (10) or SUBMIT (01) PDU must never decode as an inbox row.
        If (firstOctet And 3) <> 0 Then Return False
        pos += 1
        ' Originating address.
        Dim oaLen As Integer = data(pos)
        pos += 1
        If pos >= data.Length Then Return False
        Dim oaType As Integer = data(pos)
        pos += 1
        ' (Round 76W) A1: oaLen is a DIGIT count for a numeric OA but a
        ' SEPTET count for an alphanumeric OA (TS 23.040 9.2.3.8) - the
        ' octet count differs: ceil(septets*7/8), not (digits+1)\2.
        Dim oaIsAlpha As Boolean = (oaType And &H70) = &H50
        Dim oaOctets As Integer = If(oaIsAlpha, (oaLen * 7 + 7) \ 8, (oaLen + 1) \ 2)
        If pos + oaOctets > data.Length Then Return False
        If oaIsAlpha Then
            ' Alphanumeric sender id: oaLen septets packed in oaOctets.
            msg.FromNumber = UnpackGsm7Bit(data, pos * 8, oaLen)
        Else
            ' (Round 76W) A4: every OA digit nibble must be BCD 0-9; a
            ' 0xA nibble is malformed, never a ':' inside the number.
            Dim d As Integer = 0
            While d < oaLen
                Dim nib As Integer = If((d Mod 2) = 0, data(pos + d \ 2) And &HF, (data(pos + d \ 2) >> 4) And &HF)
                If nib > 9 Then Return False
                d += 1
            End While
            Dim digits As String = SemiOctetsToDigits(data, pos, oaLen)
            msg.FromNumber = If(oaType = &H91, "+" & digits, digits)
        End If
        pos += oaOctets
        If pos + 10 > data.Length Then Return False
        pos += 1 ' PID.
        Dim dcs As Integer = data(pos)
        pos += 1
        ' Service-centre timestamp: 7 semi-octet bytes YY MM DD hh mm ss tz.
        Dim yy As Integer = BcdSemiOctet(data(pos))
        Dim mo As Integer = BcdSemiOctet(data(pos + 1))
        Dim dd As Integer = BcdSemiOctet(data(pos + 2))
        Dim hh As Integer = BcdSemiOctet(data(pos + 3))
        Dim mm As Integer = BcdSemiOctet(data(pos + 4))
        Dim ss As Integer = BcdSemiOctet(data(pos + 5))
        Dim tzByte As Integer = data(pos + 6)
        pos += 7
        ' (Round 76W) A4: BcdSemiOctet returns -1 for a non-BCD nibble, so
        ' hh/mm/ss are checked < 0 as well (a 0xF nibble used to slip through).
        If yy < 0 OrElse mo < 1 OrElse mo > 12 OrElse dd < 1 OrElse dd > 31 OrElse hh < 0 OrElse hh > 23 OrElse mm < 0 OrElse mm > 59 OrElse ss < 0 OrElse ss > 59 Then Return False
        ' (TS 23.040 9.2.3.11) Timezone octet: units = high nibble, tens = low-nibble bits 0-2, sign = bit 3 (0x08).
        ' (Round 76W) A4: the units nibble must be BCD (0-9) and the value
        ' within +/-48 quarters; 0xF8 is malformed, never -03:45.
        Dim tzSignNeg As Boolean = (tzByte And &H8) <> 0
        Dim tzUnits As Integer = (tzByte >> 4) And &HF
        If tzUnits > 9 Then Return False
        Dim tzQuarters As Integer = ((tzByte And &H7) * 10) + tzUnits
        If tzSignNeg Then tzQuarters = -tzQuarters
        If Math.Abs(tzQuarters) > 48 Then Return False
        Dim tzSign As String = If(tzQuarters < 0, "-", "+")
        Dim tzAbs As Integer = Math.Abs(tzQuarters)
        msg.TimeText = String.Format("{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00} {6}{7:00}:{8:00}", 2000 + yy, mo, dd, hh, mm, ss, tzSign, tzAbs \ 4, (tzAbs Mod 4) * 15)
        If pos >= data.Length Then Return False
        Dim udl As Integer = data(pos)
        pos += 1
        Dim udStart As Integer = pos
        Dim udLen As Integer = data.Length - udStart
        If udLen < 0 Then Return False
        ' (Round 76W) A3: bits 3-2 select the alphabet only in DCS groups
        ' 00xx/01xx (TS 23.038 4). Group 1100/1101 (message waiting) is
        ' always GSM 7-bit, 1110 always UCS2, reserved 1000-1011 GSM 7-bit,
        ' and group 1111 uses bit 2 (1 = 8-bit data, 0 = GSM 7-bit).
        Dim dcsKind As Integer = dcs And &HC
        If dcs >= &H80 Then
            If (dcs And &HF0) = &HE0 Then
                dcsKind = &H8
            ElseIf (dcs And &HF0) = &HF0 Then
                dcsKind = If((dcs And &H4) <> 0, &H4, 0)
            Else
                dcsKind = 0
            End If
        End If
        If dcsKind = &H8 Then
            ' UCS2: UDL is in octets; skip the UDH in octets.
            If udl > udLen Then Return False
            Dim textStart As Integer = udStart
            Dim textLen As Integer = udl
            If udhi Then
                If udLen < 1 Then Return False
                Dim udhl As Integer = data(udStart)
                textStart = udStart + 1 + udhl
                textLen = udl - 1 - udhl
                If textLen < 0 Then Return False
            End If
            If textStart + textLen > data.Length Then Return False
            ' (Round 76W) A4: an odd UCS2 length is malformed, never silently truncated.
            If (textLen Mod 2) <> 0 Then Return False
            msg.Text = System.Text.Encoding.BigEndianUnicode.GetString(data, textStart, textLen)
        ElseIf dcsKind = &H4 Then
            ' 8-bit data: UDL is in octets; skip the UDH in octets; Latin-1 maps bytes 1:1.
            If udl > udLen Then Return False
            Dim textStart As Integer = udStart
            Dim textLen As Integer = udl
            If udhi Then
                If udLen < 1 Then Return False
                Dim udhl As Integer = data(udStart)
                textStart = udStart + 1 + udhl
                textLen = udl - 1 - udhl
                If textLen < 0 Then Return False
            End If
            If textStart + textLen > data.Length Then Return False
            msg.Text = System.Text.Encoding.GetEncoding("ISO-8859-1").GetString(data, textStart, textLen)
        Else
            ' GSM 7-bit: UDL is in septets; the UDH occupies ((UDHL+1)*8+6)\7 septets (fill bit included).
            ' (Round 76W) A4: UDL septets must fit in the bytes actually present.
            If udl > (udLen * 8) \ 7 Then Return False
            Dim skipSeptets As Integer = 0
            If udhi Then
                If udLen < 1 Then Return False
                skipSeptets = ((data(udStart) + 1) * 8 + 6) \ 7
            End If
            Dim textSeptets As Integer = udl - skipSeptets
            If textSeptets < 0 Then Return False
            msg.Text = UnpackGsm7Bit(data, udStart * 8 + skipSeptets * 7, textSeptets)
        End If
        Return True
    End Function

    ''' <summary>Parses a raw AT+CMGL=4 reply into inbox messages. Each "+CMGL: &lt;idx&gt;,&lt;stat&gt;,,&lt;len&gt;" line is followed
    ''' by one PDU hex line; unparseable entries are skipped (the rest still list). A reply with no +CMGL: lines
    ''' (empty inbox, or an error) yields an empty list.</summary>
    Public Shared Function ParseCmglReply(reply As String) As List(Of InboxMessage)
        Dim result As New List(Of InboxMessage)()
        If String.IsNullOrEmpty(reply) Then Return result
        Dim lines As String() = reply.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, StringSplitOptions.RemoveEmptyEntries)
        Dim i As Integer = 0
        While i < lines.Length
            Dim line As String = lines(i).Trim()
            If line.StartsWith("+CMGL:", StringComparison.Ordinal) Then
                Dim head As String = line.Substring(6).Trim()
                Dim parts As String() = head.Split(","c)
                Dim idx As Integer = 0
                Dim stat As Integer = -1
                If parts.Length >= 2 AndAlso Integer.TryParse(parts(0).Trim(), idx) AndAlso Integer.TryParse(parts(1).Trim(), stat) Then
                    If i + 1 < lines.Length Then
                        Dim pduHex As String = lines(i + 1).Trim()
                        Dim msg As InboxMessage = Nothing
                        If DecodeDeliverPdu(pduHex, msg) Then
                            msg.Index = idx
                            msg.Status = InboxStatusLabel(stat)
                            result.Add(msg)
                        End If
                        i += 1
                    End If
                End If
            End If
            i += 1
        End While
        Return result
    End Function

    ' (Fixer F) Per-storage notes from the last ReadInboxSms call
    ' (e.g. "SM: 1، ME: 1، MT: 0" or a select failure), for the
    ' caller's log line. Empty when the fallback single read ran.
    Public Property InboxReadNotes As String = String.Empty

    ''' <summary>(Fixer F) Parses an AT+CPMS=? reply into the supported
    ''' message storages, in modem order (e.g. "+CPMS: ("SM","ME","MT"),..."
    ''' -> SM, ME, MT). Only the first group (mem1, the read store)
    ''' is used. Returns an empty list when the reply carries no
    ''' usable storage names. Pure logic, no port.</summary>
    Public Shared Function ParseCpmsTestReply(reply As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrEmpty(reply) Then Return result
        Dim start As Integer = reply.IndexOf("+CPMS:", StringComparison.Ordinal)
        If start < 0 Then Return result
        Dim openParen As Integer = reply.IndexOf("("c, start)
        Dim closeParen As Integer = If(openParen >= 0, reply.IndexOf(")"c, openParen), -1)
        If openParen < 0 OrElse closeParen < 0 Then Return result
        Dim group As String = reply.Substring(openParen + 1, closeParen - openParen - 1)
        For Each part As String In group.Split(","c)
            Dim name As String = part.Trim().Trim(""""c).Trim()
            If name.Length > 0 AndAlso Not result.Contains(name) Then result.Add(name)
        Next
        Return result
    End Function

    ''' <summary>(Fixer F) Selects one message storage for read/write/
    ''' receive (AT+CPMS="m","m","m"); falls back to the one-argument
    ''' form (AT+CPMS="m") when the modem rejects the three-argument
    ''' form. Returns True when the modem answered OK to either form.</summary>
    Private Function SelectInboxStorage(storage As String, Optional stopPredicate As Func(Of Boolean) = Nothing) As Boolean
        Dim three As String = SendCommand("AT+CPMS=""" & storage & """,""" & storage & """,""" & storage & """", "OK", 5000, stopPredicate)
        If ReplyMatchesExpected(three, "OK") Then Return True
        Dim one As String = SendCommand("AT+CPMS=""" & storage & """", "OK", 5000, stopPredicate)
        Return ReplyMatchesExpected(one, "OK")
    End Function

    ''' <summary>(Fixer F) Merges per-storage inbox reads in the given
    ''' order, tagging each row with its storage. Rows are deduped by
    ''' storage+index (indices collide across storages). Pure logic,
    ''' no port.</summary>
    Public Shared Function MergeInboxStorages(ordered As List(Of KeyValuePair(Of String, List(Of InboxMessage)))) As List(Of InboxMessage)
        Dim merged As New List(Of InboxMessage)()
        If ordered Is Nothing Then Return merged
        Dim seen As New HashSet(Of String)()
        For Each pair As KeyValuePair(Of String, List(Of InboxMessage)) In ordered
            If pair.Value Is Nothing Then Continue For
            For Each msg As InboxMessage In pair.Value
                If msg Is Nothing Then Continue For
                Dim key As String = pair.Key & "|" & msg.Index.ToString()
                If Not seen.Add(key) Then Continue For
                msg.Storage = pair.Key
                merged.Add(msg)
            Next
        Next
        Return merged
    End Function

    ''' <summary>(Fixer F) Reads the inbox from the currently selected
    ''' storage (AT+CMGL=4, PDU mode). Extracted unchanged from
    ''' ReadInboxSms: same replies, same error strings. Returns Nothing
    ''' on failure (errorText explains); an empty list means the storage
    ''' is empty.</summary>
    Private Function ReadInboxFromCurrentStorage(ByRef errorText As String, Optional stopPredicate As Func(Of Boolean) = Nothing) As List(Of InboxMessage)
        errorText = String.Empty
        Dim reply As String = SendCommand("AT+CMGL=4", "OK", 15000, stopPredicate)
        If String.IsNullOrWhiteSpace(reply) Then
            errorText = "تعذرت قراءة الوارد (AT+CMGL=4) — لا يوجد رد"
            Return Nothing
        End If
        Dim messages As List(Of InboxMessage) = ParseCmglReply(reply)
        If messages.Count > 0 Then Return messages
        If LineStartsWith(reply, "+CMS ERROR") OrElse LineStartsWith(reply, "+CME ERROR") OrElse LineStartsWith(reply, "ERROR") Then
            errorText = "تعذرت قراءة الوارد: " & reply.Trim()
            Return Nothing
        End If
        ' No +CMGL: lines and no error: the inbox is empty.
        Return messages
    End Function

    ''' <summary>Reads the SIM/modem inbox on the already-open port.
    ''' (Fixer F) Storage-aware: AT+CPMS=? lists the supported storages
    ''' and each of SM, ME, MT (supported only) is selected and read with
    ''' AT+CMGL=4, merging the rows (deduped by storage+index, each row
    ''' tagged with its storage). When the CPMS query fails the old
    ''' single read on the modem default storage runs instead. A fresh
    ''' read every call - nothing is cached and nothing is deleted or
    ''' marked read. On failure returns Nothing and errorText explains.</summary>
    Public Function ReadInboxSms(ByRef errorText As String, Optional stopPredicate As Func(Of Boolean) = Nothing) As List(Of InboxMessage)
        errorText = String.Empty
        InboxReadNotes = String.Empty
        If _port Is Nothing OrElse Not _port.IsOpen Then
            errorText = "المنفذ غير مفتوح"
            Return Nothing
        End If
        Dim modeReply As String = SendCommand("AT+CMGF=0", "OK", 5000, stopPredicate)
        If Not ReplyMatchesExpected(modeReply, "OK") Then
            errorText = "تعذر ضبط وضع PDU (AT+CMGF=0)" & If(String.IsNullOrWhiteSpace(modeReply), " — لا يوجد رد", " — رد المودم: " & modeReply.Trim())
            Return Nothing
        End If
        Dim supported As List(Of String) = Nothing
        Dim cpmsReply As String = SendCommand("AT+CPMS=?", "OK", 5000, stopPredicate)
        If ReplyMatchesExpected(cpmsReply, "OK") Then supported = ParseCpmsTestReply(cpmsReply)
        If supported Is Nothing OrElse supported.Count = 0 Then
            ' (Fixer F) The modem does not answer the storage query:
            ' fall back to the old single read on the default storage.
            Return ReadInboxFromCurrentStorage(errorText, stopPredicate)
        End If
        Dim ordered As New List(Of KeyValuePair(Of String, List(Of InboxMessage)))()
        Dim notes As New System.Text.StringBuilder()
        Dim anyRead As Boolean = False
        For Each storage As String In New String() {"SM", "ME", "MT"}
            If Not supported.Contains(storage) Then Continue For
            If Not SelectInboxStorage(storage, stopPredicate) Then
                notes.Append(storage).Append(": تعذر الاختيار؛ ")
                Continue For
            End If
            Dim perError As String = String.Empty
            Dim list As List(Of InboxMessage) = ReadInboxFromCurrentStorage(perError, stopPredicate)
            If list Is Nothing Then
                notes.Append(storage).Append(": ").Append(perError).Append("؛ ")
                Continue For
            End If
            anyRead = True
            ordered.Add(New KeyValuePair(Of String, List(Of InboxMessage))(storage, list))
            notes.Append(storage).Append(": ").Append(list.Count.ToString()).Append("؛ ")
        Next
        InboxReadNotes = notes.ToString().Trim()
        If Not anyRead Then
            errorText = "تعذرت قراءة الوارد من كل التخزينات — " & InboxReadNotes
            Return Nothing
        End If
        Return MergeInboxStorages(ordered)
    End Function

    ''' <summary>(Fixer F) One USSD session step result.</summary>
    Public Class UssdResult
        Public Property Text As String = String.Empty
        ' True when the network waits for more input (m=1): the caller
        ' shows a follow-up input and sends it with SendUssd.
        Public Property NeedsMoreInput As Boolean
        Public Property ErrorText As String = String.Empty
        Public Property RawReply As String = String.Empty
    End Class

    ''' <summary>(Fixer F) Decodes a +CUSD string by its DCS: 72 = UCS2 hex,
    ''' 4 = 8-bit data (Latin-1), anything else = GSM 7-bit packed hex.
    ''' A string that is not hex at all is returned as-is (some modems
    ''' send plain text). Pure logic, no port.</summary>
    Public Shared Function DecodeCusdString(str As String, dcs As Integer) As String
        If String.IsNullOrEmpty(str) Then Return String.Empty
        Dim bytes As Byte() = HexToBytes(str.Trim())
        If bytes Is Nothing Then Return str
        If dcs = 72 Then
            ' UCS2: whole 16-bit units only; a trailing odd byte cannot
            ' form a character and is dropped, never shown as '?'.
            Dim evenLen As Integer = bytes.Length - (bytes.Length Mod 2)
            If evenLen <= 0 Then Return String.Empty
            Return System.Text.Encoding.BigEndianUnicode.GetString(bytes, 0, evenLen)
        ElseIf dcs = 4 Then
            Return System.Text.Encoding.GetEncoding("ISO-8859-1").GetString(bytes)
        End If
        ' GSM 7-bit packed: septets fill the bytes; trailing NUL
        ' septets are packing padding, never message text.
        Dim septets As Integer = (bytes.Length * 8) \ 7
        If septets <= 0 Then Return String.Empty
        Return UnpackGsm7Bit(bytes, 0, septets).TrimEnd(ChrW(0))
    End Function

    ''' <summary>(Fixer F) Finds the +CUSD: line in a modem reply and
    ''' decodes it into result (Text, NeedsMoreInput for m=1). Returns
    ''' False when the reply holds no +CUSD: line at all. Pure logic,
    ''' no port.</summary>
    Public Shared Function ParseCusdReply(reply As String, ByRef result As UssdResult) As Boolean
        result = New UssdResult()
        If String.IsNullOrEmpty(reply) Then Return False
        result.RawReply = reply.Trim()
        For Each line As String In SplitLines(reply)
            Dim t As String = line.Trim()
            If Not t.StartsWith("+CUSD:", StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim rest As String = t.Substring(6).Trim()
            Dim m As Integer = 0
            Dim comma As Integer = rest.IndexOf(","c)
            If comma < 0 Then
                If Not Integer.TryParse(rest, m) Then Return False
                result.NeedsMoreInput = (m = 1)
                Return True
            End If
            If Not Integer.TryParse(rest.Substring(0, comma).Trim(), m) Then Return False
            result.NeedsMoreInput = (m = 1)
            Dim firstQuote As Integer = rest.IndexOf(""""c, comma)
            Dim lastQuote As Integer = rest.LastIndexOf(""""c)
            If firstQuote < 0 OrElse lastQuote <= firstQuote Then Return True
            Dim str As String = rest.Substring(firstQuote + 1, lastQuote - firstQuote - 1)
            Dim dcs As Integer = 15
            Dim after As String = rest.Substring(lastQuote + 1).Trim()
            If after.StartsWith(",", StringComparison.Ordinal) Then
                Dim dcsText As String = after.Substring(1).Trim()
                Dim parsed As Integer = 0
                If Integer.TryParse(dcsText, parsed) Then dcs = parsed
            End If
            result.Text = DecodeCusdString(str, dcs)
            Return True
        Next
        Return False
    End Function

    ''' <summary>(Fixer F) Reads the port briefly for a late +CUSD: line
    ''' (some networks answer OK first and send the USSD text as a later
    ''' URC). Returns everything read; bytes that are not part of a
    ''' +CUSD: line are re-preserved for the next command instead of
    ''' being dropped.</summary>
    Private Function GraceReadForCusd(timeoutMs As Integer) As String
        Dim buffer As New System.Text.StringBuilder()
        If _port Is Nothing OrElse Not _port.IsOpen Then Return String.Empty
        buffer.Append(TakePreservedText())
        Dim watch As New System.Diagnostics.Stopwatch()
        watch.Start()
        While watch.ElapsedMilliseconds < timeoutMs
            Dim chunk As String = String.Empty
            Try
                chunk = _port.ReadExisting()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("CUSD grace read failed: " & ex.Message)
                Exit While
            End Try
            If Not String.IsNullOrEmpty(chunk) Then buffer.Append(chunk)
            If LineStartsWith(buffer.ToString(), "+CUSD:") Then Exit While
            Threading.Thread.Sleep(100)
        End While
        Dim text As String = buffer.ToString()
        If Not LineStartsWith(text, "+CUSD:") AndAlso Not String.IsNullOrEmpty(text) Then
            ' (Fixer G) C-81D-L1: a final error line IS an answer, not
            ' leftovers - return it so SendUssd's grace error check
            ' can fire instead of reporting a timeout. Only non-error,
            ' non-CUSD text is re-preserved for the next read.
            If LineStartsWith(text, "+CME ERROR") OrElse LineStartsWith(text, "+CMS ERROR") OrElse LineStartsWith(text, "ERROR") Then
                Return text
            End If
            AppendPreservedFragment(text)
            Return String.Empty
        End If
        Return text
    End Function

    ''' <summary>(Fixer F) Sends one USSD step (AT+CUSD=1,"code",15) on the
    ''' already-open port and waits up to timeoutMs for the +CUSD:
    ''' reply (the network may answer in two parts: an OK first, then
    ''' the +CUSD line - a short grace read covers that). Returns the
    ''' decoded text and whether the session waits for more input
    ''' (NeedsMoreInput). On failure ErrorText explains in Arabic and
    ''' Text is empty. Never cancels the session itself - use
    ''' CancelUssd for that.</summary>
    Public Function SendUssd(codeOrInput As String, ByRef errorText As String, Optional timeoutMs As Integer = 60000) As UssdResult
        errorText = String.Empty
        Dim result As New UssdResult()
        If _port Is Nothing OrElse Not _port.IsOpen Then
            errorText = "المنفذ غير مفتوح"
            result.ErrorText = errorText
            Return result
        End If
        Dim code As String = If(codeOrInput, String.Empty).Replace(vbCr, String.Empty).Replace(vbLf, String.Empty).Replace("""", String.Empty)
        If String.IsNullOrWhiteSpace(code) Then
            errorText = "اكتب كود USSD أولاً"
            result.ErrorText = errorText
            Return result
        End If
        Dim reply As String = SendCommand("AT+CUSD=1,""" & code & """,15", "OK", Math.Max(5000, timeoutMs), Nothing)
        Dim parsed As New UssdResult()
        If ParseCusdReply(reply, parsed) Then Return parsed
        If LineStartsWith(reply, "+CME ERROR") OrElse LineStartsWith(reply, "+CMS ERROR") OrElse LineStartsWith(reply, "ERROR") Then
            errorText = "رفض المودم/الشبكة طلب USSD: " & reply.Trim()
            result.ErrorText = errorText
            result.RawReply = reply.Trim()
            Return result
        End If
        ' No +CUSD: line yet and no error: the network may still be
        ' answering (OK first, +CUSD: as a later URC) - grace-read
        ' briefly for the late line before calling it a timeout.
        Dim grace As String = GraceReadForCusd(5000)
        If ParseCusdReply(grace, parsed) Then Return parsed
        If LineStartsWith(grace, "+CME ERROR") OrElse LineStartsWith(grace, "+CMS ERROR") OrElse LineStartsWith(grace, "ERROR") Then
            errorText = "رفض المودم/الشبكة طلب USSD: " & grace.Trim()
            result.ErrorText = errorText
            result.RawReply = (reply & vbCrLf & grace).Trim()
            Return result
        End If
        errorText = "انتهت المهلة — لم يصل رد USSD من الشبكة"
        result.ErrorText = errorText
        result.RawReply = (reply & vbCrLf & grace).Trim()
        Return result
    End Function

    ''' <summary>(Fixer F) Cancels an active USSD session (AT+CUSD=2).
    ''' Returns True when the modem answered OK.</summary>
    Public Function CancelUssd(ByRef errorText As String) As Boolean
        errorText = String.Empty
        If _port Is Nothing OrElse Not _port.IsOpen Then
            errorText = "المنفذ غير مفتوح"
            Return False
        End If
        Dim reply As String = SendCommand("AT+CUSD=2", "OK", 15000, Nothing)
        If ReplyMatchesExpected(reply, "OK") Then Return True
        errorText = "تعذر إلغاء جلسة USSD" & If(String.IsNullOrWhiteSpace(reply), " — لا يوجد رد", " — رد المودم: " & reply.Trim())
        Return False
    End Function


    ''' <summary>
    ''' Sends ESC (followed by CR, so it can never be buffered into the
    ''' next command line) to abort a stale "&gt;" prompt left by an
    ''' earlier failed send, then drains the buffer so the next command
    ''' is not swallowed as SMS text. ESC is skipped when the last reply
    ''' already proved command mode. The drain reads until ~100ms of
    ''' quiet (never discarding, so buffered bytes are really consumed
    ''' and examined, not destroyed). Command mode is NOT asserted from the drain alone:
    ''' an AT -> OK probe must prove it first. When stopping and
    ''' command mode is already proven, the drain is skipped entirely;
    ''' the drain loop also exits promptly on stop. Returns the text
    ''' consumed by the drain ONLY,
    ''' so a very-late +CMGS: arriving during the drain is not lost
    ''' (the caller checks it before reporting failure); the AT probe
    ''' reply is returned separately in probeText, so the probe's own
    ''' OK can never be mistaken for the +CMGS: confirmation.
    ''' probeTimeoutMs caps the probe wait (default 3000ms; pass ~500
    ''' on the stop path, where the port closes right after and the
    ''' probe result is discarded). If the drain still shows a "&gt;"
    ''' prompt line, the ESC abort did not take effect: the probe is
    ''' skipped (its AT would be swallowed as SMS text) and the
    ''' drained text is returned with command mode unproven. When the
    ''' drain holds neither an OK line nor a "&gt;" line the abort is
    ''' unproven, so a second ESC+CR is sent and the buffer is
    ''' re-drained briefly; D3: a silent re-drain no longer skips the
    ''' probe (a bare ESC on an idle modem in command mode is silent,
    ''' so no reply is normal) — only a "&gt;" line still skips it.
    ''' forceDrain (default False) skips both
    ''' early returns, so a recovery drain actually reads even when
    ''' command mode was already proven. The probe does NOT discard
    ''' the input buffer first, so a +CMGS: arriving between the
    ''' drain and the probe survives into the probe reply.
    ''' </summary>
    ' HARDENING (Reviewer D note): the wrapper forwards an optional
    ' parkRejoinedCmgs to the Core (default True, as before), so a
    ' future in-send recovery site using the wrapper can pass False.
    ' No call-site behaviour changes: every current caller keeps the default.
    Private Function SendEscapeAndDrain(Optional stopPredicate As Func(Of Boolean) = Nothing, Optional probeTimeoutMs As Integer = 3000, Optional forceDrain As Boolean = False, Optional parkRejoinedCmgs As Boolean = True) As String
        Dim probeText As String = String.Empty
        Return SendEscapeAndDrainCore(stopPredicate, probeTimeoutMs, probeText, forceDrain, parkRejoinedCmgs)
    End Function

    ''' <summary>SendEscapeAndDrain with the probe reply returned
    ''' separately in probeText (see SendEscapeAndDrain). forceDrain
    ''' skips both early returns so a recovery drain actually reads.</summary>
    Private Function SendEscapeAndDrainCore(stopPredicate As Func(Of Boolean), probeTimeoutMs As Integer, ByRef probeText As String, Optional forceDrain As Boolean = False, Optional parkRejoinedCmgs As Boolean = True) As String
        probeText = String.Empty
        If _port Is Nothing OrElse Not _port.IsOpen Then Return String.Empty
        ' Command mode already proven and no stop pending: a full
        ' ESC-drain + AT probe on every SMS is pure overhead.
        If _commandModeKnown AndAlso (stopPredicate Is Nothing OrElse Not stopPredicate()) AndAlso Not forceDrain Then Return String.Empty
        ' Stopping with command mode already proven: nothing to abort.
        If _commandModeKnown AndAlso stopPredicate IsNot Nothing AndAlso stopPredicate() AndAlso Not forceDrain Then Return String.Empty
        Dim drained As New System.Text.StringBuilder()
        ' Bytes preserved by an earlier SendCommand (a late +CMGS:
        ' read ahead of a command) belong to this drain's result.
        drained.Append(TakePreservedText())
        Try
            If Not _commandModeKnown Then
                ' ESC + CR as bytes: the CR terminates any command
                ' line the ESC may have been buffered into.
                _port.Write(New Byte() {CByte(&H1B), CByte(&HD)}, 0, 2)
                Dim waited As Integer = 0
                While waited < 100
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                    Threading.Thread.Sleep(20)
                    waited += 20
                End While
            End If
            Dim quietWatch As New System.Diagnostics.Stopwatch()
            quietWatch.Start()
            Dim lastDataMs As Long = 0
            While quietWatch.ElapsedMilliseconds < 1000
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                Dim chunk As String = _port.ReadExisting()
                If Not String.IsNullOrEmpty(chunk) Then
                    drained.Append(chunk)
                    lastDataMs = quietWatch.ElapsedMilliseconds
                    System.Diagnostics.Debug.WriteLine("ESC drain: " & chunk)
                ElseIf quietWatch.ElapsedMilliseconds - lastDataMs >= 100 Then
                    Exit While
                End If
                Threading.Thread.Sleep(20)
            End While
            quietWatch.Stop()
            ' Still at the ">" prompt: the ESC abort did not take
            ' effect. Sending the probe AT now would swallow it as
            ' SMS text and burn the whole probe timeout, so skip the
            ' probe and report the modem as still busy (command mode
            ' stays unproven).
            If HasPromptLine(drained.ToString()) Then
                _commandModeKnown = False
                Return DrainResult(drained)
            End If
            ' Abort unproven: the drain holds neither an OK line (the
            ' abort's final result) nor a ">" prompt line, so a
            ' still-in-text-mode modem would swallow the probe AT as
            ' SMS body. One more ESC+CR, then a brief re-drain
            ' (~300ms of quiet). D3: an OK in the re-drain still
            ' proves the abort, but a silent re-drain no longer
            ' blocks the probe — a bare ESC on an idle modem in
            ' command mode gets no reply at all, so silence is
            ' normal; only a ">" line still skips the probe.
            If Not LineMatchesExpected(drained.ToString(), "OK") Then
                Try
                    _port.Write(New Byte() {CByte(&H1B), CByte(&HD)}, 0, 2)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("Second ESC failed: " & ex.Message)
                End Try
                Dim reWatch As New System.Diagnostics.Stopwatch()
                reWatch.Start()
                Dim reLastDataMs As Long = 0
                While reWatch.ElapsedMilliseconds < 1500
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                    Dim reChunk As String = _port.ReadExisting()
                    If Not String.IsNullOrEmpty(reChunk) Then
                        drained.Append(reChunk)
                        reLastDataMs = reWatch.ElapsedMilliseconds
                        System.Diagnostics.Debug.WriteLine("ESC re-drain: " & reChunk)
                    ElseIf reWatch.ElapsedMilliseconds - reLastDataMs >= 300 Then
                        Exit While
                    End If
                    Threading.Thread.Sleep(20)
                End While
                reWatch.Stop()
                ' D3: still no OK — but a bare ESC on an idle modem
                ' in command mode is silent, so no reply is NORMAL
                ' and must not block the probe: probe to confirm
                ' instead of skipping it. Only a ">" prompt line
                ' still skips the probe (its AT would be swallowed
                ' as SMS text). The abort at the CMGF-recovery
                ' call site therefore fires only after the probe
                ' actually ran and failed, not on a skipped probe.
                If HasPromptLine(drained.ToString()) Then
                    _commandModeKnown = False
                    Return DrainResult(drained)
                End If
            End If
            ' No DiscardInBuffer here: bytes arriving between the
            ' drain loop's exit and the probe must survive into the
            ' probe reply instead of being destroyed.
            ' Do not assert command mode without proof: probe AT -> OK.
            _commandModeKnown = False
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Return DrainResult(drained)
            ' FIX B-M1/B-L1: hold back the drain's trailing unterminated
            ' line BEFORE the probe runs, preserving it in
            ' _preservedText, so the probe's SendCommand seeds it and
            ' rejoins it with its continuation exactly like SendCommand
            ' / WaitForResult already do. Without this, a +CMGS: line
            ' split across the drain/probe boundary (drain ends "... +CM",
            ' probe brings "GS: 12\r\nOK\r\n") was lost entirely: the
            ' fragment stayed inside drained, the probe carried only
            ' the bare continuation, JoinDrainAndProbe's inserted CRLF
            ' kept them two non-+CMGS: lines (count 0), and the held
            ' "+CM" could never complete (its bytes were consumed).
            ' The hold-back is generalised to ANY trailing unterminated
            ' line (B-L1, as Strip already does), so drained always
            ' ends at a line boundary. DrainResult's post-probe
            ' hold-back then finds nothing new for that fragment and
            ' cannot double-preserve it (it also guards duplicates).
            Dim drainedBeforeProbe As String = drained.ToString()
            Dim preProbeFragment As String = String.Empty
            SplitTrailingUnterminatedLine(drainedBeforeProbe, preProbeFragment)
            If preProbeFragment <> String.Empty Then PrependPreservedFragment(preProbeFragment)
            drained.Clear()
            drained.Append(drainedBeforeProbe)
            ' No discard inside the probe: a +CMGS: arriving between
            ' the drain above and this write must survive into the
            ' probe reply instead of being destroyed.
            Dim probe As String = SendCommand("AT", "OK", probeTimeoutMs, stopPredicate, discardBeforeWrite:=False, parkStaleCmgs:=parkRejoinedCmgs)
            probeText = probe
            _commandModeKnown = ReplyMatchesExpected(probe, "OK")
            Return DrainResult(drained)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("ESC abort failed: " & ex.Message)
            _commandModeKnown = False
            Return DrainResult(drained)
        End Try
    End Function

    ''' <summary>
    ''' Sends one SMS on an already-open port.
    ''' Arabic (or any non-GSM character) -> UCS2 with UTF-16BE hex.
    ''' Latin-only -> GSM 7-bit, sent as raw septet bytes (GSM 03.38),
    ''' never via Write(String), so accented Latin survives.
    ''' NEVER sends the message body without first receiving the ">" prompt.
    ''' Long messages are split into segments (GSM 160 septets,
    ''' UCS2 70 chars - each segment IS a separate standalone SMS)
    ''' and each segment is sent as its own CMGS; success is reported
    ''' only when every segment succeeds. A failed segment reports how
    ''' many standalone messages were already sent.
    ''' Empty messages and un-normalisable numbers are rejected here
    ''' and never reach the modem. When stopPredicate is supplied the
    ''' send aborts between segments and inside waits; on stop the
    ''' modem is ESC-drained and Cancelled is set.
    ''' </summary>
    ''' <summary>(Round 55) True for exceptions that mean the modem's
    ''' port itself is gone (unplugged, closed by another program,
    ''' driver fault): IOException, InvalidOperationException
    ''' (covers ObjectDisposedException / "port is closed") and
    ''' UnauthorizedAccessException. TimeoutException is included
    ''' too: a serial write that cannot complete in 5 s means the
    ''' port is not accepting bytes at all. Anything else (an
    ''' encoding or argument error) is NOT a dead port.</summary>
    Private Shared Function IsPortDeadException(ex As Exception) As Boolean
        Return TypeOf ex Is System.IO.IOException OrElse TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is UnauthorizedAccessException OrElse TypeOf ex Is TimeoutException
    End Function

    Public Function SendSmsConnected(number As String, message As String, Optional stopPredicate As Func(Of Boolean) = Nothing, Optional reuseConcatRef As Integer = 0) As SmsResult
        Dim result As New SmsResult()
        ' Set True immediately after a Ctrl-Z byte is written for the
        ' current segment (reset at the start of each segment
        ' iteration): only then may a drained +CMGS: be counted for
        ' this segment (see CountGuardedCmgs).
        Dim ctrlZSent As Boolean = False
        ' (Round 34) B-M1: set True immediately BEFORE a Ctrl-Z byte
        ' is written for the current segment (reset per segment with
        ' ctrlZSent): the write itself can throw after the 0x1A byte
        ' reached the modem, leaving ctrlZSent False while the segment
        ' was (probably) sent — the clause sites key on this flag, never
        ' the counting guard in CountGuardedCmgs.
        Dim ctrlZAttempted As Boolean = False
        ' Set when any counted +CMGS: lacked its own trailing OK.
        Dim sawCmgsWithoutOk As Boolean = False
        If String.IsNullOrWhiteSpace(number) OrElse String.IsNullOrWhiteSpace(message) Then
            result.Success = False
            result.Message = "رقم أو رسالة فارغة - لم يتم الإرسال"
            Return result
        End If
        Dim normalized As String = String.Empty
        If Not PhoneNormalizer.NormalizePhone(number, normalized) Then
            result.Success = False
            result.Message = "رقم غير صالح: " & number
            Return result
        End If
        If _port Is Nothing OrElse Not _port.IsOpen Then
            result.Success = False
            result.Message = "المنفذ غير مفتوح"
            Return result
        End If

        Try
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                SendEscapeAndDrain(stopPredicate, 500)
                MarkCancelled(result, sawCmgsWithoutOk)
                Return result
            End If

            ' Clear any stale ">" prompt left by an earlier failed send.
            SendEscapeAndDrain(stopPredicate)

            ' (Round 50) PDU mode for UCS2: the E1550 firmware rejects
            ' the UCS2 TEXT-mode AT+CMGS (live: Arabic failed in 1s
            ' while English worked), so Arabic is sent in PDU mode.
            ' The mode this send needs is set here, at the start of
            ' every send, so no restore is needed: the next send sets
            ' its own mode first, and the connection test (AT/CPIN/
            ' CSQ) never depends on CMGF.
            Dim useUcs2 As Boolean = PhoneNormalizer.NeedsUcs2(message)
            ' (Round 56) Concatenated SMS: the split is computed
            ' here, before the mode is chosen — a message longer
            ' than one part needs PDU mode even when it is GSM
            ' (text mode cannot carry a UDH), so the phone joins
            ' the parts into ONE displayed message. A single part
            ' keeps its existing path exactly (UCS2 PDU, GSM text).
            Dim segments As List(Of String) = PhoneNormalizer.SplitMessage(message)
            Dim usePdu As Boolean = useUcs2 OrElse segments.Count > 1
            Dim concatRef As Integer = 0
            ' (Round 57) C-55-L1: a retried row reuses its first
            ' reference (reuseConcatRef, 1..255) so the phone
            ' completes the SAME concatenated message; any other
            ' multi-segment send draws a fresh Shared ref. The
            ' reused ref is NOT drawn from the counter.
            If segments.Count > 1 Then
                If reuseConcatRef >= 1 AndAlso reuseConcatRef <= 255 Then
                    concatRef = reuseConcatRef
                Else
                    concatRef = NextConcatRef()
                End If
                result.ConcatRef = concatRef
            End If
            Dim response As String = SendCommand(If(usePdu, "AT+CMGF=0", "AT+CMGF=1"), "OK", 5000, stopPredicate)
            ' A +CMGS: in this reply arrived before any Ctrl-Z for
            ' this send, so the guard inside refuses to count it;
            ' the reply (with preserved bytes prepended) is still
            ' examined here instead of being discarded.
            Dim setupRefused As Boolean = False
            CountGuardedCmgs(response, String.Empty, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                SendEscapeAndDrain(stopPredicate, 500)
                MarkCancelled(result, sawCmgsWithoutOk)
                Return result
            End If
            If Not ReplyMatchesExpected(response, "OK") Then
                ' One recovery attempt: force ESC + AT probe, then retry
                ' CMGF once. If command mode cannot be proven, abort
                ' the whole send instead of failing every number.
                _commandModeKnown = False
                SendEscapeAndDrain(stopPredicate)
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                    MarkCancelled(result, sawCmgsWithoutOk)
                    Return result
                End If
                If Not _commandModeKnown Then
                    result.Success = False
                    result.Message = "تعذر الوصول إلى وضع الأوامر بالمودم — أعد اختبار الاتصال قبل الإرسال"
                    Return result
                End If
                response = SendCommand(If(usePdu, "AT+CMGF=0", "AT+CMGF=1"), "OK", 5000, stopPredicate)
                CountGuardedCmgs(response, String.Empty, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                    SendEscapeAndDrain(stopPredicate, 500)
                    MarkCancelled(result, sawCmgsWithoutOk)
                    Return result
                End If
                If Not ReplyMatchesExpected(response, "OK") Then
                    result.Success = False
                    result.Message = "فشل AT+CMGF=" & If(usePdu, "0", "1") & ": " & response.Trim()
                    Return result
                End If
            End If

            ' Charset actually in effect for this send: GSM septet
            ' bytes are only valid under CSCS="GSM". Under the IRA
            ' fallback the same bytes would silently corrupt @ $ £
            ' and accented Latin, so IRA sends pure ASCII instead.
            Dim activeCharset As String = "GSM"
            If usePdu Then
                ' (Round 50/56) PDU mode: CSCS is a text-mode setting
                ' and the PDU carries its own DCS (08 for UCS2, 00
                ' for GSM), so no CSCS command is sent; response
                ' still holds the CMGF=0 OK reply verified above, so
                ' the shared check below passes without a CSCS
                ' round-trip.
                activeCharset = If(useUcs2, "UCS2", "PDU")
            Else
                response = SendCommand("AT+CSCS=""GSM""", "OK", 5000, stopPredicate)
                CountGuardedCmgs(response, String.Empty, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                If Not ReplyMatchesExpected(response, "OK") Then
                    response = SendCommand("AT+CSCS=""IRA""", "OK", 5000, stopPredicate)
                    CountGuardedCmgs(response, String.Empty, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                    If ReplyMatchesExpected(response, "OK") Then activeCharset = "IRA"
                End If
            End If
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                SendEscapeAndDrain(stopPredicate, 500)
                MarkCancelled(result, sawCmgsWithoutOk)
                Return result
            End If
            If Not ReplyMatchesExpected(response, "OK") Then
                result.Success = False
                result.Message = "فشل AT+CSCS: " & response.Trim()
                Return result
            End If

            If segments.Count = 0 Then
                result.Success = False
                result.Message = "رسالة فارغة - لم يتم الإرسال"
                Return result
            End If
            result.SegmentsTotal = segments.Count

            ' A lone surrogate in a UCS2 message would silently
            ' become U+FFFD in ToUcs2Hex (PhoneNormalizer) (Round 51):
            ' fail the row loudly before any segment is sent instead.
            If useUcs2 AndAlso HasUnpairedSurrogate(message) Then
                result.Success = False
                result.Message = "الرسالة تحتوي حرفاً غير مكتمل (surrogate وحيد) لا يمكن ترميزه في UCS2 — احذف الحرف أو أعد كتابة الرسالة (لم يتم إرسال أي مقطع)"
                Return result
            End If

            For segIndex As Integer = 0 To segments.Count - 1
                ctrlZSent = False
                ctrlZAttempted = False  ' (Round 34) B-M1: reset per segment with ctrlZSent
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                    SendEscapeAndDrain(stopPredicate, 500)
                    MarkCancelled(result, sawCmgsWithoutOk)
                    Return result
                End If
                Dim segment As String = segments(segIndex)
                Dim prompt As String = String.Empty
                ' The flag may still be True from the CSCS/CMGF OK
                ' replies: clear it before AT+CMGS so a missed ">"
                ' prompt cannot leave SendEscapeAndDrain thinking the
                ' modem is already in command mode.
                _commandModeKnown = False
                ' (Round 50/56) PDU mode: build the SMS-SUBMIT PDU
                ' first - its TPDU length goes in AT+CMGS. A build
                ' failure means bad input, never a modem reply, so
                ' fail the row loudly before any byte is sent.
                Dim pduHex As String = String.Empty
                Dim tpduOctets As Integer = 0
                If usePdu Then
                    If useUcs2 Then
                        ' (Round 56) Concatenated UCS2: every part
                        ' carries the same reference so the phone
                        ' joins them; a single part builds exactly
                        ' the old PDU (no UDH arguments).
                        If segments.Count > 1 Then
                            pduHex = PhoneNormalizer.BuildSmsSubmitPdu(normalized, segment, concatRef, segments.Count, segIndex + 1)
                        Else
                            pduHex = PhoneNormalizer.BuildSmsSubmitPdu(normalized, segment)
                        End If
                    Else
                        ' (Round 56) Long GSM message: PDU mode with
                        ' the same UDH scheme — text mode cannot
                        ' carry a concatenation header.
                        pduHex = BuildGsmSubmitPdu(normalized, segment, concatRef, segments.Count, segIndex + 1)
                    End If
                    tpduOctets = PhoneNormalizer.SmsSubmitTpduOctets(pduHex)
                    If tpduOctets < 0 Then
                        result.Success = False
                        result.Message = "تعذر بناء رسالة PDU للرقم " & normalized & " (تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " رسائل منفصلة قبل الفشل)" & If(result.SegmentsSent > 0, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)  ' (Round 50) PDU build failure - before any byte is sent for this segment.
                        Return result
                    End If
                    prompt = SendCommand("AT+CMGS=" & tpduOctets.ToString(), ">", 10000, stopPredicate)
                Else
                    prompt = SendCommand("AT+CMGS=""" & normalized & """", ">", 10000, stopPredicate)
                End If
                ' A +CMGS: in the prompt reply is from a previous
                ' segment (no Ctrl-Z has been sent for this one), so
                ' the guard refuses to count it; it is still examined.
                CountGuardedCmgs(prompt, String.Empty, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                    SendEscapeAndDrain(stopPredicate, 500)
                    MarkCancelled(result, sawCmgsWithoutOk)
                    Return result
                End If
                If Not HasPromptLine(prompt) Then
                    ' Route through the Core drain (not the wrapper)
                    ' so a +CMGS: arriving in the probe window is
                    ' examined instead of discarded; the guard still
                    ' refuses to count it (no Ctrl-Z for this segment).
                    Dim promptProbeText As String = String.Empty
                    Dim promptDrained As String = SendEscapeAndDrainCore(stopPredicate, 3000, promptProbeText, parkRejoinedCmgs:=False)
                    CountGuardedCmgs(promptDrained, promptProbeText, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                    result.Success = False
                    ' (Round 50) include the modem's raw reply to
                    ' AT+CMGS (flattened, first 200 chars) so a
                    ' prompt failure is diagnosable from the log
                    ' alone - the v1.1 log showed only this generic
                    ' text for the E1550's fast ERROR on the UCS2
                    ' text-mode AT+CMGS.
                    Dim promptReply As String = If(prompt, String.Empty).Trim().Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
                    If promptReply.Length > 200 Then promptReply = promptReply.Substring(0, 200)
                    result.Message = "تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل" & If(result.SegmentsSent > 0, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty) & If(promptReply = String.Empty, String.Empty, " — رد المودم: " & promptReply)  ' (Round 33) B-observation: prompt failure - no Ctrl-Z for this segment yet, Sent = prior only; clause only when Sent>0.
                    Return result
                End If
                _commandModeKnown = False

                Try
                    If usePdu Then
                        ' (Round 50/56) PDU hex digits are pure ASCII
                        ' (UCS2 or GSM PDU); Ctrl-Z is emitted as a
                        ' byte, never via Write(String).
                        Dim pduBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(pduHex)
                        _port.Write(pduBytes, 0, pduBytes.Length)
                        ctrlZAttempted = True  ' (Round 34) B-M1: before the Ctrl-Z write — it can throw after the byte went out
                        _port.Write(New Byte() {CByte(&H1A)}, 0, 1)
                        ctrlZSent = True
                    ElseIf activeCharset = "IRA" Then
                        ' Under CSCS="IRA" the modem reads plain ASCII
                        ' and Encoding.ASCII maps every non-ASCII char
                        ' to '?': an accented Latin char would be sent
                        ' corrupted and reported as success. Fail the
                        ' segment instead, naming the first lossy char.
                        Dim lossyChar As Char = ChrW(0)
                        For Each ch As Char In segment
                            If AscW(ch) > &H7F Then
                                lossyChar = ch
                                Exit For
                            End If
                        Next
                        If lossyChar <> ChrW(0) Then
                            Dim iraProbeText As String = String.Empty
                            Dim iraDrained As String = SendEscapeAndDrainCore(stopPredicate, 3000, iraProbeText, parkRejoinedCmgs:=False)
                            CountGuardedCmgs(iraDrained, iraProbeText, ctrlZSent, result, sawCmgsWithoutOk, setupRefused)
                            result.Success = False
                            ' FIX B-L1 (Round 26): re-send warning added -
                            ' the counts were already here, the warning was not.
                            result.Message = "حرف لاتيني مشكّل غير مدعوم في ترميز IRA: '" & lossyChar & "' — الرسالة تحتوي أحرفاً لاتينية مشكّلة لا يمكن إرسالها بهذا الترميز (تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " رسائل منفصلة قبل الفشل)" & If(result.SegmentsSent > 0, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)  ' (Round 33) B-observation: IRA pre-write - no Ctrl-Z for this segment yet, Sent = prior only; clause only when Sent>0.
                            Return result
                        End If
                        Dim asciiBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(segment)
                        _port.Write(asciiBytes, 0, asciiBytes.Length)
                        ctrlZAttempted = True  ' (Round 34) B-M1: before the Ctrl-Z write — it can throw after the byte went out
                        _port.Write(New Byte() {CByte(&H1A)}, 0, 1)
                        ctrlZSent = True
                    Else
                        Dim badChar As Char = ChrW(0)
                        Dim badIndex As Integer = -1
                        Dim gsmBytes As Byte() = GsmEncodeSegment(segment, badChar, badIndex)
                        If gsmBytes Is Nothing Then
                            Dim gsmProbeText As String = String.Empty
                            Dim gsmDrained As String = SendEscapeAndDrainCore(stopPredicate, 3000, gsmProbeText, parkRejoinedCmgs:=False)
                            Dim gsmRefused As Boolean = False
                            CountGuardedCmgs(gsmDrained, gsmProbeText, ctrlZSent, result, sawCmgsWithoutOk, gsmRefused)
                            result.Success = False
                            ' FIX B-L1 (Round 26): same one-line warning
                            ' addition as the IRA site above - identical
                            ' pattern (counts present, warning missing).
                            result.Message = "حرف غير مدعوم في GSM: '" & badChar & "' (الموضع " & (badIndex + 1).ToString() & ") (تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " رسائل منفصلة قبل الفشل)" & If(result.SegmentsSent > 0, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)  ' (Round 33) B-observation: GSM pre-write - no Ctrl-Z for this segment yet, Sent = prior only; clause only when Sent>0.
                            Return result
                        End If
                        _port.Write(gsmBytes, 0, gsmBytes.Length)
                        ctrlZAttempted = True  ' (Round 34) B-M1: before the Ctrl-Z write — it can throw after the byte went out
                        _port.Write(New Byte() {CByte(&H1A)}, 0, 1)
                        ctrlZSent = True
                    End If
                Catch ex As Exception
                    ' Check the drained text exactly as the late-+CMGS:
                    ' recovery after WaitForResult does below: a
                    ' very-late +CMGS: means the segment WAS sent -
                    ' count it and continue with the next segment
                    ' instead of failing the row. Only when a Ctrl-Z
                    ' was actually sent for this segment (the guard
                    ' inside CountGuardedCmgs): an exception on the
                    ' body write, before Ctrl-Z, must not count a
                    ' stale +CMGS: for this segment.
                    Dim writeProbeText As String = String.Empty
                    Dim writeDrainedText As String = SendEscapeAndDrainCore(stopPredicate, 3000, writeProbeText, forceDrain:=True, parkRejoinedCmgs:=False)
                    Dim writeRefused As Boolean = False
                    If CountGuardedCmgs(writeDrainedText, writeProbeText, ctrlZSent, result, sawCmgsWithoutOk, writeRefused) > 0 Then
                        If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                            SendEscapeAndDrain(stopPredicate, 500)
                            MarkCancelled(result, sawCmgsWithoutOk)
                            Return result
                        End If
                        ' B-B6: continue only when the drain's probe
                        ' proved command mode. In the stuck-text-mode
                        ' case the next AT+CMGS line would be swallowed
                        ' as body and its body + Ctrl-Z could complete
                        ' a garbage SMS to the previous recipient —
                        ' abort the row instead, noting how many
                        ' segments were sent so the user knows a
                        ' re-send would double-send them.
                        If Not _commandModeKnown Then
                            ' (Round 55) The body/Ctrl-Z write itself threw a
                            ' port exception: the modem is dead (the drain above
                            ' proved no +CMGS: was counted for this segment, so
                            ' nothing was sent by it). A non-port exception
                            ' here leaves PortDead False.
                            result.PortDead = IsPortDeadException(ex)
                            result.Success = False
                            result.Message = "تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل — لم يتم التحقق من وضع الأوامر بعد الاسترداد، فأُوقف الإرسال لتجنب إرسال مكرر — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها" & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)
                            Return result
                        End If
                        Continue For
                    End If
                    ' (Round 55) As the not-proven branch above: the
                    ' write threw — a port exception here means the
                    ' modem is dead.
                    result.PortDead = IsPortDeadException(ex)
                    result.Success = False
                    result.Message = "تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل" & If(result.SegmentsSent > 0 OrElse ctrlZAttempted, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)  ' (Round 34) B-M1: write-Catch fall-through - CountGuardedCmgs counted nothing and ctrlZSent is False at every reachable entry here, but the Ctrl-Z byte may already have gone out (the write can throw after transmitting), hence ctrlZAttempted; clause when prior Sent>0 or a Ctrl-Z was attempted.
                    Return result
                End Try

                ' (Round 56) Wait up to 60 seconds for +CMGS: / OK /
                ' ERROR: live E1550 confirmations arrived at ~26-28 s,
                ' at the old 30 s edge, so slow confirmations timed
                ' out as failures. A stop is still checked every
                ' ~50 ms poll, so Stop answers quickly. No auto-retry
                ' of a timed-out submit: the modem may have sent it
                ' while the confirmation was missed (the late-+CMGS:
                ' machinery and the re-send warning cover that).
                Dim cmgsErrorTrailed As Boolean = False
                response = WaitForResult(60000, stopPredicate)
                Dim hasCmgs As Boolean = LineStartsWith(response, "+CMGS:")
                ' (Round 45) B-H2: a +CMGS: line whose trailing
                ' result is an ERROR is never this segment's
                ' confirmation (see CountGuardedCmgs above) -
                ' park all lines found and fall through to the
                ' failure message with the re-send clause,
                ' instead of counting the segment.
                If hasCmgs AndAlso CurrentCmgsHasTrailingError(response) Then
                    Dim errFlags As System.Collections.Generic.List(Of Boolean) = CmgsLineConfirmations(response)
                    Dim errCount As Integer = errFlags.Count
                    Dim errCountedConfirmed As Boolean = True
                    Dim errParkedWithoutOk As Boolean = False
                    Dim errParkedWithoutOkCount As Integer = 0
                    SplitCmgsFlags(errFlags, 0, errCountedConfirmed, errParkedWithoutOk, errParkedWithoutOkCount, CmgsLineReferences(response))
                    ParkPendingLateCmgs(errCount, errParkedWithoutOkCount)
                    cmgsErrorTrailed = True
                End If
                If hasCmgs AndAlso Not cmgsErrorTrailed Then
                    ' D2: count the +CMGS: lines actually present
                    ' (a reply can hold two — a late previous one
                    ' plus this segment's) and add them bounded, so
                    ' SegmentsSent can never pass SegmentsTotal.
                    ' Any excess over the bounded add belongs to an
                    ' earlier segment: park it in the D1 pending
                    ' store (D1/D2) instead of dropping it.
                    ' B-B1: at most ONE line here belongs to the
                    ' current segment — the line with the GREATEST
                    ' message reference; earlier lines are previous
                    ' segments/rows and are parked, not counted.
                    Dim cmgsFlags As List(Of Boolean) = CmgsLineConfirmations(response)
                    Dim cmgsCount As Integer = cmgsFlags.Count
                    Dim cmgsAdded As Integer = Math.Min(BoundedCmgsAdd(result, cmgsCount), 1)
                    result.SegmentsSent += cmgsAdded
                    Dim cmgsCountedConfirmed As Boolean = True
                    Dim cmgsParkedWithoutOk As Boolean = False
                    Dim cmgsParkedWithoutOkCount As Integer = 0
                    SplitCmgsFlags(cmgsFlags, cmgsAdded, cmgsCountedConfirmed, cmgsParkedWithoutOk, cmgsParkedWithoutOkCount, CmgsLineReferences(response))
                    ' B-B5: the current segment's OK-status comes from
                    ' the COUNTED line (greatest mr; last-confirmed
                    ' fallback) only; the parked lines' status goes
                    ' to the pending store's withoutOk flag.
                    If cmgsAdded > 0 AndAlso Not cmgsCountedConfirmed Then sawCmgsWithoutOk = True
                    ParkPendingLateCmgs(cmgsCount - cmgsAdded, cmgsParkedWithoutOkCount)
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                        SendEscapeAndDrain(stopPredicate, 500)
                        MarkCancelled(result, sawCmgsWithoutOk)
                        Return result
                    End If
                    Continue For
                End If
                ' No +CMGS: yet: it may have arrived just as the wait
                ' was cut short by a stop/timeout, and the ESC drain
                ' below would discard it. Grace-read first so a sent
                ' segment is not undercounted.
                Dim graceText As String = String.Empty
                If Not LineStartsWith(response, "ERROR") _
                    AndAlso Not LineStartsWith(response, "+CMS ERROR") _
                    AndAlso Not LineStartsWith(response, "+CME ERROR") Then
                    graceText = GraceReadForCmgs(1500, stopPredicate)
                End If
                ' (Round 45) B-H2: same refusal as the wait branch
                ' above - an error-trailed +CMGS: in the grace
                ' window is a previous segment's late confirmation
                ' and this segment's failure, never a sent segment.
                If LineStartsWith(graceText, "+CMGS:") AndAlso CurrentCmgsHasTrailingError(graceText) Then
                    Dim graceErrFlags As System.Collections.Generic.List(Of Boolean) = CmgsLineConfirmations(graceText)
                    Dim graceErrCount As Integer = graceErrFlags.Count
                    Dim graceErrCountedConfirmed As Boolean = True
                    Dim graceErrParkedWithoutOk As Boolean = False
                    Dim graceErrParkedWithoutOkCount As Integer = 0
                    SplitCmgsFlags(graceErrFlags, 0, graceErrCountedConfirmed, graceErrParkedWithoutOk, graceErrParkedWithoutOkCount, CmgsLineReferences(graceText))
                    ParkPendingLateCmgs(graceErrCount, graceErrParkedWithoutOkCount)
                    cmgsErrorTrailed = True
                End If
                If LineStartsWith(graceText, "+CMGS:") AndAlso Not cmgsErrorTrailed Then
                    ' D2: as above — count the lines present, add
                    ' bounded; park any excess (previous-segment
                    ' lines) in the D1 pending store (D1/D2).
                    ' B-B1/B-B5: as the wait path above — at most
                    ' ONE line (the one with the GREATEST message
                    ' reference) belongs to the current segment;
                    ' earlier lines are parked, and only the counted
                    ' line's OK-status marks this segment.
                    Dim graceFlags As List(Of Boolean) = CmgsLineConfirmations(graceText)
                    Dim graceCount As Integer = graceFlags.Count
                    Dim graceAdded As Integer = Math.Min(BoundedCmgsAdd(result, graceCount), 1)
                    result.SegmentsSent += graceAdded
                    Dim graceCountedConfirmed As Boolean = True
                    Dim graceParkedWithoutOk As Boolean = False
                    Dim graceParkedWithoutOkCount As Integer = 0
                    SplitCmgsFlags(graceFlags, graceAdded, graceCountedConfirmed, graceParkedWithoutOk, graceParkedWithoutOkCount, CmgsLineReferences(graceText))
                    If graceAdded > 0 AndAlso Not graceCountedConfirmed Then sawCmgsWithoutOk = True
                    ParkPendingLateCmgs(graceCount - graceAdded, graceParkedWithoutOkCount)
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                        SendEscapeAndDrain(stopPredicate, 500)
                        MarkCancelled(result, sawCmgsWithoutOk)
                        Return result
                    End If
                    Continue For
                End If
                If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                    ' Stop after Ctrl-Z: drain WITHOUT a predicate so
                    ' the drain actually reads (bounded ~1s) instead
                    ' of exiting on the first stop check and dropping
                    ' a waiting +CMGS:. A late
                    ' +CMGS: here means the segment WAS sent - count
                    ' it before reporting the stop.
                    ' Capped 500ms probe: the port closes right after
                    ' this, so a full 3000ms probe only delays the
                    ' cancelled row.
                    Dim stopProbeText As String = String.Empty
                    Dim drainedAfterStop As String = SendEscapeAndDrainCore(Nothing, 500, stopProbeText, forceDrain:=True, parkRejoinedCmgs:=False)
                    Dim stopRefused As Boolean = False
                    CountGuardedCmgs(drainedAfterStop, stopProbeText, ctrlZSent, result, sawCmgsWithoutOk, stopRefused)
                    MarkCancelled(result, sawCmgsWithoutOk)
                    Return result
                End If
                Dim drainProbeText As String = String.Empty
                Dim lateDrainedText As String = SendEscapeAndDrainCore(stopPredicate, 3000, drainProbeText, forceDrain:=True, parkRejoinedCmgs:=False)
                Dim drainRefused As Boolean = False
                If CountGuardedCmgs(lateDrainedText, drainProbeText, ctrlZSent, result, sawCmgsWithoutOk, drainRefused) > 0 Then
                    ' Very-late +CMGS: arrived during the drain: the
                    ' segment WAS sent - count it, do not fail it.
                    If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                        SendEscapeAndDrain(stopPredicate, 500)
                        MarkCancelled(result, sawCmgsWithoutOk)
                        Return result
                    End If
                    ' B-B6: continue only when the drain's probe
                    ' proved command mode (see the write-Catch site
                    ' above); otherwise abort the row with the
                    ' segments-sent count, so a re-send is known to
                    ' double-send.
                    If Not _commandModeKnown Then
                        result.Success = False
                        result.Message = "تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل — لم يتم التحقق من وضع الأوامر بعد الاسترداد، فأُوقف الإرسال لتجنب إرسال مكرر — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها" & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)
                        Return result
                    End If
                    Continue For
                End If
                ' (Round 45) B-H2: a drain refusal (error-trailed
                ' +CMGS: found late) is a failed segment too - the
                ' failure message below carries the re-send clause.
                cmgsErrorTrailed = cmgsErrorTrailed OrElse drainRefused
                ' (Round 56) The modem's raw reply to the failed
                ' submit, flattened to one line, first 200 chars
                ' (like the prompt-failure branch): a timeout with
                ' no reply visible is undiagnosable — the next live
                ' failure must show whether the modem answered
                ' +CMS ERROR or nothing at all.
                ' (Round 57) C-55-L2: the ~3 s post-grace drain can
                ' hold the modem's final answer to this segment (a
                ' bare +CMS ERROR / +CME ERROR / ERROR arriving after
                ' 61.5 s of silence) — CountGuardedCmgs only examines
                ' +CMGS: lines, so without this the log wrongly says
                ' "لا يوجد رد" although the modem answered.
                ' (Round 58) C-56-L1: drainFinal mines the late drain
                ' ONLY — never the probe text. The probe is a NEW AT
                ' command sent after the timeout, so its ERROR
                ' answers the probe, not the failed submit: mining
                ' it turned an unknown fate into a false definite
                ' failure (and dropped the re-send warning). Only
                ' bytes that arrived unsolicited during the drain
                ' can carry the submit's late answer.
                ' (Round 59) C-57-L1: one exception — a +CMS ERROR /
                ' +CME ERROR in the probe window IS the submit's late
                ' refusal: plain AT is never answered with a
                ' CMGS-family error, and the only CMGS in flight is
                ' the timed-out submit (the code already trusts the
                ' same window's +CMGS: confirmations). Bare ERROR
                ' and OK from the probe stay excluded.
                ' (Round 58) C-56-L2: identical lines are shown once
                ' (BuildTimeoutReply) — the same +CMS ERROR can
                ' arrive in both the grace window and the drain
                ' window, and repeating it wastes the 200-char
                ' window.
                Dim drainFinal As String = DrainFinalResultLines(lateDrainedText) & vbCrLf & DrainFinalResultLines(drainProbeText, includeBareError:=False)
                ' (Round 76) D-L1/B-L1: a bare OK in the 3 s drain is the submit's
                ' late answer - the drain holds unsolicited bytes only (the AT
                ' probe's OK is separate in drainProbeText), so it counts exactly
                ' like an OK in the wait or grace window. Mine it into drainFinal
                ' too, so the reply suffix shows it instead of "لا يوجد رد".
                Dim drainHasOk As Boolean = ReplyHasTerminalOk(lateDrainedText)
                If drainHasOk Then drainFinal = If(String.IsNullOrWhiteSpace(drainFinal), "OK", drainFinal & vbCrLf & "OK")
                Dim timeoutReply As String = BuildTimeoutReply(response, graceText, drainFinal, lateDrainedText)
                If timeoutReply.Length > 200 Then timeoutReply = timeoutReply.Substring(0, 200)
                Dim timeoutReplySuffix As String = If(timeoutReply = String.Empty, " — رد المودم: لا يوجد رد", " — رد المودم: " & timeoutReply)
                result.Success = False
                ' (Round 76) A-L1/D-L1: the submit's fate is UNKNOWN - never a
                ' definite failure - when Ctrl-Z went out and no terminal result
                ' answers it anywhere: either the modem answered OK with no
                ' +CMGS: (wait, grace or drain) or it answered no terminal line
                ' at all, only unsolicited lines like BOOT:/^RSSI:. An error
                ' anywhere still beats an OK anywhere, and a non-OK final line
                ' (NO CARRIER, BUSY, ...) still fails definitely in the Else
                ' below. A +CMGS: line cannot reach here uncounted (it was
                ' counted above), so it can only ever keep unknownFate False.
                ' (Round 77) R76-A-L2/D-O1: DrainFinalResultLines now
                ' mines those non-OK final lines from the drain too,
                ' so the claim above holds for drain arrivals as well.
                Dim hasTerminalError As Boolean = ReplyHasTerminalError(response) OrElse ReplyHasTerminalError(graceText) OrElse ReplyHasTerminalError(drainFinal)
                Dim hasTerminalOk As Boolean = ReplyHasTerminalOk(response) OrElse ReplyHasTerminalOk(graceText) OrElse drainHasOk
                Dim hasAnyTerminalLine As Boolean = HasFinalResultLine(response) OrElse HasFinalResultLine(graceText) OrElse HasFinalResultLine(drainFinal)
                Dim unknownFate As Boolean = ctrlZSent AndAlso Not hasTerminalError AndAlso (hasTerminalOk OrElse Not hasAnyTerminalLine)
                ' (Round 57) C-55-L2: a drain final-result line IS a
                ' reply, so it takes the reply-failure template
                ' below (which keeps the suffix on a late-credit
                ' rebuild) instead of the "قد تكون الرسالة أُرسلت"
                ' template, which would contradict the error shown.
                If String.IsNullOrWhiteSpace(response) AndAlso String.IsNullOrWhiteSpace(graceText) AndAlso String.IsNullOrWhiteSpace(drainFinal) AndAlso String.IsNullOrWhiteSpace(lateDrainedText) Then
                    ' Timeout with no reply at all: the segment may
                    ' still have been sent, so say so explicitly.
                    ' FIX B-NEW-L1 (Round 27): re-send warning clause
                    ' added - the count was here, the warning was not,
                    ' unlike every sibling partial-failure message.
                    ' (Round 77) R76-A-L1: total silence is the
                    ' limiting case of the unknownFate rule below
                    ' (Ctrl-Z out + no terminal line anywhere) —
                    ' but this branch fires first, so the flag must
                    ' be set here too when Ctrl-Z went out. A row
                    ' whose Ctrl-Z never went out stays a definite
                    ' failure. (Round 77) R76-D-L1: the branch now
                    ' also requires the raw drain text to be empty —
                    ' unsolicited drain bytes are a reply, so the
                    ' flow must reach the ElseIf below instead.
                    If ctrlZSent Then result.Unconfirmed = True
                    result.Message = "انتهت المهلة — قد تكون الرسالة أُرسلت (تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " رسائل منفصلة قبل الفشل) — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها" & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty) & timeoutReplySuffix
                ElseIf unknownFate Then
                    ' (Round 75) LIVE 2026-11-23: OK without +CMGS:
                    ' Ctrl-Z went out, the modem answered OK, the
                    ' grace + drain found no +CMGS: and no error —
                    ' the segment MAY have been sent (jony received
                    ' it: reply "OK BOOT:32264700,0,0,0,87 ^RSSI:15").
                    ' Fate UNKNOWN: never a definite failure, never
                    ' counted sent — the caller reports it as
                    ' unconfirmed (see SmsResult.Unconfirmed).
                    ' (Round 76) A-L1/D-L1: the same unknown fate when
                    ' the modem answered no terminal line at all, only
                    ' unsolicited lines, or its OK arrived only in the
                    ' 3 s drain (see unknownFate above).
                    result.Unconfirmed = True
                    result.Message = "لم يتم تأكيد الإرسال — قد تكون الرسالة أُرسلت (تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " رسائل منفصلة) — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها" & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty) & timeoutReplySuffix
                Else
                    result.Message = "تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل" & If(result.SegmentsSent > 0 OrElse cmgsErrorTrailed, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty) & timeoutReplySuffix  ' (Round 33) B-observation: reply failure - modem gave a terminal reply and grace+drain found no +CMGS:, segment not sent; clause when prior Sent>0 or on an error-trailed +CMGS: refusal (Round 45 B-H2).
                End If
                Return result
            Next

            result.Success = True
            If segments.Count = 1 Then
                result.Message = If(sawCmgsWithoutOk, "تم الإرسال (بدون تأكيد OK)", "تم الإرسال")
            Else
                result.Message = If(sawCmgsWithoutOk, "تم الإرسال (" & segments.Count.ToString() & " رسائل منفصلة) (بدون تأكيد OK)", "تم الإرسال (" & segments.Count.ToString() & " رسائل منفصلة)")
            End If
        Catch ex As Exception
            Dim drainedText As String = String.Empty
            Dim outerProbeText As String = String.Empty
            Try
                drainedText = SendEscapeAndDrainCore(stopPredicate, 3000, outerProbeText, forceDrain:=True, parkRejoinedCmgs:=False)
            Catch ex2 As Exception
                System.Diagnostics.Debug.WriteLine("ESC drain failed: " & ex2.Message)
            End Try
            ' Check the drained text exactly as the write-Catch
            ' recovery above does: a very-late +CMGS: means a segment
            ' WAS sent - count it so the run totals do not undercount
            ' it. Only when a Ctrl-Z was actually sent for the
            ' current segment (the guard inside
            ' CountGuardedCmgs): an exception in the CMGF/CSCS/
            ' prompt phase, before any Ctrl-Z for this segment, must
            ' not count a stale +CMGS:. An OK counts only when it
            ' follows the +CMGS: in the same portion (drain or probe
            ' reply).
            Dim outerRefused As Boolean = False
            CountGuardedCmgs(drainedText, outerProbeText, ctrlZSent, result, sawCmgsWithoutOk, outerRefused)
            ' A stop that coincides with the exception is a stop, not
            ' a failure: honour it exactly as the write Catch does.
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                SendEscapeAndDrain(stopPredicate, 500)
                MarkCancelled(result, sawCmgsWithoutOk)
                Return result
            End If
            ' (Round 55) A port exception escaping here means the
            ' modem itself is dead (e.g. dongle unplugged while
            ' waiting): mark it so the caller stops this modem and
            ' requeues its remaining numbers instead of failing
            ' them one by one.
            result.PortDead = IsPortDeadException(ex)
            result.Success = False
            ' FIX B-L1 (Round 26): the outer-Catch message used to
            ' carry no sent-count and no re-send warning, so a
            ' failure here (e.g. dongle unplugged while
            ' WaitForResult polls segment 2 of 3) showed a bare
            ' "فشل: خطأ أثناء الإرسال: ..." and RebuildPartialMessage
            ' could not repair it - its count-refresh regex only
            ' replaces an EXISTING "تم إرسال \d+ من \d+" substring.
            ' Same shape as the sibling partial-failure messages
            ' now: count + re-send warning, ex.Message detail kept.
            ' (Round 42) A-L1: the setup phase (SegmentsTotal = 0 -
            ' the count is first assigned after setup) must report a
            ' bare error with no segment count: nothing was sent yet,
            ' and "تم إرسال 0 من 0" contradicts README_AR.txt:170.
            ' The count fragment is emitted only once a total exists;
            ' the in-send shape (count + clause keying + "(بدون تأكيد
            ' OK)" suffix) is unchanged. When Total = 0 the Ctrl-Z
            ' flags are necessarily False (Ctrl-Z writes only happen
            ' after the count is assigned) and sawCmgsWithoutOk is
            ' False (CountGuardedCmgs only sets it when ctrlZSent),
            ' so the bare message loses nothing.
            If result.SegmentsTotal > 0 Then
                result.Message = "خطأ أثناء الإرسال: " & ex.Message & " — تم إرسال " & result.SegmentsSent.ToString() & " من " & result.SegmentsTotal.ToString() & " (رسائل منفصلة) قبل الفشل" & If(result.SegmentsSent > 0 OrElse ctrlZSent OrElse ctrlZAttempted, " — إعادة الإرسال ستعيد إرسال المقاطع التي تم إرسالها", String.Empty) & If(sawCmgsWithoutOk, " (بدون تأكيد OK)", String.Empty)  ' (Round 34) B-M1: outer Catch - keep clause when prior Sent>0, or when this segment's Ctrl-Z went out (ctrlZSent) or was attempted (ctrlZAttempted) and its fate is unknown (exception cut the wait short).
            Else
                result.Message = "خطأ أثناء الإرسال: " & ex.Message
            End If
        End Try
        Return result
    End Function

    ''' <summary>Reads until a final-result line (OK / ERROR / +CMGS: /
    ''' +CMS ERROR / +CME ERROR / NO CARRIER / ...) or the timeout
    ''' elapses. Exits early on any final-result line instead of burning
    ''' the whole timeout on an unrecognised failure. +CMGS: alone is
    ''' success (OK optional). +CMGS: / ERROR are matched line-anchored
    ''' only, so message text echoed back can never false-pass or
    ''' false-fail the result. When stopPredicate is supplied it is
    ''' checked every poll (~50ms).</summary>
    Private Function WaitForResult(timeoutMs As Integer, Optional stopPredicate As Func(Of Boolean) = Nothing) As String
        Dim buffer As New System.Text.StringBuilder()
        If _port Is Nothing OrElse Not _port.IsOpen Then Return String.Empty
        ' FIX B-M1: rejoin a fragment held by an earlier Strip (e.g. "+CM"
        ' read ahead of the AT+CMGS prompt) with the bytes arriving
        ' during this wait, BEFORE any +CMGS: attribution — the fragment
        ' must not be stranded across the read.
        ' TRACED CASE (ii): fragment "+CM" held before this wait + live
        ' "GS: 12" + vbCrLf + "OK" + vbCrLf — the joined "+CMGS: 12" line is
        ' in the returned text, so the caller counts it for the current
        ' segment (bounded, last line) instead of missing the send.
        Dim heldFragment As String = TakePreservedText()
        buffer.Append(heldFragment)
        Dim watch As New System.Diagnostics.Stopwatch()
        watch.Start()
        While watch.ElapsedMilliseconds < timeoutMs
            Dim chunk As String = _port.ReadExisting()
            If Not String.IsNullOrEmpty(chunk) Then
                buffer.Append(chunk)
                Dim text As String = buffer.ToString()
                ' FIX B-D3: joined check - a live OK glued onto the
                ' held fragment ("+CMOK") still ends the wait.
                If JoinedHasFinalResultLine(text, heldFragment.Length) Then
                    ' FIX B-M1 (Round 22): the grace decision is about
                    ' the CURRENT line only - an OK confirming an
                    ' EARLIER line must not skip the grace, and an
                    ' ERROR on an earlier line must not suppress it.
                    If LineStartsWith(text, "+CMGS:") _
                        AndAlso Not CurrentCmgsConfirmed(text) _
                        AndAlso Not CurrentCmgsHasTrailingError(text) Then
                        ' +CMGS: arrived but its trailing OK has not
                        ' yet: grace-read briefly instead of leaving
                        ' the send "(بدون تأكيد OK)" on a timing artefact.
                        Dim grace As New System.Diagnostics.Stopwatch()
                        grace.Start()
                        While grace.ElapsedMilliseconds < 1000
                            Dim graceChunk As String = _port.ReadExisting()
                            If Not String.IsNullOrEmpty(graceChunk) Then
                                buffer.Append(graceChunk)
                                Dim graceText As String = buffer.ToString()
                                ' FIX B-M1 (Round 22): leave the grace
                                ' only when the CURRENT line is confirmed
                                ' or errored - see the entry check above.
                                If CurrentCmgsConfirmed(graceText) _
                                    OrElse CurrentCmgsHasTrailingError(graceText) Then Exit While
                            End If
                            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                            Threading.Thread.Sleep(50)
                        End While
                        grace.Stop()
                    End If
                    Exit While
                End If
            End If
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
            Threading.Thread.Sleep(50)
        End While
        watch.Stop()
        Dim reply As String = buffer.ToString()
        ' FIX B-M1: as in SendCommand — drop a first line that straddles
        ' the wait's start unless it is a +CMGS: line, and hold back a
        ' trailing unterminated +CMGS: fragment for the next read
        ' (GraceReadForCmgs re-seeds it and can still complete it).
        If heldFragment <> String.Empty Then reply = RejoinGluedFragmentReply(reply, heldFragment)
        Dim trailingFragment As String = String.Empty
        SplitTrailingCmgsFragment(reply, trailingFragment)
        If trailingFragment <> String.Empty Then AppendPreservedFragment(trailingFragment)
        ' (Round 45) B-M2: a bare +CMGS: never proves command mode
        ' (Round 40 B-L2 principle) - use the OtherThanCmgs helper,
        ' so a stale +CMGS: arriving during a wait whose Ctrl-Z
        ' was lost cannot set the flag and make the next
        ' SendEscapeAndDrain skip its ESC.
        If HasFinalResultLineOtherThanCmgs(reply) Then _commandModeKnown = True
        Return reply
    End Function

    ''' <summary>Brief grace read (up to ~1.5s) for a +CMGS: that may
    ''' have arrived just as a stop/timeout cut the main wait short.
    ''' Returns whatever was read (empty when nothing arrived). The
    ''' stop predicate IS polled: once a stop is seen, at most ~300ms
    ''' more grace is allowed before giving up. After +CMGS: is seen,
    ''' reading continues briefly (up to 500ms) for the trailing OK.</summary>
    Private Function GraceReadForCmgs(timeoutMs As Integer, Optional stopPredicate As Func(Of Boolean) = Nothing) As String
        Dim buffer As New System.Text.StringBuilder()
        If _port Is Nothing OrElse Not _port.IsOpen Then Return String.Empty
        ' FIX B-M1: same rejoin as WaitForResult — a fragment held for
        ' the next read (e.g. held back by WaitForResult's own return
        ' path) is completed by the bytes read here, not stranded.
        Dim heldFragment As String = TakePreservedText()
        buffer.Append(heldFragment)
        Dim watch As New System.Diagnostics.Stopwatch()
        watch.Start()
        Dim stopSeenAt As Long = -1
        While watch.ElapsedMilliseconds < timeoutMs
            Dim chunk As String = _port.ReadExisting()
            If Not String.IsNullOrEmpty(chunk) Then
                buffer.Append(chunk)
                If LineStartsWith(buffer.ToString(), "+CMGS:") Then
                    ' FIX B-R37-L1 (Round 38): the Round 37 pre-check
                    ' that skipped the 500ms OK-watch when the current
                    ' line was already confirmed is REVERTED. A lone
                    ' stale +CMGS: line in the first chunk (the previous
                    ' row's very-late confirmation) counted as the
                    ' current segment, so the current +CMGS: arriving
                    ' <500ms later was never read and was dropped as
                    ' unattributed. The watch below is unconditional
                    ' again; the Round 35 stall returns as an accepted
                    ' latency-only cost on the timeout path.
                    ' +CMGS: seen: keep reading briefly (up to 500ms)
                    ' for the trailing OK.
                    Dim okWatch As New System.Diagnostics.Stopwatch()
                    okWatch.Start()
                    While okWatch.ElapsedMilliseconds < 500
                        Dim okChunk As String = _port.ReadExisting()
                        If Not String.IsNullOrEmpty(okChunk) Then
                            buffer.Append(okChunk)
                            Dim okText As String = buffer.ToString()
                            ' FIX B-M1 (Round 22): leave the OK watch
                            ' only when the CURRENT line is confirmed
                            ' or errored - see WaitForResult above.
                            If CurrentCmgsConfirmed(okText) _
                                OrElse CurrentCmgsHasTrailingError(okText) Then Exit While
                        End If
                        If stopPredicate IsNot Nothing AndAlso stopPredicate() Then Exit While
                        Threading.Thread.Sleep(50)
                    End While
                    okWatch.Stop()
                    Exit While
                End If
            End If
            If stopPredicate IsNot Nothing AndAlso stopPredicate() Then
                If stopSeenAt < 0 Then stopSeenAt = watch.ElapsedMilliseconds
                If watch.ElapsedMilliseconds - stopSeenAt >= 300 Then Exit While
            End If
            Threading.Thread.Sleep(50)
        End While
        watch.Stop()
        Dim reply As String = buffer.ToString()
        ' FIX B-M1: as in WaitForResult — drop a straddling first line
        ' unless it is a +CMGS: line, and hold back a trailing
        ' unterminated +CMGS: fragment for the next read.
        If heldFragment <> String.Empty Then reply = RejoinGluedFragmentReply(reply, heldFragment)
        Dim trailingFragment As String = String.Empty
        SplitTrailingCmgsFragment(reply, trailingFragment)
        If trailingFragment <> String.Empty Then AppendPreservedFragment(trailingFragment)
        ' (Round 45) B-M2: same rule as the WaitForResult tail above
        ' - a reply whose only final-result evidence is a +CMGS:
        ' line is a late confirmation, not proof of command mode.
        If HasFinalResultLineOtherThanCmgs(reply) Then _commandModeKnown = True
        Return reply
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not _disposed Then
            If disposing Then Close()
            _disposed = True
        End If
    End Sub

End Class
