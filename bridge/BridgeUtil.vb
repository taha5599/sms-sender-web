Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' Small JSON + multipart helpers for SmsSenderBridge.
''' Hand-rolled on purpose: no extra dependencies, so the same code
''' compiles on .NET Framework 4.8 (the deliverable) and on the
''' .NET 10 test harness. Nothing here touches the modem.
''' </summary>
Public Module BridgeUtil

    ''' <summary>Escapes a string for use inside a JSON string literal (without the quotes).</summary>
    Public Function JsonEscape(value As String) As String
        If value Is Nothing Then Return String.Empty
        Dim sb As New StringBuilder(value.Length + 8)
        For Each ch As Char In value
            Select Case ch
                Case """"c
                    sb.Append("\""")
                Case "\"c
                    sb.Append("\\")
                Case ControlChars.Cr
                    sb.Append("\r")
                Case ControlChars.Lf
                    sb.Append("\n")
                Case ControlChars.Tab
                    sb.Append("\t")
                Case Else
                    If AscW(ch) < 32 Then
                        sb.Append("\u" & AscW(ch).ToString("X4"))
                    Else
                        sb.Append(ch)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function

    ''' <summary>Returns a JSON string literal (with quotes), or null when the value is Nothing.</summary>
    Public Function JsonStr(value As String) As String
        If value Is Nothing Then Return "null"
        Return """" & JsonEscape(value) & """"
    End Function

    Public Function JsonBool(value As Boolean) As String
        If value Then Return "true"
        Return "false"
    End Function

    ''' <summary>Reads a flat JSON string field: {"port": "COM18"} -> "COM18". Returns Nothing when absent.</summary>
    Public Function JsonGetString(json As String, key As String) As String
        If String.IsNullOrEmpty(json) Then Return Nothing
        Dim keyToken As String = """" & key & """"
        Dim keyIdx As Integer = json.IndexOf(keyToken, StringComparison.Ordinal)
        If keyIdx < 0 Then Return Nothing
        Dim colonIdx As Integer = json.IndexOf(":"c, keyIdx + keyToken.Length)
        If colonIdx < 0 Then Return Nothing
        Dim i As Integer = colonIdx + 1
        While i < json.Length AndAlso Char.IsWhiteSpace(json(i))
            i += 1
        End While
        If i >= json.Length OrElse json(i) <> """"c Then Return Nothing
        i += 1
        Dim sb As New StringBuilder()
        While i < json.Length
            Dim ch As Char = json(i)
            If ch = "\"c Then
                If i + 1 >= json.Length Then Return Nothing
                Dim esc As Char = json(i + 1)
                Select Case esc
                    Case """"c
                        sb.Append(""""c)
                    Case "\"c
                        sb.Append("\"c)
                    Case "/"c
                        sb.Append("/"c)
                    Case "n"c
                        sb.Append(ControlChars.Lf)
                    Case "r"c
                        sb.Append(ControlChars.Cr)
                    Case "t"c
                        sb.Append(ControlChars.Tab)
                    Case "b"c
                        sb.Append(ControlChars.Back)
                    Case "f"c
                        sb.Append(ControlChars.FormFeed)
                    Case "u"c
                        If i + 5 >= json.Length Then Return Nothing
                        Dim hex As String = json.Substring(i + 2, 4)
                        Dim code As Integer = 0
                        If Not Integer.TryParse(hex, Global.System.Globalization.NumberStyles.HexNumber, Global.System.Globalization.CultureInfo.InvariantCulture, code) Then Return Nothing
                        sb.Append(ChrW(code))
                        i += 4
                    Case Else
                        sb.Append(esc)
                End Select
                i += 2
            ElseIf ch = """"c Then
                Return sb.ToString()
            Else
                sb.Append(ch)
                i += 1
            End If
        End While
        Return Nothing
    End Function

    ''' <summary>Reads a flat JSON integer field. Returns defaultValue when absent or unparseable.</summary>
    Public Function JsonGetInt(json As String, key As String, defaultValue As Integer) As Integer
        If String.IsNullOrEmpty(json) Then Return defaultValue
        Dim keyToken As String = """" & key & """"
        Dim keyIdx As Integer = json.IndexOf(keyToken, StringComparison.Ordinal)
        If keyIdx < 0 Then Return defaultValue
        Dim colonIdx As Integer = json.IndexOf(":"c, keyIdx + keyToken.Length)
        If colonIdx < 0 Then Return defaultValue
        Dim i As Integer = colonIdx + 1
        While i < json.Length AndAlso Char.IsWhiteSpace(json(i))
            i += 1
        End While
        Dim start As Integer = i
        If i < json.Length AndAlso (json(i) = "-"c OrElse json(i) = "+"c) Then i += 1
        While i < json.Length AndAlso Char.IsDigit(json(i))
            i += 1
        End While
        If i = start Then Return defaultValue
        Dim value As Integer = defaultValue
        If Integer.TryParse(json.Substring(start, i - start), value) Then Return value
        Return defaultValue
    End Function

    ''' <summary>Truncates text to maxLen characters (for logs - never log a full message).</summary>
    Public Function TruncateText(value As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(value) Then Return String.Empty
        If value.Length <= maxLen Then Return value
        Return value.Substring(0, maxLen) & "…"
    End Function

    ''' <summary>Finds a byte sequence inside a byte array. Returns -1 when absent.</summary>
    Public Function FindBytes(haystack As Byte(), needle As Byte(), startIndex As Integer) As Integer
        If haystack Is Nothing OrElse needle Is Nothing OrElse needle.Length = 0 Then Return -1
        If haystack.Length < needle.Length Then Return -1
        Dim lastStart As Integer = haystack.Length - needle.Length
        For i As Integer = Math.Max(0, startIndex) To lastStart
            Dim found As Boolean = True
            For j As Integer = 0 To needle.Length - 1
                If haystack(i + j) <> needle(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then Return i
        Next
        Return -1
    End Function

    ''' <summary>One uploaded file part of a multipart/form-data body.</summary>
    Public Class MultipartFile
        Public Property FileName As String = String.Empty
        Public Property ContentType As String = String.Empty
        Public Property Data As Byte()
    End Class

    ''' <summary>Parsed multipart/form-data body: text fields + uploaded files.</summary>
    Public Class MultipartData
        Public Property Fields As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public Property Files As New Dictionary(Of String, MultipartFile)(StringComparer.OrdinalIgnoreCase)
    End Class

    Private Function MultipartHeaderValue(headers As String, name As String) As String
        Dim token As String = name & "=" & """"
        Dim idx As Integer = headers.IndexOf(token, StringComparison.OrdinalIgnoreCase)
        If idx < 0 Then Return Nothing
        Dim start As Integer = idx + token.Length
        Dim [end] As Integer = headers.IndexOf(""""c, start)
        If [end] < 0 Then Return Nothing
        Return headers.Substring(start, [end] - start)
    End Function

    ''' <summary>Parses a multipart/form-data body. Returns Nothing when the Content-Type has no boundary.</summary>
    Public Function ParseMultipart(body As Byte(), contentType As String) As MultipartData
        Dim result As New MultipartData()
        If body Is Nothing OrElse String.IsNullOrEmpty(contentType) Then Return result
        Dim boundary As String = Nothing
        Dim bIdx As Integer = contentType.IndexOf("boundary=", StringComparison.OrdinalIgnoreCase)
        If bIdx >= 0 Then
            boundary = contentType.Substring(bIdx + "boundary=".Length).Trim()
            If boundary.StartsWith("""") AndAlso boundary.EndsWith("""") AndAlso boundary.Length >= 2 Then
                boundary = boundary.Substring(1, boundary.Length - 2)
            End If
        End If
        If String.IsNullOrEmpty(boundary) Then Return Nothing
        Dim delim As Byte() = Encoding.ASCII.GetBytes("--" & boundary)
        Dim crlf As Byte() = Encoding.ASCII.GetBytes(vbCrLf)
        Dim doubleCrlf As Byte() = Encoding.ASCII.GetBytes(vbCrLf & vbCrLf)
        ' Collect delimiter positions; the first is the preamble.
        Dim positions As New List(Of Integer)()
        Dim pos As Integer = FindBytes(body, delim, 0)
        While pos >= 0
            positions.Add(pos)
            pos = FindBytes(body, delim, pos + delim.Length)
        End While
        If positions.Count < 2 Then Return result
        For p As Integer = 0 To positions.Count - 2
            Dim partStart As Integer = positions(p) + delim.Length
            ' Skip the CRLF after the delimiter line.
            If partStart + 1 < body.Length AndAlso body(partStart) = 13 AndAlso body(partStart + 1) = 10 Then
                partStart += 2
            End If
            Dim partEnd As Integer = positions(p + 1)
            ' The CRLF before the next delimiter belongs to the delimiter, not the content.
            If partEnd - 2 >= partStart AndAlso body(partEnd - 2) = 13 AndAlso body(partEnd - 1) = 10 Then
                partEnd -= 2
            End If
            If partEnd <= partStart Then Continue For
            Dim headerEnd As Integer = FindBytes(body, doubleCrlf, partStart)
            If headerEnd < 0 OrElse headerEnd >= partEnd Then Continue For
            Dim headers As String = Encoding.ASCII.GetString(body, partStart, headerEnd - partStart)
            Dim fieldName As String = MultipartHeaderValue(headers, "name")
            If String.IsNullOrEmpty(fieldName) Then Continue For
            Dim dataStart As Integer = headerEnd + doubleCrlf.Length
            Dim dataLen As Integer = partEnd - dataStart
            If dataLen < 0 Then Continue For
            Dim fileName As String = MultipartHeaderValue(headers, "filename")
            If fileName IsNot Nothing Then
                Dim f As New MultipartFile()
                f.FileName = fileName
                Dim ctIdx As Integer = headers.IndexOf("Content-Type:", StringComparison.OrdinalIgnoreCase)
                If ctIdx >= 0 Then
                    Dim lineEnd As Integer = headers.IndexOf(vbCrLf, ctIdx)
                    If lineEnd < 0 Then lineEnd = headers.Length
                    f.ContentType = headers.Substring(ctIdx + "Content-Type:".Length, lineEnd - (ctIdx + "Content-Type:".Length)).Trim()
                End If
                Dim data(dataLen - 1) As Byte
                If dataLen > 0 Then Array.Copy(body, dataStart, data, 0, dataLen)
                f.Data = data
                result.Files(fieldName) = f
            Else
                Dim textValue As String = Encoding.UTF8.GetString(body, dataStart, dataLen)
                ' A repeated text field (e.g. several "ports" fields) joins with ",".
                If result.Fields.ContainsKey(fieldName) Then
                    result.Fields(fieldName) = result.Fields(fieldName) & "," & textValue
                Else
                    result.Fields(fieldName) = textValue
                End If
            End If
        Next
        Return result
    End Function

End Module
