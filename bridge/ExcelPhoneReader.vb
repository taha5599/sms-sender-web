Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports System.Xml.Linq

''' <summary>
''' Reads phone numbers from Excel / CSV files.
''' NO Excel Interop, NO NuGet, NO Excel install required.
''' .xlsx / .xlsm are read as a ZIP: xl/sharedStrings.xml + the first
''' worksheet resolved through xl/workbook.xml and its relationships.
''' First sheet only. Numbers only - no name column is expected or used.
''' </summary>
Public Class ExcelPhoneReader

    Public Class LoadResult
        Public Property TotalFound As Integer
        Public Property ValidUnique As New List(Of String)()
        Public Property DuplicatesRemoved As Integer
        Public Property InvalidSkipped As Integer
        Public Property SheetCount As Integer
        Public Property SheetName As String = String.Empty
        Public Property ErrorMessage As String = String.Empty
        Public ReadOnly Property HasError As Boolean
            Get
                Return Not String.IsNullOrEmpty(ErrorMessage)
            End Get
        End Property
    End Class

    Private _lastSheetCount As Integer
    Private _lastSheetName As String = String.Empty
    Private _skippedInvalidCells As Integer
    Private _allSheetsHidden As Boolean
    Private _csvUnclosedQuote As Boolean
    Private _dateStyleIndexes As New HashSet(Of Integer)()

    Public Function LoadPhones(filePath As String) As LoadResult
        Dim result As New LoadResult()
        _skippedInvalidCells = 0
        _csvUnclosedQuote = False
        If String.IsNullOrEmpty(filePath) Then
            result.ErrorMessage = "لم يتم تحديد ملف."
            Return result
        End If
        Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()

        If ext = ".xls" Then
            result.ErrorMessage = "هذا الملف بصيغة .xls القديمة." & vbCrLf &
                "من فضلك افتحه في Excel ثم احفظه بصيغة .xlsx (Save As -> Excel Workbook) وحاول مرة أخرى."
            Return result
        End If

        Dim rawValues As New List(Of String)()
        Try
            If ext = ".xlsx" OrElse ext = ".xlsm" Then
                rawValues = ReadXlsxValues(filePath)
                result.SheetCount = _lastSheetCount
                result.SheetName = _lastSheetName
            ElseIf ext = ".csv" Then
                rawValues = ReadCsvValues(filePath)
                result.SheetCount = 1
            Else
                result.ErrorMessage = "صيغة الملف غير مدعومة. استخدم ملف .xlsx أو .xlsm أو .csv"
                Return result
            End If
        Catch ex As Exception
            result.ErrorMessage = "تعذر قراءة الملف: " & ex.Message
            Return result
        End Try

        If _csvUnclosedQuote Then
            result.ErrorMessage = "ملف CSV غير مكتمل: علامة اقتباس غير مغلقة."
            Return result
        End If

        Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
        Dim seenInvalid As New HashSet(Of String)(StringComparer.Ordinal)
        For Each rawValue As String In rawValues
            For Each candidate As String In PhoneNormalizer.ExtractPhoneCandidates(rawValue)
                result.TotalFound += 1
                Dim normalized As String = String.Empty
                If PhoneNormalizer.NormalizePhone(candidate, normalized) Then
                    If seen.Contains(normalized) Then
                        result.DuplicatesRemoved += 1
                    Else
                        seen.Add(normalized)
                        result.ValidUnique.Add(normalized)
                    End If
                Else
                    ' Dedupe invalid candidates too (by their digits, so
                    ' "0101 234" and "0101-234" count once): a repeated
                    ' bad value counts as a duplicate, so the invariant
                    ' TotalFound = Valid + Duplicates + Invalid holds.
                    Dim invalidKey As String = InvalidCandidateKey(candidate)
                    If seenInvalid.Contains(invalidKey) Then
                        result.DuplicatesRemoved += 1
                    Else
                        seenInvalid.Add(invalidKey)
                        result.InvalidSkipped += 1
                    End If
                End If
            Next
        Next
        ' Cells skipped for an out-of-range shared-string index (M11)
        ' count as invalid as well - and they increment TotalFound too,
        ' so TotalFound = Valid + Duplicates + Invalid always holds.
        result.TotalFound += _skippedInvalidCells
        result.InvalidSkipped += _skippedInvalidCells

        Return result
    End Function

    ''' <summary>
    ''' Dedupe key for an invalid candidate: its ASCII digits only, so
    ''' the same bad number written with different separators - or with
    ''' and without a leading "+" - counts once.
    ''' </summary>
    Private Function InvalidCandidateKey(candidate As String) As String
        Dim sb As New System.Text.StringBuilder()
        For Each ch As Char In candidate
            If ch >= "0"c AndAlso ch <= "9"c Then sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Reads every cell value from the FIRST VISIBLE worksheet of an
    ''' .xlsx/.xlsm file. The first visible worksheet is found through
    ''' xl/workbook.xml (first visible &lt;sheet&gt;) resolved via
    ''' xl/_rels/workbook.xml.rels; sheet1.xml is only a last-resort
    ''' fallback. Hidden sheets (state="hidden"/"veryHidden") are skipped;
    ''' when every sheet is hidden the load fails with an error. Handles
    ''' shared strings, inline strings and plain numbers. Sets
    ''' _lastSheetCount to the workbook's sheet count, or to -1 when the
    ''' first sheet could not be determined (resolver failed, or the
    ''' fallback ran because the resolved part was missing) so the UI can
    ''' warn the user instead of assuming sheet 1; _lastSheetName is the
    ''' first visible sheet's name, or empty when the sheet actually read
    ''' is unknown (fallback used). Numeric cells whose style is a date
    ''' format are skipped: they hold a date serial, not a phone number.
    ''' </summary>
    Private Function ReadXlsxValues(filePath As String) As List(Of String)
        Dim values As New List(Of String)()
        Dim sharedStrings As New List(Of String)()
        _skippedInvalidCells = 0
        _allSheetsHidden = False
        _lastSheetName = String.Empty

        Using archive As ZipArchive = ZipFile.OpenRead(filePath)
            Dim sharedEntry As ZipArchiveEntry = archive.GetEntry("xl/sharedStrings.xml")
            If sharedEntry IsNot Nothing Then
                Using stream As Stream = sharedEntry.Open()
                    Dim doc As XDocument = XDocument.Load(stream)
                    Dim ns As XNamespace = doc.Root.Name.Namespace
                    For Each si As XElement In doc.Root.Elements(ns + "si")
                        sharedStrings.Add(ExtractRichText(si, ns))
                    Next
                End Using
            End If

            _dateStyleIndexes = LoadDateStyleIndexes(archive)

            Dim sheetCount As Integer = 0
            Dim sheetPart As String = ResolveFirstSheetPart(archive, sheetCount)
            Dim resolverFailed As Boolean = String.IsNullOrEmpty(sheetPart)
            If _allSheetsHidden Then
                Throw New InvalidDataException("كل أوراق ملف الإكسيل مخفية، ولا توجد ورقة ظاهرة للقراءة.")
            End If
            _lastSheetCount = sheetCount

            Dim sheetEntry As ZipArchiveEntry = Nothing
            Dim fallbackUsed As Boolean = False
            If Not String.IsNullOrEmpty(sheetPart) Then
                sheetEntry = archive.GetEntry(sheetPart)
            End If
            If sheetEntry Is Nothing Then
                ' Last-resort fallback only: prefer sheet1.xml, otherwise
                ' the first worksheet entry in the archive.
                fallbackUsed = True
                sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")
            End If
            If sheetEntry Is Nothing Then
                For Each entry As ZipArchiveEntry In archive.Entries
                    If entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) _
                        AndAlso entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) Then
                        sheetEntry = entry
                        Exit For
                    End If
                Next
            End If
            If sheetEntry Is Nothing Then Throw New InvalidDataException("لم يتم العثور على أي ورقة داخل ملف الإكسيل.")
            ' CONTRACT (M10/M1): when the first sheet could not be
            ' determined (resolver failed, or the fallback ran because
            ' the resolved part was missing), SheetCount is -1 - never
            ' silently asserted to 1 - so the UI can warn the user. The
            ' sheet actually read is then unknown too, so its name is
            ' cleared as well.
            If resolverFailed OrElse fallbackUsed Then
                _lastSheetCount = -1
                _lastSheetName = String.Empty
            End If

            Using stream As Stream = sheetEntry.Open()
                Dim doc As XDocument = XDocument.Load(stream)
                Dim ns As XNamespace = doc.Root.Name.Namespace
                For Each cell As XElement In doc.Descendants(ns + "c")
                    Dim cellValue As String = GetCellValue(cell, ns, sharedStrings)
                    If Not String.IsNullOrEmpty(cellValue) Then values.Add(cellValue)
                Next
            End Using
        End Using

        ' Only fail the file when no usable cells remain at all: a file
        ' whose every cell was skipped for a bad shared-string index is
        ' unreadable, not empty.
        If values.Count = 0 AndAlso _skippedInvalidCells > 0 Then
            Throw New InvalidDataException("تعذر قراءة خلايا ملف الإكسيل: كل فهارس النصوص المشتركة خارج النطاق.")
        End If

        Return values
    End Function

    ''' <summary>
    ''' Returns the part name (e.g. "xl/worksheets/sheet2.xml") of the
    ''' FIRST VISIBLE WORKSHEET in xl/workbook.xml, resolved through its
    ''' relationship id, and outputs the workbook's total sheet count.
    ''' Sheets with state="hidden" or "veryHidden" are skipped, and so
    ''' are sheets whose relationship is not a worksheet (e.g. a
    ''' chartsheet — FIX C-L2, Round 14). Returns an
    ''' empty string when workbook.xml / the relationships cannot be read:
    ''' any parse failure is caught here, never thrown, so the caller can
    ''' fall back to sheet1.xml (M9). When every sheet is hidden,
    ''' _allSheetsHidden is set and an empty string is returned (L3).
    ''' FIX C-M1 (Round 17): a visible sheet whose relationship cannot
    ''' be resolved (missing r:id / missing relationship / missing
    ''' Target) also returns empty - the caller then takes the -1
    ''' "alternative sheet read" warning path instead of silently
    ''' loading the NEXT sheet's numbers. Only a relationship that
    ''' POSITIVELY identifies a non-worksheet (chartsheet /
    ''' dialogsheet) is skipped for the next sheet.
    ''' On success _lastSheetName is set to the first visible worksheet's name.
    ''' </summary>
    Private Function ResolveFirstSheetPart(archive As ZipArchive, ByRef sheetCount As Integer) As String
        sheetCount = 0
        _allSheetsHidden = False
        Try
            Dim workbookEntry As ZipArchiveEntry = archive.GetEntry("xl/workbook.xml")
            If workbookEntry Is Nothing Then Return String.Empty
            Using stream As Stream = workbookEntry.Open()
                Dim doc As XDocument = XDocument.Load(stream)
                Dim ns As XNamespace = doc.Root.Name.Namespace
                Dim sheetsElement As XElement = doc.Root.Element(ns + "sheets")
                If sheetsElement Is Nothing Then Return String.Empty
                Dim sheetElements As List(Of XElement) = sheetsElement.Elements(ns + "sheet").ToList()
                sheetCount = sheetElements.Count
                If sheetCount = 0 Then Return String.Empty
                ' FIX C-M1 (Round 17): determine visibility HERE, before
                ' the .rels early return below - an all-hidden workbook
                ' with a missing .rels must fail with the "all sheets
                ' hidden" error (via _allSheetsHidden), not fall back
                ' to reading hidden sheet1.xml with only the -1
                ' warning. Reused at the end of the loop (see below),
                ' so it is computed once.
                Dim anyVisible As Boolean = False
                For Each sheetElement0 As XElement In sheetElements
                    Dim stateAttr0 As XAttribute = sheetElement0.Attribute("state")
                    Dim state0 As String = String.Empty
                    If stateAttr0 IsNot Nothing Then state0 = stateAttr0.Value
                    If Not String.Equals(state0, "hidden", StringComparison.OrdinalIgnoreCase) AndAlso
                        Not String.Equals(state0, "veryHidden", StringComparison.OrdinalIgnoreCase) Then
                        anyVisible = True
                        Exit For
                    End If
                Next
                If Not anyVisible Then _allSheetsHidden = True
                ' FIX C-L2 (Round 14): the relationships are read up
                ' front, so each visible sheet's relationship TYPE can
                ' be checked before its target is used: a chartsheet
                ' (Type ".../relationships/chartsheet", target under
                ' xl/chartsheets/) is skipped and the next visible
                ' sheet is tried, instead of resolving to a part with
                ' no <c> elements and loading zero numbers silently.
                Dim relsEntry As ZipArchiveEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
                If relsEntry Is Nothing Then Return String.Empty
                Dim rels As New Dictionary(Of String, XElement)(StringComparer.OrdinalIgnoreCase)
                Using relsStream As Stream = relsEntry.Open()
                    Dim relsDoc As XDocument = XDocument.Load(relsStream)
                    Dim relsNs As XNamespace = relsDoc.Root.Name.Namespace
                    For Each rel As XElement In relsDoc.Root.Elements(relsNs + "Relationship")
                        Dim idAttr As XAttribute = rel.Attribute("Id")
                        If idAttr IsNot Nothing AndAlso Not rels.ContainsKey(idAttr.Value) Then rels.Add(idAttr.Value, rel)
                    Next
                End Using
                For Each sheetElement As XElement In sheetElements
                    Dim stateAttr As XAttribute = sheetElement.Attribute("state")
                    Dim state As String = String.Empty
                    If stateAttr IsNot Nothing Then state = stateAttr.Value
                    If String.Equals(state, "hidden", StringComparison.OrdinalIgnoreCase) OrElse
                        String.Equals(state, "veryHidden", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If
                    Dim relId As String = String.Empty
                    Dim relAttr As XAttribute = sheetElement.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                    If relAttr IsNot Nothing Then relId = relAttr.Value
                    ' FIX C-M1 (Round 17): a visible sheet whose
                    ' relationship cannot be resolved (missing r:id
                    ' here, missing relationship or missing Target
                    ' below) must NOT be silently skipped for the
                    ' next sheet - the caller would then read the
                    ' next sheet's numbers while asserting the first
                    ' visible sheet was read. Return empty so the
                    ' load takes the -1 "alternative sheet read"
                    ' warning path. Only a relationship that
                    ' POSITIVELY identifies a non-worksheet (a
                    ' chartsheet/dialogsheet, see isWorksheet below)
                    ' is skipped for the next sheet.
                    If String.IsNullOrEmpty(relId) Then Return String.Empty
                    Dim rel As XElement = Nothing
                    If Not rels.TryGetValue(relId, rel) Then Return String.Empty
                    Dim targetAttr As XAttribute = rel.Attribute("Target")
                    If targetAttr Is Nothing Then Return String.Empty
                    Dim target As String = targetAttr.Value.Replace("\", "/")
                    If target.StartsWith("/", StringComparison.Ordinal) Then target = target.TrimStart("/"c)
                    If Not target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) Then target = "xl/" & target
                    ' FIX C-L2: only a worksheet relationship resolves
                    ' here — a chartsheet (or any other part type) is
                    ' skipped for the next visible sheet. The target
                    ' check is the fallback for a missing Type: a part
                    ' under xl/worksheets/ is a worksheet.
                    Dim relType As String = String.Empty
                    Dim typeAttr As XAttribute = rel.Attribute("Type")
                    If typeAttr IsNot Nothing Then relType = typeAttr.Value
                    Dim isWorksheet As Boolean = String.Equals(relType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", StringComparison.OrdinalIgnoreCase) _
                        OrElse target.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                    If Not isWorksheet Then Continue For
                    Dim nameAttr As XAttribute = sheetElement.Attribute("name")
                    If nameAttr IsNot Nothing Then _lastSheetName = nameAttr.Value
                    Return target
                Next
                ' No visible sheet resolved to a worksheet: when every
                ' sheet was hidden, flag it (L3) so the caller fails
                ' with an error instead of falling back silently;
                ' otherwise (e.g. every visible sheet was a chartsheet)
                ' return empty so the caller falls back to sheet1.xml
                ' as before.
                ' anyVisible was computed up front (FIX C-M1, Round
                ' 17) so the .rels early return above cannot bypass
                ' it; it is only re-asserted here.
                If Not anyVisible Then _allSheetsHidden = True
                Return String.Empty
            End Using
        Catch ex As Exception
            ' Resolver failure must never abort the load: return empty so
            ' the caller falls back to sheet1.xml. The first sheet is then
            ' undetermined, which the caller reports as SheetCount = -1.
            sheetCount = 0
            Return String.Empty
        End Try
        Return String.Empty
    End Function

    ''' <summary>
    ''' Concatenates the text runs of a shared/inline string element:
    ''' direct &lt;t&gt; children plus &lt;r&gt;/&lt;t&gt; runs only.
    ''' Phonetic runs (&lt;rPh&gt;/&lt;t&gt;) are deliberately excluded.
    ''' </summary>
    Private Function ExtractRichText(container As XElement, ns As XNamespace) As String
        Dim textParts As New System.Text.StringBuilder()
        For Each t As XElement In container.Elements(ns + "t")
            textParts.Append(t.Value)
        Next
        For Each r As XElement In container.Elements(ns + "r")
            Dim t As XElement = r.Element(ns + "t")
            If t IsNot Nothing Then textParts.Append(t.Value)
        Next
        Return textParts.ToString()
    End Function

    Private Function GetCellValue(cell As XElement, ns As XNamespace, sharedStrings As List(Of String)) As String
        Dim cellType As String = String.Empty
        Dim typeAttr As XAttribute = cell.Attribute("t")
        If typeAttr IsNot Nothing Then cellType = typeAttr.Value

        If cellType = "inlineStr" Then
            Dim isElement As XElement = cell.Element(ns + "is")
            If isElement Is Nothing Then Return String.Empty
            Return ConvertScientificIfIntegral(ExtractRichText(isElement, ns))
        End If

        ' L4: boolean cells (t="b") and error cells (t="e", e.g.
        ' #DIV/0!) hold no phone number - return empty like an empty
        ' cell, instead of letting the "0" in "#DIV/0!" count invalid.
        If cellType = "b" OrElse cellType = "e" Then Return String.Empty

        Dim valueElement As XElement = cell.Element(ns + "v")
        If valueElement Is Nothing Then Return String.Empty
        Dim raw As String = valueElement.Value
        If String.IsNullOrEmpty(raw) Then Return String.Empty

        If cellType = "s" Then
            Dim index As Integer = 0
            If Integer.TryParse(raw, index) AndAlso index >= 0 AndAlso index < sharedStrings.Count Then
                Return ConvertScientificIfIntegral(sharedStrings(index))
            End If
            ' One bad shared-string index must not abort the whole file:
            ' skip this cell, count it invalid, keep reading (M11).
            _skippedInvalidCells += 1
            Return String.Empty
        End If

        ' FIX C-D2: the date-style skip moved inside the plain-number
        ' gate below — it must apply only to ""/"n" cells (see there).
        ' A t="str" formula result is TEXT, never a date serial, so a
        ' date-formatted t="str" cell must load normally here.

        ' Plain number. Excel may store large numbers in scientific notation.
        ' Convert to plain digits only when the value has no fractional part,
        ' so 1012345678.5 is never rounded into a phantom candidate.
        ' D7: coerce ONLY a plain-number cell (no t attribute — a
        ' t="str" formula result is text, not a number) and with
        ' NumberStyles.Float, not Any: Any also accepts thousands
        ' separators, parentheses-negatives and currency, so junk
        ' text like "1,012,345,678" would coerce into a valid phone.
        ' C-M1: OOXML ST_CellType defines "n" as number (the
        ' default); some producers write t="n" explicitly, and
        ' such a cell must take the same Float coercion as an
        ' untyped one — otherwise "1.012345678E+9" splits at
        ' E+ into a valid part plus a phantom invalid "9".
        ' "s"/"inlineStr"/"b"/"e"/"str" still stay out: only
        ' "" and "n" are plain numbers.
        If String.IsNullOrEmpty(cellType) OrElse cellType = "n" Then
            ' FIX C-D2: date-style skip applies ONLY to plain-number
            ' cells ("" or "n"). A t="str" value is text and can never
            ' be a date serial, so it is no longer skipped; t="d" cells
            ' are not skipped either — they return an ISO date string
            ' the extractor already drops as date+time either way.
            ' M6: a plain number whose cell style is a date format is an
            ' Excel date serial (e.g. 45977 for 2026-11-16), not a phone
            ' number: skip the cell entirely instead of counting it invalid.
            ' D6: the same holds for a pure-time format — a time serial
            ' (e.g. 0.4375 for 10:30) is not a phone number either.
            Dim styleAttr As XAttribute = cell.Attribute("s")
            If styleAttr IsNot Nothing AndAlso _dateStyleIndexes IsNot Nothing Then
                Dim styleIndex As Integer = 0
                If Integer.TryParse(styleAttr.Value, styleIndex) AndAlso _dateStyleIndexes.Contains(styleIndex) Then
                    Return String.Empty
                End If
            End If
            Dim numberValue As Decimal = 0D
            If Decimal.TryParse(raw, Globalization.NumberStyles.Float,
                                Globalization.CultureInfo.InvariantCulture, numberValue) Then
                If numberValue = Decimal.Truncate(numberValue) Then
                    ' FIX C-L1 (Round 23): do NOT Decimal-convert an
                    ' integral raw value whose digits have a leading
                    ' zero followed by more digits ("001012345678"): the
                    ' conversion strips the zeros and invents a valid
                    ' phone ("1012345678"), while the raw digits match
                    ' no valid branch and are invalid. Return the raw
                    ' text instead so the extractor rejects it. Plain
                    ' integrals without a leading zero, "0" alone and
                    ' non-integral values are unchanged; the scientific
                    ' conversion path (ConvertScientificIfIntegral) is
                    ' untouched.
                    Dim trimmedRaw As String = raw.Trim()
                    ' FIX C-L4 (Round 24): apply the leading-zero
                    ' test AFTER stripping an optional leading
                    ' "+"/"-" - Decimal strips the sign and the
                    ' zeros, so "<v>+001012345678</v>" converted
                    ' to "1012345678" and loaded as a valid phone.
                    ' This aligns the numeric path with the text
                    ' path (shared string "+01012345678" was
                    ' already 1 invalid); a signed value without
                    ' a leading zero ("+1012345678") still
                    ' converts and validates as before.
                    Dim signStripped As String = trimmedRaw
                    If signStripped.StartsWith("+") OrElse signStripped.StartsWith("-") Then signStripped = signStripped.Substring(1)
                    If signStripped.Length > 1 AndAlso signStripped(0) = "0"c AndAlso Char.IsDigit(signStripped(1)) Then
                        Return raw
                    End If
                    Return numberValue.ToString("0", Globalization.CultureInfo.InvariantCulture)
                End If
                Return raw
            End If
        End If
        ' FIX C-D1: a t="str" formula result in scientific notation
        ' ("1.012345678E+9") must be converted to plain digits here, or
        ' the extractor splits it at E+ into a valid part plus a phantom
        ' invalid "9". ConvertScientificIfIntegral only fires when an
        ' E/e marker is present and the value is integral, so it cannot
        ' coerce "1,012,345,678" or ISO dates, and it does not violate
        ' D7's "text, not a number" rule — the same text conversion
        ' already used for shared/inline strings.
        Return ConvertScientificIfIntegral(raw)
    End Function

    ''' <summary>
    ''' Converts a scientific-notation value (e.g. "1.012345678E+9" stored
    ''' as text in a shared/inline string) to plain digits, but only when
    ''' it has no fractional part, so a leading zero on a plain digit
    ''' string is never lost and "1012345678.5" is never rounded into a
    ''' phantom candidate (L7). Plain digit strings are returned as-is.
    ''' </summary>
    Private Function ConvertScientificIfIntegral(value As String) As String
        If String.IsNullOrEmpty(value) Then Return value
        ' L6: only convert when the value actually contains an exponent
        ' marker (E/e) - a bare integral decimal such as "1.0" or
        ' "10.0" is left as-is and is NOT rewritten to "1" / "10".
        If value.IndexOf("E"c) < 0 AndAlso value.IndexOf("e"c) < 0 Then Return value
        Dim numberValue As Decimal = 0D
        If Decimal.TryParse(value, Globalization.NumberStyles.Float,
                            Globalization.CultureInfo.InvariantCulture, numberValue) _
            AndAlso numberValue = Decimal.Truncate(numberValue) Then
            Return numberValue.ToString("0", Globalization.CultureInfo.InvariantCulture)
        End If
        ' FIX C-M3 (Round 18): a SPACED scientific value ("1.012345678
        ' E+9") CAN reach this function: shared strings (t="s") and
        ' inline strings hold verbatim user-typed text - ExtractRichText
        ' does not trim - and a t="str" formula result can hold spaces
        ' too. Decimal.TryParse rejects an internal space, so when the
        ' direct parse above fails, match the spaced shape (same pattern
        ' as CloseCandidate in PhoneNormalizer) and strip the spaces
        ' before parsing. Strictly additive: anything that parsed above
        ' is returned before this branch runs. The plain-number path
        ' above needs no such handling: Excel writes a numeric <v>
        ' value without internal spaces.
        Dim trimmedValue As String = value.Trim()
        ' FIX C-M3 (Round 20): faithful port of the extractor's M4
        ' gate (PhoneNormalizer ExtractFieldCandidates) — the Round 19
        ' gate did NOT mirror it, and overreached in two ways:
        ' "[Ee]\s*[+-]" let spaces sit between E and the sign (the
        ' extractor requires the sign IMMEDIATELY after E), and
        ' Contains(".") fired regardless of gap shape (the extractor
        ' AND-s the dot test with the gap-shape test). So spaced
        ' signed forms the extractor would split ("01012345678 E + 9",
        ' "01000000000 E -9") were converted here into one huge number
        ' and lost a valid phone on the .xlsx path only. Parse the
        ' spaced value into mantissa / gap text / exponent group — the
        ' group starts at "+" when "+" immediately precedes the
        ' exponent digits, else at the first exponent digit — then
        ' apply the extractor's OWN gap regex to the gap text, AND-ed
        ' with (mantissa contains "." OR gap contains "+" / "-" OR the
        ' group starts with "+"): exactly the extractor's condition
        ' with the same pieces. The direct TryParse path above stays
        ' ungated: the adjacent form ("01012345678E9") is genuinely
        ' unambiguous scientific syntax and its conversion matches the
        ' extractor's adjacent-form handling.
        ' FIX C-Low-1 (Round 38): SPACE-only gaps in this branch -
        ' "\s*" also matched tab/newline, but the extractor's field
        ' split makes \t/\r/\n STRONG delimiters, so it can never
        ' see (or merge) a tab/newline-spaced exponent - converting
        ' one here invented or destroyed candidates the extractor
        ' never produces. "[ ]*" matches exactly what the extractor
        ' can see; the mantissa backtrack and gap regex below use
        ' " " the same way.
        If Regex.IsMatch(trimmedValue, "^\+?\d+(?:\.\d+)?[ ]*[Ee][ ]*[+-]?[ ]*\d+$") Then
            Dim ePos As Integer = trimmedValue.IndexOf("E"c)
            If ePos < 0 Then ePos = trimmedValue.IndexOf("e"c)
            Dim mantissaEnd As Integer = ePos
            While mantissaEnd > 0 AndAlso trimmedValue(mantissaEnd - 1) = " "c ' FIX C-Low-1 (Round 38): space-only, as the pattern above.
                mantissaEnd -= 1
            End While
            Dim mantissa As String = trimmedValue.Substring(0, mantissaEnd)
            Dim exponentDigitStart As Integer = trimmedValue.Length - 1
            While exponentDigitStart >= 0 AndAlso Char.IsDigit(trimmedValue(exponentDigitStart))
                exponentDigitStart -= 1
            End While
            exponentDigitStart += 1
            Dim exponentGroupStart As Integer = exponentDigitStart
            If exponentGroupStart > 0 AndAlso trimmedValue(exponentGroupStart - 1) = "+"c Then exponentGroupStart -= 1
            Dim gapText As String = trimmedValue.Substring(mantissaEnd, exponentGroupStart - mantissaEnd)
            Dim isExponentGap As Boolean = Regex.IsMatch(gapText, "^ *[Ee][+-]? *$") AndAlso (mantissa.Contains(".") OrElse gapText.Contains("+") OrElse gapText.Contains("-") OrElse exponentGroupStart < exponentDigitStart) ' FIX C-Low-1 (Round 38): space-only gap, as the pattern above.
            If isExponentGap Then
                Dim compactValue As String = Regex.Replace(trimmedValue, "\s", String.Empty)
                If Decimal.TryParse(compactValue, Globalization.NumberStyles.Float,
                                    Globalization.CultureInfo.InvariantCulture, numberValue) _
                    AndAlso numberValue = Decimal.Truncate(numberValue) Then
                    Return numberValue.ToString("0", Globalization.CultureInfo.InvariantCulture)
                End If
            End If
        End If
        Return value
    End Function

    ''' <summary>
    ''' Returns the set of cell-style indexes (positions in xl/styles.xml
    ''' cellXfs) whose number format is a date format — or a pure-time
    ''' format (D6): built-in date ids 14-22, 27-36, 45-47, 50-58
    ''' (those ranges already cover the built-in time ids), or a custom
    ''' numFmt whose format code contains y/m/d (or is a pure-time code,
    ''' see IsTimeFormatCode). Numeric cells with one of these styles hold
    ''' an Excel date/time serial, not a phone number. Returns an empty set
    ''' when styles.xml is absent or unreadable: the load then behaves
    ''' exactly as before instead of failing (M6).
    ''' </summary>
    Private Function LoadDateStyleIndexes(archive As ZipArchive) As HashSet(Of Integer)
        Dim dateStyles As New HashSet(Of Integer)()
        Try
            Dim stylesEntry As ZipArchiveEntry = archive.GetEntry("xl/styles.xml")
            If stylesEntry Is Nothing Then Return dateStyles
            Using stream As Stream = stylesEntry.Open()
                Dim doc As XDocument = XDocument.Load(stream)
                Dim ns As XNamespace = doc.Root.Name.Namespace
                Dim customDateFormats As New HashSet(Of Integer)()
                Dim numFmts As XElement = doc.Root.Element(ns + "numFmts")
                If numFmts IsNot Nothing Then
                    For Each fmt As XElement In numFmts.Elements(ns + "numFmt")
                        Dim idAttr As XAttribute = fmt.Attribute("numFmtId")
                        Dim codeAttr As XAttribute = fmt.Attribute("formatCode")
                        If idAttr Is Nothing OrElse codeAttr Is Nothing Then Continue For
                        Dim fmtId As Integer = 0
                        If Not Integer.TryParse(idAttr.Value, fmtId) Then Continue For
                        If IsDateFormatCode(codeAttr.Value) OrElse IsTimeFormatCode(codeAttr.Value) Then customDateFormats.Add(fmtId)
                    Next
                End If
                Dim cellXfs As XElement = doc.Root.Element(ns + "cellXfs")
                If cellXfs Is Nothing Then Return dateStyles
                Dim index As Integer = 0
                For Each xf As XElement In cellXfs.Elements(ns + "xf")
                    Dim fmtAttr As XAttribute = xf.Attribute("numFmtId")
                    Dim fmtId As Integer = 0
                    If fmtAttr IsNot Nothing AndAlso Integer.TryParse(fmtAttr.Value, fmtId) Then
                        If IsBuiltinDateFormatId(fmtId) OrElse customDateFormats.Contains(fmtId) Then
                            dateStyles.Add(index)
                        End If
                    End If
                    index += 1
                Next
            End Using
        Catch ex As Exception
            ' styles.xml absent/unreadable: leave the set empty so the
            ' load behaves as before instead of failing.
            Return New HashSet(Of Integer)()
        End Try
        Return dateStyles
    End Function

    Private Function IsBuiltinDateFormatId(fmtId As Integer) As Boolean
        Return (fmtId >= 14 AndAlso fmtId <= 22) OrElse
            (fmtId >= 27 AndAlso fmtId <= 36) OrElse
            (fmtId >= 45 AndAlso fmtId <= 47) OrElse
            (fmtId >= 50 AndAlso fmtId <= 58)
    End Function

    Private Function IsDateFormatCode(formatCode As String) As Boolean
        If String.IsNullOrEmpty(formatCode) Then Return False
        ' M1: ignore quoted literals, bracketed colour/condition
        ' sections such as [Red] (whose letters are not date tokens -
        ' "#,##0;[Red](#,##0)" must NOT count as a date format), and
        ' backslash escapes (\x), before testing for date tokens.
        ' Then test for date tokens below: a lone "d" / "m" counts as
        ' a date token, but a pure time pattern ("h:mm", "mm:ss")
        ' with no y/d token does not (L4) — D6 catches those in
        ' IsTimeFormatCode instead, so time serials are skipped too.
        Dim code As String = StripFormatCodeLiterals(formatCode).ToLowerInvariant()
        ' L4: any year (y) or day (d) token makes it a date format,
        ' even alongside a time ("m/d/yy h:mm" is a datetime serial).
        If code.Contains("y"c) OrElse code.Contains("d"c) Then Return True
        ' No y/d token: a colon means a pure time format ("h:mm",
        ' "mm:ss", "[h]:mm:ss" - the "[h]" was stripped above, leaving
        ' ":mm:ss"), never a date.
        If code.Contains(":"c) Then Return False
        ' An hour token alone ("h", "hh") is a pure time format too.
        If code.Contains("h"c) Then Return False
        ' A lone month token ("m", "mm", "mmm") with no time context
        ' is a date format (month number/name of a date serial).
        If code.Contains("m"c) Then Return True
        Return False
    End Function

    ''' <summary>
    ''' D6: strips quoted literals, bracketed colour/condition
    ''' sections and backslash escapes from an Excel number-format
    ''' code — the shared pre-step of IsDateFormatCode and
    ''' IsTimeFormatCode (factored out of IsDateFormatCode, behaviour
    ''' unchanged).
    ''' </summary>
    Private Function StripFormatCodeLiterals(formatCode As String) As String
        Dim stripped As New System.Text.StringBuilder()
        Dim inQuote As Boolean = False
        Dim i As Integer = 0
        While i < formatCode.Length
            Dim ch As Char = formatCode(i)
            If inQuote Then
                If ch = """"c Then inQuote = False
                i += 1
            ElseIf ch = """"c Then
                inQuote = True
                i += 1
            ElseIf ch = "["c Then
                ' Skip to the matching "]" (colour / condition section).
                Dim close As Integer = formatCode.IndexOf("]"c, i)
                If close < 0 Then Exit While
                i = close + 1
            ElseIf ch = "\"c Then
                ' Backslash escape: skip the escaped character too.
                i += 2
            Else
                stripped.Append(ch)
                i += 1
            End If
        End While
        Return stripped.ToString()
    End Function

    ''' <summary>
    ''' FIX C-D3: strips ONLY quoted literals ("...") from an Excel
    ''' number-format code, keeping bracketed sections ([h], [Red], ...)
    ''' and everything else. Used by IsTimeFormatCode to test bracketed
    ''' elapsed tokens: a quoted "[h]" is a literal, not an elapsed
    ''' token, while a real unquoted [h] survives the strip.
    ''' </summary>
    Private Function StripFormatCodeQuotes(formatCode As String) As String
        Dim stripped As New System.Text.StringBuilder()
        Dim inQuote As Boolean = False
        Dim i As Integer = 0
        While i < formatCode.Length
            Dim ch As Char = formatCode(i)
            If inQuote Then
                If ch = """"c Then inQuote = False
                i += 1
            ElseIf ch = """"c Then
                inQuote = True
                i += 1
            ElseIf ch = "\"c Then
                ' FIX C-D1: backslash escape, mirroring
                ' StripFormatCodeLiterals - in Excel, \" is a literal "
                ' character, not a quote delimiter: skip the escaped
                ' character too (append neither, do not toggle quote
                ' state). A trailing lone \ just ends the loop (i += 2
                ' runs past the end safely). So 0\"[h]" strips to 0[h]:
                ' the genuine elapsed token survives and IsTimeFormatCode
                ' = True is a true positive, not a false one. (FIX C-D2:
                ' the old comment said the new code strips to 0 and the
                ' old code let [h] survive - both wrong. The OLD code
                ' returned 0\ for 0\"[h]": the \" toggled quote state
                ' and [h] was swallowed inside the "quote", a
                ' false-negative elapsed miss, same as 0\"[h].) A plain
                ' quoted literal 0"[h]" still strips to 0, correctly
                ' no time drop.
                i += 2
            Else
                stripped.Append(ch)
                i += 1
            End If
        End While
        Return stripped.ToString()
    End Function

    ''' <summary>
    ''' D6: True when a custom format code is a PURE TIME format
    ''' ("h:mm", "mm:ss", "[h]:mm:ss", "h", "ss", "[ss]"): a colon,
    ''' an hour token, a seconds token, or a bracketed elapsed token,
    ''' with no year/day token. A code with a y/d token is a
    ''' date/datetime, handled by IsDateFormatCode, not here.
    ''' </summary>
    Private Function IsTimeFormatCode(formatCode As String) As Boolean
        If String.IsNullOrEmpty(formatCode) Then Return False
        Dim code As String = StripFormatCodeLiterals(formatCode).ToLowerInvariant()
        If code.Contains("y"c) OrElse code.Contains("d"c) Then Return False
        If code.Contains(":"c) Then Return True
        If code.Contains("h"c) Then Return True
        ' FIX C-D3: bracketed elapsed tokens ("[h]", "[m]", "[s]",
        ' "[mm]", "[ss]" ...) are deleted entirely by
        ' StripFormatCodeLiterals, so they are tested on the code with
        ' quoted literals removed but bracketed sections KEPT — a quoted
        ' literal such as 0"[h]" is removed with its quotes, so it can
        ' no longer misfire, while a real unquoted [h]/[s] survives.
        Dim unquoted As String = StripFormatCodeQuotes(formatCode).ToLowerInvariant()
        For Each elapsed As String In New String() {"[h]", "[hh]", "[m]", "[mm]", "[s]", "[ss]"}
            If unquoted.Contains(elapsed) Then Return True
        Next
        ' C-L1: seconds-only formats ("s", "ss", "ss.0" — no colon,
        ' no h). Test for a bare seconds TOKEN (a run of letters
        ' that is all "s"), not a substring, so an "s" inside some
        ' other token can never misfire.
        Dim token As New System.Text.StringBuilder()
        For i As Integer = 0 To code.Length - 1
            Dim ch As Char = code(i)
            If ch >= "a"c AndAlso ch <= "z"c Then
                token.Append(ch)
            Else
                If IsSecondsToken(token.ToString()) Then Return True
                token.Length = 0
            End If
        Next
        If IsSecondsToken(token.ToString()) Then Return True
        Return False
    End Function

    ''' <summary>C-L1: True when a letter-run token from a stripped
    ''' format code is a bare seconds token: one or more "s" and
    ''' nothing else ("s", "ss").</summary>
    Private Shared Function IsSecondsToken(token As String) As Boolean
        If String.IsNullOrEmpty(token) Then Return False
        For Each ch As Char In token
            If ch <> "s"c Then Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' Reads a .csv file as plain text (one value per field). The whole
    ''' file is read at once: strict UTF-8 first, and on invalid UTF-8 it
    ''' is re-read with Encoding.Default (the Windows ANSI code page, e.g.
    ''' Windows-1256), so no-BOM Arabic files no longer decode to U+FFFD
    ''' with their digits lost (M12). UTF-16 files with a byte-order mark
    ''' (FF FE / FE FF) are decoded with the matching encoding (M2).
    ''' Records are accumulated line by line
    ''' until quotes balance, so a quoted field containing an embedded
    ''' newline is parsed as one field (L5). Quoted fields may contain
    ''' commas. A field in scientific notation (e.g. 1.012345678E+9) is
    ''' converted to plain digits, but only when it has no fractional part;
    ''' plain digit fields are kept raw so a leading zero is never lost.
    ''' A file that ends inside a quoted field is incomplete: SplitCsvRecords
    ''' flags it and LoadPhones reports an error instead (L1).
    ''' </summary>
    Private Function ReadCsvValues(filePath As String) As List(Of String)
        Dim values As New List(Of String)()
        Dim text As String = ReadCsvText(filePath)
        For Each record As String In SplitCsvRecords(text)
            If String.IsNullOrWhiteSpace(record) Then Continue For
            For Each field As String In ParseCsvLine(record)
                Dim cleaned As String = field.Trim()
                If String.IsNullOrEmpty(cleaned) Then Continue For
                ' L6: same rule as ConvertScientificIfIntegral - only
                ' convert when an exponent marker (E/e) is present, so a
                ' bare "1.0" / "10.0" field is left as-is.
                If cleaned.IndexOf("E"c) >= 0 OrElse cleaned.IndexOf("e"c) >= 0 Then
                    Dim numberValue As Decimal = 0D
                    If Decimal.TryParse(cleaned, Globalization.NumberStyles.Float,
                                        Globalization.CultureInfo.InvariantCulture, numberValue) _
                        AndAlso numberValue = Decimal.Truncate(numberValue) Then
                        cleaned = numberValue.ToString("0", Globalization.CultureInfo.InvariantCulture)
                    End If
                End If
                values.Add(cleaned)
            Next
        Next
        Return values
    End Function

    ''' <summary>
    ''' Reads a CSV file's whole text. A byte-order mark wins first:
    ''' FF FE is UTF-16LE (what PowerShell Export-Csv writes by default
    ''' for "Unicode"), FE FF is UTF-16BE, EF BB BF is UTF-8. Without a
    ''' mark, strict UTF-8 first (a byte-order mark is stripped); on
    ''' invalid UTF-8, re-read with Encoding.Default so ANSI/Windows-1256
    ''' files keep their Arabic text and Arabic-Indic digits. UTF-16
    ''' WITHOUT a mark is not detected: its bytes decode as valid UTF-8
    ''' (NUL bytes are legal UTF-8), so such a file is unsupported.
    ''' </summary>
    Private Function ReadCsvText(filePath As String) As String
        Dim bytes As Byte() = File.ReadAllBytes(filePath)
        ' M2: check the mark before the UTF-8 attempt - UTF-16LE bytes
        ' decode as valid UTF-8 (NUL is legal UTF-8), so no
        ' DecoderFallbackException would ever fire and every digit would
        ' become its own candidate.
        If bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
            Return System.Text.Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2)
        End If
        If bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
            Return System.Text.Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2)
        End If
        If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
            ' FIX C-L4 (Round 14): the strict UTF-8 attempt on the BOM
            ' branch runs inside the same try/fallback as the no-BOM
            ' path below — one invalid byte used to throw
            ' DecoderFallbackException out of this branch (it sat
            ' OUTSIDE the Try/Catch) and reject the whole file, while
            ' the identical bytes without a BOM fell back to
            ' Encoding.Default. The fallback decodes past the BOM.
            Try
                Return New System.Text.UTF8Encoding(False, True).GetString(bytes, 3, bytes.Length - 3)
            Catch ex As System.Text.DecoderFallbackException
                Return System.Text.Encoding.Default.GetString(bytes, 3, bytes.Length - 3)
            End Try
        End If
        Dim text As String = String.Empty
        Try
            text = New System.Text.UTF8Encoding(False, True).GetString(bytes)
        Catch ex As System.Text.DecoderFallbackException
            text = System.Text.Encoding.Default.GetString(bytes)
        End Try
        If text.StartsWith(ChrW(&HFEFF)) Then text = text.Substring(1)
        Return text
    End Function

    ''' <summary>
    ''' Splits whole CSV text into records. Lines are accumulated until
    ''' the double quotes balance, so a quoted field may contain an
    ''' embedded newline without splitting the record. Handles CRLF, LF
    ''' and lone CR line endings. L5: per RFC 4180 a '"' only opens a
    ''' quoted field at the start of a field - a stray '"' in the middle
    ''' of a field is a literal character and never swallows the rest of
    ''' the file. L1: a file that ends inside a quoted field is
    ''' incomplete: the trailing partial record is NOT added and
    ''' _csvUnclosedQuote is set so LoadPhones reports an error.
    ''' </summary>
    Private Function SplitCsvRecords(text As String) As List(Of String)
        Dim records As New List(Of String)()
        _csvUnclosedQuote = False
        Dim normalized As String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim lines As String() = normalized.Split(New Char() {ControlChars.Lf})
        Dim current As New System.Text.StringBuilder()
        Dim inQuotes As Boolean = False
        Dim atFieldStart As Boolean = True
        For Each line As String In lines
            If current.Length > 0 Then current.Append(vbLf)
            current.Append(line)
            Dim i As Integer = 0
            While i < line.Length
                If inQuotes Then
                    If line(i) = """"c Then
                        If i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                            i += 2
                            Continue While
                        End If
                        inQuotes = False
                        atFieldStart = False
                    End If
                    i += 1
                Else
                    If line(i) = """"c AndAlso atFieldStart Then
                        inQuotes = True
                        atFieldStart = False
                        i += 1
                    ElseIf line(i) = ","c Then
                        atFieldStart = True
                        i += 1
                    Else
                        atFieldStart = False
                        i += 1
                    End If
                End If
            End While
            If Not inQuotes AndAlso current.Length > 0 Then
                records.Add(current.ToString())
                current.Clear()
                atFieldStart = True
            End If
        Next
        If current.Length > 0 Then
            If inQuotes Then
                ' L1: the file ends inside a quoted field - an incomplete
                ' record, not a complete one. Do not add it; LoadPhones
                ' reports it as an error instead.
                _csvUnclosedQuote = True
            Else
                records.Add(current.ToString())
            End If
        End If
        Return records
    End Function

    ''' <summary>
    ''' Splits one CSV line into fields, respecting double-quoted fields
    ''' (commas inside quotes do not split; "" inside quotes is one quote).
    ''' Per RFC 4180 a '"' only opens a quoted field at the start of a
    ''' field; a stray '"' in the middle of a field is a literal character
    ''' (L5, same rule as SplitCsvRecords).
    ''' </summary>
    Private Function ParseCsvLine(line As String) As List(Of String)
        Dim fields As New List(Of String)()
        Dim sb As New System.Text.StringBuilder()
        Dim inQuotes As Boolean = False
        Dim atFieldStart As Boolean = True
        Dim i As Integer = 0
        While i < line.Length
            Dim ch As Char = line(i)
            If inQuotes Then
                If ch = """"c Then
                    If i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        sb.Append(""""c)
                        i += 2
                        Continue While
                    End If
                    inQuotes = False
                    atFieldStart = False
                    i += 1
                    Continue While
                End If
                sb.Append(ch)
                atFieldStart = False
                i += 1
            Else
                If ch = """"c AndAlso atFieldStart Then
                    inQuotes = True
                    atFieldStart = False
                    i += 1
                ElseIf ch = ","c Then
                    fields.Add(sb.ToString())
                    sb.Clear()
                    atFieldStart = True
                    i += 1
                Else
                    sb.Append(ch)
                    atFieldStart = False
                    i += 1
                End If
            End If
        End While
        fields.Add(sb.ToString())
        Return fields
    End Function

End Class
