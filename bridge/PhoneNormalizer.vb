Imports System.Text.RegularExpressions

''' <summary>
''' Normalization and validation of Egyptian mobile numbers.
''' No phone numbers are hard-coded here - only the country code (+20)
''' and the valid Egyptian mobile prefixes (10 / 11 / 12 / 15).
''' </summary>
Public Module PhoneNormalizer

    ''' <summary>
    ''' Maps Arabic-Indic digits (U+0660-U+0669), Extended Arabic-Indic
    ''' digits (U+06F0-U+06F9) and full-width digits (U+FF10-U+FF19)
    ''' to ASCII 0-9. The Arabic decimal separator (U+066B), the
    ''' full-width full stop (U+FF0E) and the middle dot (U+00B7)
    ''' map to "." instead - they are decimal separators, so the
    ''' decimal machinery treats them exactly like "." (FIX
    ''' C-L3b, Round 24). Other characters are unchanged.
    ''' </summary>
    Public Function ConvertArabicIndicDigits(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        Dim chars As Char() = text.ToCharArray()
        For i As Integer = 0 To chars.Length - 1
            Dim code As Integer = AscW(chars(i))
            If code >= &H660 AndAlso code <= &H669 Then
                chars(i) = ChrW(AscW("0"c) + (code - &H660))
            ElseIf code >= &H6F0 AndAlso code <= &H6F9 Then
                chars(i) = ChrW(AscW("0"c) + (code - &H6F0))
            ElseIf code >= &HFF10 AndAlso code <= &HFF19 Then
                ' FIX C-L2 (Round 23): full-width digits
                ' (U+FF10-U+FF19) map to ASCII 0-9 as well - without
                ' this a full-width number was silently dropped
                ' (0 candidates) instead of counted.
                chars(i) = ChrW(AscW("0"c) + (code - &HFF10))
            ElseIf code = &H66B Then
                ' FIX C-L3b (Round 24): Arabic decimal separator
                ' (U+066B) maps to "." - without this the
                ' extractor split on it (2 invalid) while
                ' NormalizePhone saw no decimal and concatenated
                ' the digits into a phantom valid phone. The
                ' existing C-H1 machinery now handles it: the
                ' extractor merges via the "." gap and
                ' NormalizePhone rejects the single decimal.
                chars(i) = "."c
            ElseIf code = &HFF0E Then
                ' FIX C-L3b (Round 24): full-width full stop
                ' (U+FF0E) maps to "." for the same reason - it
                ' is the full-width form of the decimal separator.
                chars(i) = "."c
            ElseIf code = &HB7 Then
                ' FIX C-L3b (Round 24): middle dot (U+00B7) maps
                ' to "." for the same reason - it is used as a
                ' decimal separator. Only the two call sites in
                ' this module (NormalizePhone, the extractor) use
                ' this function, so the mapping can only add
                ' decimal rejections in NormalizePhone and make
                ' the lookalike behave like "." in the extractor.
                chars(i) = "."c
            End If
        Next
        Return New String(chars)
    End Function

    ''' <summary>
    ''' FIX Unicode-direct (Round 27): True when ch is a minus
    ''' sign or a Unicode minus lookalike - ASCII "-" (U+002D)
    ''' plus U+2212, U+2013, U+FF0D, U+2010, U+2011, U+2012,
    ''' U+2014, U+2015, U+FE58, U+FE63. Used only for the
    ''' leading-character check in NormalizePhone: a lookalike
    ''' in mid-field is NOT mapped to "-" (that would turn it
    ''' into an intra-number separator), it simply stays an
    ''' unmapped character there.
    ''' </summary>
    Private Function IsMinusLookalike(ch As Char) As Boolean
        Select Case AscW(ch)
            Case &H2D, &H2212, &H2013, &HFF0D, &H2010, &H2011, &H2012, &H2014, &H2015, &HFE58, &HFE63
                Return True
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Normalizes a raw value to international format (+20xxxxxxxxxx).
    ''' Returns True when the value is a valid Egyptian mobile number.
    ''' </summary>
    Public Function NormalizePhone(raw As String, ByRef normalized As String) As Boolean
        normalized = String.Empty
        If String.IsNullOrWhiteSpace(raw) Then Return False

        ' Map Arabic-Indic digits to ASCII first, then keep only the
        ' leading "+" (if any) and the ASCII digits.
        Dim text As String = ConvertArabicIndicDigits(raw.Trim())
        ' L5: "++20..." or a "+" anywhere but position 0 is not a
        ' valid number - reject before digit extraction strips the
        ' extra "+" away and makes it look valid.
        If text.StartsWith("++") OrElse text.IndexOf("+"c, 1) >= 0 Then Return False
        ' C-L2: a leading "-" is a negative number, not a phone:
        ' reject it before digit extraction strips it away and
        ' makes "-01012345678" look valid. Only position 0: a
        ' "-" INSIDE a number ("010-12345678") stays accepted.
        ' FIX Unicode-direct (Round 27): the ASCII-only check
        ' above let the 10 Unicode minus lookalikes through,
        ' so direct NormalizePhone("–01012345678") returned
        ' True while the extractor rejected the same input. A
        ' leading lookalike is the same negative number, not
        ' a phone - check the first char against all of them.
        If text.Length > 0 AndAlso IsMinusLookalike(text(0)) Then Return False
        ' FIX C-L1 (Round 31, gated Round 32, reworked Round 33, restructured Round 35): "+<sign>20..." — a minus lookalike after "+" is a signed number, not a phone, when the sign run or the "." run that follows it reaches a digit: the extractor carries the sign only then, and carries a "." the same way (so "+-.201..." fires via the dot phase, not via the sign reaching the digit immediately). A space/separator after the sign ("+- 20...", "+-/20...") breaks the run, the extractor drops the gap and loads a valid phone, so direct must stay valid there too. (Round 35) The gate now mirrors ExtractPhoneCandidates exactly: it splits text with the SAME field-split rule (on the same mapped text, modulo NormalizePhone's leading/trailing Trim) — ";" / CR / LF / tab / "|" / Arabic comma (،) / Arabic semicolon (؛), a comma that is not a decimal separator, and a colon with no digit on either side (ignoring adjacent spaces) — and applies the two-phase scan (sign run of "+" / minus lookalikes, then "." run) to each field's TrimStart'ed leading gap, firing (Return False) iff ANY field's scan reaches an ASCII digit with hasMinus or dotCount > 0. Entry is whenever text starts with "+" (not only when text(1) is a minus — a later field can carry the sign, e.g. "+,-201..."). A "," is never swallowed into the sign/dot runs: the split removes it. A bare "+" before a digit after a comma must NOT fire (the gate leaves "+-,+201..." to the pre-existing :96 second-'+' rule, which rejects it direct while the extractor loads it valid — out of gate scope).
        If text.StartsWith("+") Then
            Dim fields As String() = Regex.Split(text, "[;\r\n\t|،؛]+|(?<!\d),|,(?!\d{1,2}(?!\d))|(?<!\d\s*):(?!\s*\d)")
            For Each field As String In fields
                Dim f As String = field.TrimStart()
                Dim i As Integer = 0
                Dim hasMinus As Boolean = False
                Dim dotCount As Integer = 0
                While i < f.Length AndAlso (IsMinusLookalike(f(i)) OrElse f(i) = "+"c)
                    If IsMinusLookalike(f(i)) Then hasMinus = True
                    i += 1
                End While
                While i < f.Length AndAlso f(i) = "."c
                    dotCount += 1
                    i += 1
                End While
                If i < f.Length AndAlso f(i) >= "0"c AndAlso f(i) <= "9"c AndAlso (hasMinus OrElse dotCount > 0) Then Return False
            Next
        End If
        ' D9: any letter in the raw value rejects it — digit
        ' extraction below keeps only ASCII digits, so without this
        ' "0101234abc5678" would normalise to a valid number.
        ' Spaces, dashes, dots, slashes, parentheses and a leading
        ' "+" are still accepted (they are not letters). Verified
        ' callers pass only grid numbers (already normalised, no
        ' letters) or extractor candidates (letters close a
        ' candidate upstream), so no legitimate number is rejected.
        For Each ch As Char In text
            If Char.IsLetter(ch) Then Return False
        Next
        ' (Round 40) C-M1: a real date followed by a dot-digit tail
        ' ("01/02/2026.123") is a date with junk on it, not a phone -
        ' its digits concatenate to a valid pattern. The date match
        ' must end on a group boundary (the "." follows the day/year
        ' group directly), so "2011-11-223344" (the day group runs
        ' into the tail digits) is unaffected. A time with a
        ' dot-digit tail ("10:30:45.1234", "10:30.123456") is a time
        ' with fractional seconds / junk on it, not a phone, for the
        ' same reason. Both shapes contain separators no phone has
        ' ("/" or ":"), so no legitimate number is rejected.
        Dim dotTailDate As Match = Regex.Match(text, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))\.(\d+)\s*$")
        If dotTailDate.Success AndAlso DatePartsAreRealEitherOrder(dotTailDate) Then Return False
        ' (Round 41) C-3: a complete real date followed by a
        ' space-separated pure-digit tail ("01/02/2026 123",
        ' "01/02/2026 10.9", "01/02/2026 123 456") is a date with
        ' junk on it, not a phone - its digits concatenate to a valid
        ' pattern. The date match must end on a group boundary (the
        ' tail is separated by whitespace), so "2011-11-223344"
        ' (the day group runs into the tail digits) is unaffected,
        ' and the tail must hold >= 3 digits, so a 1-2 digit stray
        ' tail keeps its existing treatment. A colon in the tail
        ' ("11-12-2026 10:30") never matches, so datetimes are
        ' unaffected.
        Dim spaceTailDate As Match = Regex.Match(text, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))\s+([\d\s.]+)\s*$")
        If spaceTailDate.Success AndAlso DatePartsAreRealEitherOrder(spaceTailDate) AndAlso CountAsciiDigits(spaceTailDate.Groups(7).Value) >= 3 Then Return False
        ' (Round 42) C-H-1: a complete real date followed by a
        ' separator-joined digit tail ("01/02/2026 1-23",
        ' "01/02/2026-123", "01/02/2026 (123)", "01/02/2026.1-23")
        ' is a date with junk on it, not a phone. The C-3 space tail
        ' (digits/spaces/dots only) and the dot tail (pure digits
        ' after the dot) both missed tails joined by "-", "/" or
        ' parentheses, so the candidate fell through to the 10-14
        ' digit branch and its digits concatenated into a phantom
        ' valid phone. The tail must hold >= 3 digits (like C-3), and
        ' it must start with a separator after the date group: the
        ' date match ends on a group boundary, so "2011-11-223344"
        ' (the day group runs into the tail digits) cannot match.
        ' A colon in the tail never matches, so datetimes are
        ' unaffected.
        Dim sepTailDate As Match = Regex.Match(text, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))(?:[\s\-/.()]+)([\d\s\-/.()]+)\s*$")
        If sepTailDate.Success AndAlso DatePartsAreRealEitherOrder(sepTailDate) AndAlso CountAsciiDigits(sepTailDate.Groups(7).Value) >= 3 Then Return False
        If IsTimeWithDotTail(text.Trim()) Then Return False
        ' (Round 44) C-H2/C-H3: a time with a timezone offset
        ' and/or a trailing digit group ("10:30:45-02:00",
        ' "10:30:45 1234", "10:30+02:00 99") is a time with junk
        ' on it, not a phone - its digits concatenate to a valid
        ' pattern. See IsTimeWithOffsetOrTrailingTail below.
        If IsTimeWithOffsetOrTrailingTail(text.Trim()) Then Return False
        ' FIX C-H1 (Round 22): a SINGLE decimal value ("1.012345678",
        ' "0101234567.8") is not a phone - its digits concatenate to
        ' a valid pattern, so the extractor's decimal merge (H4) used
        ' to invent a phantom valid phone (+201012345678). The H4
        ' comment below claimed the merged decimal "becomes a single
        ' candidate that NormalizePhone rejects" - false until now:
        ' NormalizePhone had NO decimal rejection. Exactly one "."
        ' or "," with digits on both sides and nothing else;
        ' two-dot forms ("0101.234.5678") are not single decimals
        ' and still validate. Scientific values never reach this
        ' check as decimals: CloseCandidate converts them to plain
        ' digits first, and the letter check above rejects any
        ' unconverted exponent outright.
        ' FIX C-M2 (Round 25): a LEADING decimal separator (".10
        ' 12345678", "+.10 12345678", ",10 12345678") is a
        ' zero-less decimal (0.10 ...), not a phone -
        ' IsSingleDecimalValue needs "^\d+" and
        ' IsDecimalHeadWithTrailingGroups needs "^\+?\d+",
        ' so a leading "." matched neither and digit extraction
        ' concatenated the digits into a phantom valid phone.
        ' A leading "." / "," followed by digits can never
        ' start a legitimate phone, so reject it outright (see
        ' IsLeadingDecimalHead below). Bare "10 12345678" and
        ' junk prefixes ("$10 ...", "x10 ...") have no leading
        ' separator and are unaffected; a separator NOT followed
        ' by a digit (", 01012345678" - a list separator) does
        ' not match either.
        ' FIX D-1 (Round 26, doc only): the ",10 12345678" arm
        ' above fires on DIRECT NormalizePhone calls only. Via
        ' the extractor the leading-comma variant remains a
        ' documented ACCEPTED residual: ExtractPhoneCandidates
        ' splits on "(?<!\d)," first, so the comma never reaches
        ' this check - it is the list separator, and rejecting
        ' the bare "10 12345678" behind it would false-invalidate
        ' the list reading "empty field, then phone". No
        ' splitting, carrying or rejection behaviour changes.
        If IsLeadingDecimalHead(text) Then Return False
        If IsSingleDecimalValue(text) Then Return False
        ' FIX C-M2 (Round 24): a decimal head followed by further
        ' digit groups joined by intra-number separators ("0.10
        ' 12345678") is not a phone either - IsSingleDecimalValue
        ' needs the WHOLE string to be one decimal, so this shape
        ' slipped past it and digit extraction concatenated it
        ' into a phantom valid phone (+201012345678). See
        ' IsDecimalHeadWithTrailingGroups below.
        If IsDecimalHeadWithTrailingGroups(text) Then Return False
        Dim hasPlus As Boolean = text.StartsWith("+")
        Dim digitsBuilder As New System.Text.StringBuilder()
        For Each ch As Char In text
            If ch >= "0"c AndAlso ch <= "9"c Then digitsBuilder.Append(ch)
        Next
        Dim digits As String = digitsBuilder.ToString()
        If digits.Length = 0 Then Return False

        Dim intl As String = String.Empty

        If hasPlus Then
            ' +20xxxxxxxxxx, or +0020xxxxxxxxxx (14 digits)
            If digits.StartsWith("20") AndAlso digits.Length = 12 Then
                intl = "+" & digits
            ElseIf digits.StartsWith("0020") AndAlso digits.Length = 14 Then
                intl = "+20" & digits.Substring(4)
            Else
                Return False
            End If
        ElseIf digits.StartsWith("0020") AndAlso digits.Length = 14 Then
            ' 0020xxxxxxxxxx -> +20xxxxxxxxxx
            intl = "+20" & digits.Substring(4)
        ElseIf digits.StartsWith("20") AndAlso digits.Length = 12 Then
            ' 20xxxxxxxxxx -> +20xxxxxxxxxx
            intl = "+" & digits
        ElseIf digits.StartsWith("01") AndAlso digits.Length = 11 Then
            ' 01xxxxxxxxx -> +20 + last 10 digits (drop the leading 0)
            intl = "+20" & digits.Substring(1)
        ElseIf digits.StartsWith("1") AndAlso digits.Length = 10 Then
            ' 1xxxxxxxxx -> +201xxxxxxxxx
            intl = "+20" & digits
        Else
            Return False
        End If

        ' Validate: exactly 10 digits after +20, prefix 10 / 11 / 12 / 15.
        Dim afterCountry As String = intl.Substring(3)
        If afterCountry.Length <> 10 Then Return False
        Dim prefix As String = afterCountry.Substring(0, 2)
        If prefix <> "10" AndAlso prefix <> "11" AndAlso prefix <> "12" AndAlso prefix <> "15" Then Return False

        normalized = intl
        Return True
    End Function

    ''' <summary>
    ''' FIX C-H1 (Round 22): True when the text is one decimal value -
    ''' digits, exactly one "." or "," separator, digits, and
    ''' nothing else. A leading "+" is something else, so a dotted
    ''' international form is NOT one. The "," is tried as "."
    ''' (invariant culture). Used by NormalizePhone, so the rejection
    ''' applies everywhere IsValidCandidate is consulted, including
    ''' the extractor's merge-loop decisions.
    ''' </summary>
    Private Function IsSingleDecimalValue(text As String) As Boolean
        Dim trimmed As String = text.Trim()
        If Not Regex.IsMatch(trimmed, "^\d+[.,]\d+$") Then Return False
        Dim value As Decimal = 0D
        Return Decimal.TryParse(trimmed.Replace(","c, "."c), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, value)
    End Function

    ''' <summary>
    ''' FIX C-M2 (Round 24): True when the text is a decimal head
    ''' followed by further digit groups joined by intra-number
    ''' separators ("0.10 12345678", "0.10-12345678"): a decimal
    ''' head with digits on both sides, then one or more
    ''' separator-joined digit groups. A leading "+" is allowed
    ''' (the same phantom with a "+"). Two-dot forms
    ''' ("0101.234.5678") have two separators and still validate;
    ''' forms with no dot ("+20 10 1234 5678") never match.
    ''' Coordinator ruling: the mixed dot+space form "0101.234
    ''' 5678" (valid before) becomes invalid under this fix,
    ''' consistent with the C-H1 precedent that "0101.2345678"
    ''' is already invalid - no head-length carve-out.
    ''' FIX C-M1 (Round 25): a SECOND dot/comma in the trailing
    ''' groups ("0.10 1234.5678") defeated the Round 24 count
    ''' (sepCount had to be exactly 1, but the head dot plus the
    ''' trailing dot made 2), so the digits concatenated into a
    ''' phantom valid phone. The distinguishing feature is the
    ''' FIRST trailing separator: in the phantom it is a non-dot
    ''' separator (space/dash/...), in the valid two-dot phone
    ''' ("0101.234.5678", "0.10.12345678") it IS the dot/comma.
    ''' So the count is dropped: reject when a decimal head
    ''' ("^\+?\d+[.,]\d+") is followed by a NON-DOT separator
    ''' + further digit groups (any number of dot/comma groups
    ''' may follow); do NOT reject when the first trailing
    ''' separator is another "." / ",". "0.10 1234,5678" is
    ''' rejected here too (direct call); the extractor still
    ''' splits it on the comma (0 valid, 2 total) by design.
    ''' </summary>
    Private Function IsDecimalHeadWithTrailingGroups(text As String) As Boolean
        Dim trimmed As String = text.Trim()
        Dim head As Match = Regex.Match(trimmed, "^\+?\d+[.,]\d+")
        If Not head.Success Then Return False
        Dim rest As String = trimmed.Substring(head.Length)
        If rest.Length = 0 Then Return False
        ' FIX C-M1 (Round 25): the first trailing separator
        ' decides - a dot/comma here means a two-dot phone, not
        ' a phantom. The rest must then be digit groups joined
        ' by separators (dots/commas allowed between groups), so
        ' "0.10 1234.5678" and "0.10 1234,5678" both reject.
        ' A trailing separator run ("0101.234(5678)") is allowed
        ' at the end so the direct call agrees with the extractor,
        ' which never produces a trailing separator (gaps are
        ' only appended before a group, never after the last one).
        Dim first As Char = rest(0)
        If first = "."c OrElse first = ","c Then Return False
        Return Regex.IsMatch(rest, "^(?:[\s\-/():.,]+\d+)+(?:[\s\-/():.,]+)?$")
    End Function

    ''' <summary>
    ''' FIX C-M2 (Round 25): True when the text starts with a
    ''' decimal separator (optionally after a "+") followed by a
    ''' digit - ".10 12345678", "+.10 12345678", ",10 12345678"
    ''' (and the Unicode separators, already mapped to "." by
    ''' ConvertArabicIndicDigits). ".10" is a genuine zero-less
    ''' spelling of 0.10 - the same value as the C-M2 case "0.10
    ''' 12345678" - so it is a decimal, never a phone. A
    ''' separator run ("..10") counts too. A separator NOT
    ''' followed by a digit (", 01012345678" - a list separator)
    ''' does not match.
    ''' FIX D-1 (Round 26, doc only): the ",10" arm fires on
    ''' DIRECT NormalizePhone calls only - via the extractor a
    ''' leading comma is split off first, so ",10 12345678"
    ''' still loads as the valid bare phone behind it (accepted
    ''' residual, comma is the list separator).
    ''' </summary>
    Private Function IsLeadingDecimalHead(text As String) As Boolean
        Dim trimmed As String = text.Trim()
        Return Regex.IsMatch(trimmed, "^\+?[.,]+\d")
    End Function

    ''' <summary>
    ''' Scans a text fragment and returns every substring that looks like
    ''' a phone number. The text is first split ONLY on strong delimiters
    ''' (semicolon, newline, tab, |, Arabic comma/semicolon, and a comma
    ''' that is NOT a decimal separator, and a colon with NO digit on
    ''' either side (ignoring adjacent spaces): a comma between digits
    ''' followed by a 1-2 digit run ("1012345678,5") is a decimal
    ''' separator, not a field separator, so the decimal stays one field.
    ''' a SPACE is NOT a separator between numbers - spaces, dots,
    ''' dashes, slashes, parentheses and a time colon inside a field are
    ''' separators WITHIN one number, so
    ''' "0101 234 5678" is one candidate and "2026-11-16 10:30" stays one
    ''' datetime candidate instead of leaving a stray ":30" fragment.
    ''' Digit groups inside a field are merged into one candidate while
    ''' the total stays at or below 14 digits; the current candidate is
    ''' closed as soon as it is a COMPLETE VALID number, so two numbers
    ''' separated only by a space ("01012345678 01112345678") are never
    ''' merged into one invalid candidate, and a short fragment is closed
    ''' before a following group that alone is a valid number ("123
    ''' 01012345678"). Every digit group is found with Regex.Matches, so
    ''' a second number in the same field is never lost. A candidate only
    ''' counts when its digit count is phone-plausible (10-14 digits).
    ''' Short digit fragments (under 10 digits after merging) are still
    ''' returned, so the caller counts them as invalid instead of
    ''' silently dropping them. Only date-like values (e.g. 2026-11-16)
    ''' are excluded outright and never counted.
    ''' </summary>
    Public Function ExtractPhoneCandidates(text As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrWhiteSpace(text) Then Return result
        Dim mapped As String = ConvertArabicIndicDigits(text)
        Dim fields As String() = Regex.Split(mapped, "[;\r\n\t|،؛]+|(?<!\d),|,(?!\d{1,2}(?!\d))|(?<!\d\s*):(?!\s*\d)")
        For Each field As String In fields
            ExtractFieldCandidates(field, result)
        Next
        Return result
    End Function

    ''' <summary>
    ''' Merges the digit groups of one field (text between strong
    ''' delimiters) into candidates. Groups joined only by space / dot /
    ''' dash / slash / parentheses / time-colon merge; anything else
    ''' (letters, a "+" starting a new number) or a merge past 14 digits
    ''' closes the current candidate and starts a new one. H1: the
    ''' current candidate is ALSO closed as soon as it is a COMPLETE
    ''' VALID number, before another group is merged into it, so
    ''' "01012345678 011 12345678" splits into two valid numbers; and a
    ''' short fragment is closed before a following group that alone is
    ''' a valid number (and merging would not itself be valid), so "123
    ''' 01012345678" keeps its valid number. M1: a fragment is also
    ''' closed before a following run of groups that together form a
    ''' valid number ("123 010 12345678", a date/time before a
    ''' space-split phone) - unless the groups are still a date/time
    ''' continuation, or merging the whole would itself be valid
    ''' ("+20 10 1234 5678"). "0101 234 5678" still merges because it
    ''' is only valid once the last group arrives. H4: a decimal gap
    ''' ("." or ",") never closes a valid candidate - the decimal merges
    ''' instead. M1: a valid candidate that is a date head plus a
    ''' time-like tail is not closed until the time part has arrived.
    ''' </summary>
    Private Sub ExtractFieldCandidates(field As String, result As List(Of String))
        ' FIX C-L2 (Round 17): a digit group may carry an exponent suffix
        ' ("012345678E-9", "012345678E+9") - it is part of the SAME value,
        ' not a new group. Without this the "." gap merges via
        ' IsDecimalGapMerge into "1.012345678", the "E-" gap then closes
        ' it, and CloseCandidate adds it as a PHANTOM VALID phone
        ' (+201012345678) plus a phantom invalid for the exponent
        ' fragment. The exponent is bounded to 1-2 digits: a real
        ' phone-magnitude scientific value never needs more, and an
        ' "E" joining two plain numbers then cannot swallow more than
        ' two digits of the second one. The spaced form ("...678 E-9")
        ' is merged by the exponent-gap rule in the loop below.
        Dim matches As MatchCollection = Regex.Matches(field, "[+]*[0-9]+(?:[Ee][+-]?[0-9]{1,2})?")
        ' L5: "[+]*" keeps a "++" run inside the candidate, so
        ' NormalizePhone sees "++201..." and rejects it instead
        ' of the extractor silently dropping the extra "+". A "++"
        ' separated from the digits ("++ 201..." / "++-201...") is
        ' carried into the following candidate below for the same
        ' reason - it must reach NormalizePhone, not be discarded.
        If matches.Count = 0 Then Return
        Dim current As New System.Text.StringBuilder()
        Dim currentDigits As Integer = 0
        Dim prevEnd As Integer = 0
        For matchIndex As Integer = 0 To matches.Count - 1
            Dim m As Match = matches(matchIndex)
            Dim gap As String = String.Empty
            If m.Index > prevEnd Then gap = field.Substring(prevEnd, m.Index - prevEnd)
            ' FIX C-L2 (Round 17): a gap that is only an exponent marker
            ' ("E", "E-", "E+", with optional spaces - the spaced form
            ' of "1.012345678 E-9") joins the same scientific value,
            ' never closes the candidate: closing here would leave the
            ' mantissa ("1.012345678") to count as a phantom valid phone.
            ' (The adjacent form is glued by the match regex itself.)
            ' FIX C-M4 (Round 18): the merge is gated - a SPACED
            ' exponent gap counts as scientific only when the mantissa
            ' is decimal (current contains ".") or the exponent is
            ' signed (gap contains "+" or "-"). An unsigned spaced
            ' exponent after an integer mantissa ("01012345678 E9") is
            ' NOT an exponent gap: the close block below (H1) then
            ' closes the already-complete valid number first, and "9"
            ' becomes its own invalid fragment. "01012345678 E+9" /
            ' "E-9" still merge (signed) and count 1 invalid.
            ' FIX C-M4b (Round 18, proven by simulation): the "+" of
            ' "E+" never appears in gap - the :176 match regex opens
            ' with "[+]*", so the sign is swallowed into the NEXT
            ' group ("01012345678 E+9" matches "01012345678" then
            ' "+9", gap " E"). A literal gap-only sign test can
            ' therefore never fire for E+ (only E- shows its sign in
            ' the gap). The sign is read from the next group's leading
            ' "+" as well (m.Value.StartsWith("+")); the gate still
            ' only fires when the gap itself is an exponent marker, so
            ' a plain "+" group after a space ("01012345678 +9") is
            ' unaffected. Verified: "01012345678 E+9" merges to 1
            ' invalid, "01012345678 E9" still splits 1 valid + 1 invalid.
            Dim isExponentGap As Boolean = Regex.IsMatch(gap, "^\s*[Ee][+-]?\s*$") AndAlso (current.ToString().Contains(".") OrElse gap.Contains("+") OrElse gap.Contains("-") OrElse m.Value.StartsWith("+"))
            ' FIX C-Low-2 (Round 38): a trailing-dot exponent gap
            ' (".E" / ".e", the sign in the gap or glued into the
            ' next group) is the same scientific value, not a decimal
            ' followed by a fragment: "01012345678.E9" and "1.E+9"
            ' must merge into ONE candidate - the direct TryParse
            ' path converts them, so the extractor must glue them or
            ' the raw counts disagree with the converted ones. The
            ' dot sits adjacent to the E, so - unlike the spaced M4
            ' gate above - no decimal/signed gate applies: the form
            ' is unambiguous scientific syntax. Checked with the
            ' close branches below, like isExponentGap.
            Dim isTrailingDotExponentGap As Boolean = Regex.IsMatch(gap, "^\.[Ee][+-]?[ ]*$") ' (Round 39) C Defect 1: no leading "[ ]*" - a space-detached ".E" tail is not scientific syntax (no parser converts it), so it must not merge and destroy the valid phone; the trailing "[ ]*" stays ("1.E 9" converts fine).
            ' L5: a "++" in the gap (separated from the digits) starts
            ' a new candidate carrying the "++", so NormalizePhone
            ' still sees and rejects it.
            Dim gapHasPlusPlus As Boolean = gap.Contains("++")
            Dim prependPlusPlus As Boolean = gapHasPlusPlus
            ' C-L2: a "-" that is the leading sign of the FIRST
            ' candidate of a field (whitespace, then "-", immediately
            ' before the first digit group) is carried into the
            ' candidate, so NormalizePhone sees and rejects the
            ' negative number instead of the extractor silently
            ' dropping the sign. A "-" between two numbers is a
            ' separator, not a sign, and is never carried.
            ' FIX C-M1 (Round 22): widened to a pure SIGN RUN -
            ' "--", "-+", "+-" before the first group were silently
            ' dropped and the number loaded as valid. The run must
            ' contain at least one "-" and sit immediately before
            ' the digits (leading whitespace only): "- 010..." (a
            ' space AFTER the sign) is still the deliberate dash-
            ' bullet scope and stays valid. The "++" path above
            ' still wins when the gap also holds "++".
            Dim leadingMinus As Boolean = (matchIndex = 0 AndAlso current.Length = 0 AndAlso IsLeadingSignRun(gap))
            ' FIX C-M2 (Round 25): a LEADING decimal separator run
            ' (".", "+.", "..") before the first digit group is
            ' carried into the candidate, exactly like the leading
            ' sign run above - without this the "." was dropped and
            ' ".10 12345678" reached NormalizePhone as
            ' "10 12345678", a phantom valid phone. NormalizePhone
            ' rejects the carried "." (IsLeadingDecimalHead). Only
            ' at matchIndex = 0 with an empty candidate, and only
            ' when the gap is NOTHING but leading whitespace,
            ' optional "+"s and decimal separators immediately
            ' before the digits (a space after the dot, ". 10 ...",
            ' is junk + phone and stays dropped, like "- 010..."
            ' stays valid). FIX R24-C-M (Round 26): a sign run
            ' before the separators ("-.", "-+.", "--.", Unicode
            ' minus + ".") is carried the same way - the sign
            ' itself is dropped and only "." is appended, so
            ' NormalizePhone still rejects the candidate. A ","
            ' never reaches this point:
            ' ExtractPhoneCandidates splits a leading comma first
            ' ("(?<!\d),"), so the comma variant stays an
            ' extractor residual - see the Round 25 report.
            Dim leadingDecimalSep As Boolean = (matchIndex = 0 AndAlso current.Length = 0 AndAlso IsLeadingDecimalSeparator(gap))
            Dim groupDigits As Integer = CountAsciiDigits(m.Value)
            Dim startsNewNumber As Boolean = m.Value.StartsWith("+") AndAlso current.Length > 0
            Dim closeCurrent As Boolean = False
            If current.Length > 0 AndAlso Not isExponentGap AndAlso Not isTrailingDotExponentGap Then
                ' H4: a "." or "," gap that makes current + gap + group a
                ' decimal number (e.g. "1012345678.5") is one decimal
                ' value, not a finished phone followed by a fragment: do
                ' not close on it - merge instead, so the whole decimal
                ' becomes a single candidate that NormalizePhone rejects.
                ' FIX C-H1 (Round 22): that rejection is now real -
                ' IsSingleDecimalValue in NormalizePhone rejects the
                ' merged decimal, so no phantom valid phone is born.
                ' (Before, the digits concatenated to a valid pattern.)
                Dim isDecimalMerge As Boolean = IsDecimalGapMerge(current.ToString(), gap, m.Value)
                If gapHasPlusPlus Then closeCurrent = True
                If Not isDecimalMerge AndAlso (startsNewNumber OrElse Not GapIsIntraNumberSeparator(gap) OrElse currentDigits + groupDigits > 14) Then
                    closeCurrent = True
                ElseIf IsValidCandidate(current.ToString()) AndAlso Not isDecimalMerge AndAlso Not IsDateTimeContinuation(current.ToString(), matches, field, matchIndex) Then
                    ' H1: already a complete valid number - close it
                    ' before merging another group into it. M1: unless
                    ' it is a date head plus a time-like tail (e.g.
                    ' "11-12-2026 10" before the "30"), where the next
                    ' group is still a time part, not a new number.
                    closeCurrent = True
                ElseIf IsValidCandidate(m.Value) AndAlso Not isDecimalMerge AndAlso Not IsValidCandidate(current.ToString() & gap & m.Value) Then
                    ' FIX C-M1 (Round 23): this branch had NO
                    ' isDecimalMerge guard (the first close branch and
                    ' the H1 branch both have one) - a decimal gap
                    ' whose fractional part alone is a valid phone
                    ' ("0.1012345678") closed the "0" and let the
                    ' fraction through as a PHANTOM VALID phone,
                    ' bypassing the C-H1 rejection of the merged
                    ' decimal. A decimal gap always merges now, so the
                    ' whole value counts once as invalid. Space gaps
                    ' ("123 01012345678") have isDecimalMerge=False,
                    ' so the branch's legitimate job is unaffected.
                    ' The next group alone is a complete valid number and
                    ' merging it onto the current fragment would not be
                    ' valid either: close the fragment as invalid first.
                    closeCurrent = True
                ElseIf Not IsDateTimeContinuation(current.ToString(), matches, field, matchIndex) Then
                    ' M1: the groups from m onward (joined across
                    ' intra-number separators only) already form a
                    ' complete valid number, and merging them onto the
                    ' current fragment would not be valid either: close
                    ' the fragment first, so "123 010 12345678" keeps
                    ' its phone, and a date/time ("2026-11-16", "10:30")
                    ' is closed before a following space-split phone.
                    ' The combined-valid check keeps "+20 10 1234 5678"
                    ' merging (suffix and whole are both valid).
                    Dim suffixValid As String = SuffixValidPhone(matches, field, matchIndex)
                    ' FIX C-M1 (Round 23): same missing-guard root
                    ' cause as the m-valid branch above - a decimal
                    ' gap ("5.10 12345678") closed the "5" and let
                    ' "10 12345678" through as a phantom valid phone.
                    ' Guarded with Not isDecimalMerge so a decimal gap
                    ' always merges; space gaps are unaffected.
                    ' (Round 42) C-H-2: do NOT close on the suffix
                    ' when current plus a PREFIX of the suffix groups
                    ' is already a complete valid phone ("10 12 345678
                    ' 99": the suffix "12 345678 99" is itself valid,
                    ' but "10" + "12 345678" is the real phone - closing
                    ' "10" first loses it and validates the phantom).
                    ' The full-suffix case is already excluded by the
                    ' Not IsValidCandidate(current & suffixValid) test
                    ' below; this guards the shorter prefixes.
                    If suffixValid IsNot Nothing AndAlso Not isDecimalMerge AndAlso Not IsValidCandidate(current.ToString() & suffixValid) AndAlso Not CurrentPlusSuffixPrefixIsValid(current.ToString(), matches, field, matchIndex) Then
                        closeCurrent = True
                    End If
                End If
            End If
            If closeCurrent Then
                CloseCandidate(current.ToString(), result)
                current.Clear()
                currentDigits = 0
            End If
            Dim groupValue As String = m.Value
            If prependPlusPlus Then
                ' L2: strip any leading "+"s off the group before
                ' prepending "++", so "++ +201..." becomes "++201..."
                ' (still rejected by NormalizePhone) instead of the
                ' corrupted "+++201...".
                groupValue = groupValue.TrimStart("+"c)
                current.Append("++")
            ElseIf leadingMinus Then
                current.Append("-")
            ElseIf leadingDecimalSep Then
                ' FIX C-M2 (Round 25): carry the separator
                ' normalised to "." (ConvertArabicIndicDigits
                ' already mapped U+066B / U+FF0E / U+00B7 to "."
                ' upstream, so the gap here is "."; a "," arm is
                ' unreachable - see the comment above).
                current.Append(".")
            ElseIf current.Length > 0 Then
                current.Append(gap)
            End If
            current.Append(groupValue)
            currentDigits += groupDigits
            prevEnd = m.Index + m.Length
        Next
        If current.Length > 0 Then CloseCandidate(current.ToString(), result)
    End Sub

    ''' <summary>
    ''' M1: the first prefix of the digit groups from startIndex onward,
    ''' joined across intra-number separators only, that is already a
    ''' complete valid phone number - or Nothing when no such prefix
    ''' exists. Stops at a non-separator gap, a group starting a new
    ''' number ("+"), a "++" gap, or past 14 digits.
    ''' </summary>
    Private Function SuffixValidPhone(matches As MatchCollection, field As String, startIndex As Integer) As String
        Dim joined As New System.Text.StringBuilder()
        Dim joinedDigits As Integer = 0
        Dim prevEnd As Integer = 0
        Dim sawDotGap As Boolean = False
        For j As Integer = startIndex To matches.Count - 1
            Dim mj As Match = matches(j)
            If j > startIndex Then
                Dim gap2 As String = field.Substring(prevEnd, mj.Index - prevEnd)
                If gap2.Contains("++") OrElse mj.Value.StartsWith("+") OrElse Not GapIsIntraNumberSeparator(gap2) Then Return Nothing
                ' (Round 40) C-L1: a "." / "," gap into the LAST group,
                ' with a 1-2 digit tail that parses as a decimal fraction
                ' of the group before it ("345678.99"), is a decimal
                ' point, not a phone separator: the suffix must not
                ' complete its validity by swallowing that fraction
                ' ("12 345678.99" -> +201234567899 steals the real
                ' phone "10 12 345678"). The block fires only on the
                ' last group, only for a 1-2 digit tail, and only when
                ' no dot gap was crossed yet in this suffix - a dotted
                ' phone ("0101.234.5678", "01012.345.678") still joins
                ' across its dots, and a tail that is not a decimal
                ' fraction still joins.
                If (gap2 = "." OrElse gap2 = ",") AndAlso j = matches.Count - 1 AndAlso Not sawDotGap AndAlso mj.Value.Length >= 1 AndAlso mj.Value.Length <= 2 Then
                    Dim prevGroup As String = matches(j - 1).Value
                    Dim decValue As Decimal = 0D
                    If Decimal.TryParse(prevGroup & "." & mj.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, decValue) Then Return Nothing
                End If
                If gap2 = "." OrElse gap2 = "," Then sawDotGap = True
            End If
            joined.Append(mj.Value)
            joinedDigits += CountAsciiDigits(mj.Value)
            If joinedDigits > 14 Then Return Nothing
            If IsValidCandidate(joined.ToString()) Then Return joined.ToString()
            prevEnd = mj.Index + mj.Length
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' (Round 42) C-H-2: True when text is exactly one complete real
    ''' date (either order, any of the date separators) - the guard
    ''' that keeps the M1 suffix close from re-slicing a date into a
    ''' phone. Shares the date shape and the real-date test with
    ''' NormalizePhone's tail gates, so all four stay consistent.
    ''' </summary>
    Private Function IsCompleteRealDate(text As String) As Boolean
        Dim m As Match = Regex.Match(text.Trim(), "^(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))$")
        If Not m.Success Then Return False
        Return DatePartsAreRealEitherOrder(m)
    End Function

    ''' <summary>
    ''' (Round 42) C-H-2: True when current plus a PREFIX of the digit
    ''' groups from startIndex onward (joined across intra-number
    ''' separators only, same join rules as SuffixValidPhone) is already
    ''' a complete valid phone number. Used to gate the M1 suffix close:
    ''' the suffix may itself be a valid-looking number while a shorter
    ''' prefix of it completes the real phone in current ("10" + "12
    ''' 345678" = +201012345678, while the whole suffix "12 345678 99"
    ''' is the phantom +201234567899). Stops at a non-separator gap, a
    ''' group starting a new number ("+"), a "++" gap, or past 14 digits,
    ''' exactly like SuffixValidPhone; a "." / "," gap into the last
    ''' group with a 1-2 digit decimal tail ends the scan the same way
    ''' (the tail is a fraction, not a phone part).
    ''' </summary>
    Private Function CurrentPlusSuffixPrefixIsValid(current As String, matches As MatchCollection, field As String, startIndex As Integer) As Boolean
        ' (Round 42) C-H-2 carve-out: a COMPLETE real date must never
        ' re-slice into a phone ("11-12-2026" + "10" concatenates to
        ' the valid-looking "1112202610", and "20 11 2026" + "10 12"
        ' to "201120261012"): when current is exactly a date, the
        ' M1 close must fire so the date closes before the phone that
        ' follows it ("11-12-2026 10 30.123456" keeps +201030123456).
        If IsCompleteRealDate(current) Then Return False
        Dim joined As New System.Text.StringBuilder()
        Dim joinedDigits As Integer = 0
        Dim prevEnd As Integer = 0
        If startIndex > 0 Then prevEnd = matches(startIndex - 1).Index + matches(startIndex - 1).Length
        Dim sawDotGap As Boolean = False
        ' (Round 42) C-H-2 date-head carve-out: a prefix of the suffix
        ' that completes current into a complete real date ("11" +
        ' "12-2026" = "11-12-2026") means current is a date being
        ' built, not a phone fragment: the guard must not fire there,
        ' or the date never closes and a following phone is lost - and
        ' the malformed-timezone forms ("+HH:MM:SS", "+002:00") keep
        ' their existing treatment instead of being silently covered.
        Dim dateText As New System.Text.StringBuilder()
        dateText.Append(current)
        For k As Integer = startIndex To matches.Count - 1
            Dim mk As Match = matches(k)
            Dim gapk As String = String.Empty
            If mk.Index > prevEnd Then gapk = field.Substring(prevEnd, mk.Index - prevEnd)
            If k > startIndex Then
                Dim gap2 As String = field.Substring(prevEnd, mk.Index - prevEnd)
                If gap2.Contains("++") OrElse mk.Value.StartsWith("+") OrElse Not GapIsIntraNumberSeparator(gap2) Then Return False
                If (gap2 = "." OrElse gap2 = ",") AndAlso k = matches.Count - 1 AndAlso Not sawDotGap AndAlso mk.Value.Length >= 1 AndAlso mk.Value.Length <= 2 Then
                    Dim prevGroup As String = matches(k - 1).Value
                    Dim decValue As Decimal = 0D
                    If Decimal.TryParse(prevGroup & "." & mk.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, decValue) Then Return False
                End If
                If gap2 = "." OrElse gap2 = "," Then sawDotGap = True
            End If
            dateText.Append(gapk)
            dateText.Append(mk.Value)
            ' (Round 43) H-1: a prefix of the suffix that completes
            ' current into a complete real date ("11" + "12-2026" =
            ' "11-12-2026") used to return False at once — before the
            ' phone-prefix check below it — so the date closed and its
            ' groups re-sliced into a phantom valid phone whenever
            ' the date was followed by a time with a fraction/offset
            ' ("11-12-2026 10:30.5+02:00:00" became +201220261030).
            ' The remainder after the date decides now: a phone-shaped
            ' remainder still closes the date (the real phone wins);
            ' a time-shaped remainder — a time with a fraction and an
            ' optional offset tail, canonical or malformed — blocks
            ' the close, so the field merges as one datetime (0 valid;
            ' no phantom, and no real phone lost: there is none, and
            ' a real phone after the time is still split off later by
            ' the valid-number close); any other remainder keeps the
            ' old treatment (close).
            If IsCompleteRealDate(dateText.ToString()) Then
                If k + 1 < matches.Count AndAlso SuffixValidPhone(matches, field, k + 1) IsNot Nothing Then Return False
                If k + 1 < matches.Count AndAlso RemainderStartsWithTimeTail(field.Substring(matches(k).Index + matches(k).Length)) Then Return True
                Return False
            End If
            joined.Append(mk.Value)
            joinedDigits += CountAsciiDigits(mk.Value)
            If joinedDigits > 14 Then Return False
            If IsValidCandidate(current & joined.ToString()) Then Return True
            prevEnd = mk.Index + mk.Length
        Next
        Return False
    End Function

    ''' <summary>
    ''' FIX C-M1 (Round 22): True when the gap before the first digit
    ''' group is a pure sign run - after stripping LEADING whitespace
    ''' only, every character is "+" or "-" and at least one is
    ''' "-". Whitespace between the sign run and the digits ("- ")
    ''' fails the test, so the dash-bullet scope stays untouched.
    ''' FIX C-L3 (Round 23): U+2212 MINUS SIGN and U+2013 EN DASH
    ''' count as "-" here too - otherwise they were dropped and the
    ''' number loaded as valid. The carried character stays ASCII
    ''' "-" (appended by the caller), so NormalizePhone still
    ''' rejects it. U+2014 EM DASH is deliberately NOT included:
    ''' its use as a separator at position 0 could not be ruled out.
    ''' FIX C-L2b (Round 24): eight more minus lookalikes count
    ''' as "-" too - U+FF0D full-width hyphen-minus, U+2010
    ''' hyphen, U+2011 non-breaking hyphen, U+2012 figure dash,
    ''' U+2014 em dash, U+2015 horizontal bar, U+FE58 small em
    ''' dash, U+FE63 small hyphen-minus. U+2014 is now included:
    ''' the phantom ("—01012345678" loading valid) is proven
    ''' and inclusion is safe - a space anywhere in the run makes
    ''' this function return False ("— 010..." stays valid),
    ''' so the Round 23 separator worry does not apply.
    ''' </summary>
    Private Function IsLeadingSignRun(gap As String) As Boolean
        Dim t As String = gap.TrimStart()
        If t.Length = 0 Then Return False
        Dim hasMinus As Boolean = False
        For Each ch As Char In t
            If ch = "-"c OrElse ch = ChrW(&H2212) OrElse ch = ChrW(&H2013) OrElse ch = ChrW(&HFF0D) OrElse ch = ChrW(&H2010) OrElse ch = ChrW(&H2011) OrElse ch = ChrW(&H2012) OrElse ch = ChrW(&H2014) OrElse ch = ChrW(&H2015) OrElse ch = ChrW(&HFE58) OrElse ch = ChrW(&HFE63) Then
                hasMinus = True
            ElseIf ch <> "+"c Then
                Return False
            End If
        Next
        Return hasMinus
    End Function

    ''' <summary>
    ''' FIX C-M2 (Round 25): True when the gap before the first digit
    ''' group is nothing but leading whitespace, optional "+"s and
    ''' one or more decimal separators (".", "..", "+.") 
    ''' immediately before the digits - the zero-less decimal marker
    ''' (".10" = 0.10) that the extractor used to drop. A space
    ''' after the separator (". 10") fails, so junk + phone stays
    ''' valid. A "," here is unreachable (leading commas are split
    ''' off in ExtractPhoneCandidates) but is accepted by the
    ''' pattern for symmetry with the decimal-separator contract.
    ''' FIX R24-C-M (Round 26): a leading SIGN RUN before the
    ''' separators counts too ("-.", "-+.", "--.", and the same
    ''' Unicode minus lookalikes IsLeadingSignRun accepts) -
    ''' without this "-.10 12345678" was dropped to the bare
    ''' "10 12345678", a phantom valid phone, while "+.10"
    ''' was carried. The carried character stays "." (the
    ''' caller appends only "."), so NormalizePhone rejects
    ''' it via IsLeadingDecimalHead exactly as for "+.10".
    ''' A sign run with NO separator ("-") still returns
    ''' False here - that is the leadingMinus path's job - and
    ''' a space after the separator ("-. 10") still fails, so
    ''' the deliberate "- " scope is preserved.
    ''' </summary>
    Private Function IsLeadingDecimalSeparator(gap As String) As Boolean
        Dim t As String = gap.TrimStart()
        If t.Length = 0 Then Return False
        Dim i As Integer = 0
        While i < t.Length AndAlso IsDecimalSeparatorSign(t(i))
            i += 1
        End While
        If i = t.Length Then Return False
        While i < t.Length
            If t(i) <> "."c AndAlso t(i) <> ","c Then Return False
            i += 1
        End While
        Return True
    End Function

    ''' <summary>
    ''' FIX R24-C-M (Round 26): a sign character for the leading-
    ''' decimal-separator carry - ASCII "+" / "-" plus the same
    ''' Unicode minus lookalikes IsLeadingSignRun accepts
    ''' (U+2212, U+2013, U+FF0D, U+2010, U+2011, U+2012, U+2014,
    ''' U+2015, U+FE58, U+FE63). Kept as a separate helper so the
    ''' two sign-run definitions cannot drift apart silently.
    ''' </summary>
    Private Function IsDecimalSeparatorSign(ch As Char) As Boolean
        Return ch = "+"c OrElse ch = "-"c OrElse ch = ChrW(&H2212) OrElse ch = ChrW(&H2013) OrElse ch = ChrW(&HFF0D) OrElse ch = ChrW(&H2010) OrElse ch = ChrW(&H2011) OrElse ch = ChrW(&H2012) OrElse ch = ChrW(&H2014) OrElse ch = ChrW(&H2015) OrElse ch = ChrW(&HFE58) OrElse ch = ChrW(&HFE63)
    End Function

    ''' <summary>
    ''' H4: True when the gap joining the current candidate to the next
    ''' group is a decimal separator ("." or ",") and current + gap +
    ''' group parses as one decimal number (e.g. "1012345678.5"):
    ''' the two parts are one decimal value, not a finished phone
    ''' followed by a fragment, so the candidate must merge, not close.
    ''' A "," is tried with "." as the separator (invariant culture);
    ''' NumberStyles.Float (no thousands separators) is used, so two
    ''' numbers joined by a dot or comma do not parse as one decimal
    ''' merely by grouping - the m-valid branch still splits those.
    ''' </summary>
    Private Function IsDecimalGapMerge(current As String, gap As String, groupValue As String) As Boolean
        If gap <> "." AndAlso gap <> "," Then Return False
        Dim text As String = (current & gap & groupValue).Replace(","c, "."c)
        Dim value As Decimal = 0D
        Return Decimal.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, value)
    End Function

    ''' <summary>
    ''' M1 guard: True when the group at matchIndex is still a
    ''' continuation of the date/time in current (a month/day of a
    ''' partial date, the hour/minute of a time after a date, or the
    ''' short leftover after a complete time) - so the suffix lookahead
    ''' must NOT fire there: those groups are date/time parts, not the
    ''' start of a phone. H1/H2: the guard now consults SuffixValidPhone
    ''' before returning True - when the groups from m onward already
    ''' form a valid phone, it is a phone, not a date/time part
    ''' ("2026-11-16 10 12 345678", "7 10 12345678"). A bare fragment
    ''' that is the start of a date is protected by the date-prefix
    ''' guards below, so a date is never re-sliced into a phone
    ''' ("11-12-2026 10:30" must not become 1220261030, and
    ''' "2026-11-16 10:30 99" must not become 1116103099). A complete
    ''' date/time is NOT a continuation when the next group cannot
    ''' extend it (e.g. "010" after "2026-11-16" is not an hour, since
    ''' the group after it is not a 2-digit minute), so the lookahead
    ''' may fire there.
    ''' </summary>
    Private Function IsDateTimeContinuation(current As String, matches As MatchCollection, field As String, matchIndex As Integer) As Boolean
        Dim nextGroup As Match = matches(matchIndex)
        If nextGroup.Value.StartsWith("+") Then Return False
        Dim nextDigits As Integer = CountAsciiDigits(nextGroup.Value)
        Dim hasFollowing As Boolean = matchIndex + 1 < matches.Count
        Dim followingDigits As Integer = 0
        If hasFollowing Then followingDigits = CountAsciiDigits(matches(matchIndex + 1).Value)
        Dim cur As String = current.Trim()
        Dim dateMatch As Match = Regex.Match(cur, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))")
        ' (Round 40) C-L1: the date regex can match only the first 4 digits
        ' of a longer digit run as the "year" ("10 12 345678" matches year
        ' 3456, leaving "78") - when the match ends in the middle of a
        ' digit run the head is NOT a date at all: skip the date branch
        ' (same guard as CloseCandidate and FormsDateTimeReading). Without
        ' this the phantom date "10/12/3456" plus "hour" 78 made the H1
        ' close refuse, and the ".99" tail merged the real phone
        ' "10 12 345678" into one invalid decimal instead of closing it.
        If dateMatch.Success AndAlso DatePartsAreRealEitherOrder(dateMatch) AndAlso Not (dateMatch.Length < cur.Length AndAlso Char.IsDigit(cur(dateMatch.Length))) Then ' FIX C-D1: US dates (11/16/2026) are real dates too - the guard must match CloseCandidate, else the minute leaks as a stray invalid fragment.
            Dim rest As String = cur.Substring(dateMatch.Length).Trim().TrimStart(" "c, ":"c, "-"c, "/"c, "."c)
            If rest.Length = 0 Then
                ' Complete date: the next group is a time start only
                ' when it is an hour followed by a 2-digit minute
                ' (or is the last group, i.e. a bare hour). H1: unless
                ' the groups from m onward already form a valid phone
                ' ("2026-11-16 10 12 345678") - then it is a phone, not
                ' a time start, and the date must close before it.
                ' (Round 40) C-M1: unless the suffix reads as a time
                ' with a dot-digit tail ("10:30.123456") - then it is
                ' the datetime's own tail, not a phone, and the date
                ' must keep building.
                ' (Round 43) H-2: the suffix may start with a time
                ' tail (fraction/offset) followed by a phone or junk
                ' (SuffixIsTimeWithDotTail is prefix-aware now): a
                ' phone after the tail still wins — the date closes
                ' before it, exactly as before — but a tail followed
                ' by junk (or nothing) is the datetime's own tail and
                ' the date must keep building.
                If SuffixValidPhone(matches, field, matchIndex) IsNot Nothing AndAlso (Not SuffixIsTimeWithDotTail(matches, field, matchIndex) OrElse SuffixTimeTailThenPhone(matches, field, matchIndex)) Then Return False
                Return (nextDigits = 1 OrElse nextDigits = 2) AndAlso (Not hasFollowing OrElse followingDigits = 2)
            End If
            If Regex.IsMatch(rest, "^\d{1,2}$") Then
                ' H1: same - a valid phone suffix starting at m
                ' ("2026-11-16 10 15 12345678") is a phone, not the
                ' minute of the time after the date.
                If SuffixValidPhone(matches, field, matchIndex) IsNot Nothing Then Return False
                Return nextDigits = 2
            End If
            If IsTimeLike(rest) Then Return (nextDigits = 1 OrElse nextDigits = 2) AndAlso Not hasFollowing
            Return False
        End If
        If Regex.IsMatch(cur, "^\d{4}$") Then
            ' H2: a year followed by a real month/day ("2026-11-16") is
            ' a date being built, not a stray fragment - let it build;
            ' the complete-date branch above sorts out phone-vs-time
            ' once the date is whole. Otherwise a valid phone suffix
            ' starting at m ("2026 10 12345678") is a phone, not a month.
            Dim yr As Integer = 0
            Dim mo As Integer = 0
            Dim dy As Integer = 0
            If hasFollowing AndAlso nextDigits >= 1 AndAlso nextDigits <= 2 AndAlso followingDigits >= 1 AndAlso followingDigits <= 2 AndAlso Integer.TryParse(cur, yr) AndAlso Integer.TryParse(nextGroup.Value, mo) AndAlso Integer.TryParse(matches(matchIndex + 1).Value, dy) AndAlso RealDateParts(yr, mo, dy) Then Return True
            If SuffixValidPhone(matches, field, matchIndex) IsNot Nothing Then Return False
            Return (nextDigits = 1 OrElse nextDigits = 2)
        End If
        If Regex.IsMatch(cur, "^\d{4}[-/. ]\d{1,2}$") Then
            ' H2: same - a year-month followed by a real day is a date
            ' being built ("2026-11-15"); otherwise a valid phone suffix
            ' starting at m is a phone, not the day.
            Dim ym As Match = Regex.Match(cur, "^(\d{4})[-/. ](\d{1,2})$")
            Dim yr2 As Integer = 0
            Dim mo2 As Integer = 0
            Dim dy2 As Integer = 0
            If ym.Success AndAlso nextDigits >= 1 AndAlso nextDigits <= 2 AndAlso Integer.TryParse(ym.Groups(1).Value, yr2) AndAlso Integer.TryParse(ym.Groups(2).Value, mo2) AndAlso Integer.TryParse(nextGroup.Value, dy2) AndAlso RealDateParts(yr2, mo2, dy2) Then Return True
            If SuffixValidPhone(matches, field, matchIndex) IsNot Nothing Then Return False
            Return (nextDigits = 1 OrElse nextDigits = 2)
        End If
        If Regex.IsMatch(cur, "^\d{1,2}[-/. ]\d{1,2}$") Then Return nextDigits = 4
        If Regex.IsMatch(cur, "^\d{1,2}$") Then
            ' H2: a valid phone suffix starting at m ("7 10 12345678")
            ' is a phone, not a date part - unless current plus the
            ' groups from m onward read as a complete real date followed
            ' by a time, a phone, or a time then a phone
            ' ("11-12-2026 10:30", "16/11/2026 10 12 345678"), in which
            ' case the date is real and must be allowed to build.
            If SuffixValidPhone(matches, field, matchIndex) IsNot Nothing AndAlso Not FormsDateTimeReading(current, matches, field, matchIndex) Then Return False
            Return (nextDigits = 1 OrElse nextDigits = 2)
        End If
        If IsTimeLike(cur) Then Return (nextDigits = 1 OrElse nextDigits = 2) AndAlso Not hasFollowing
        Return False
    End Function

    ''' <summary>
    ''' Range-free shape version of IsTimeLike: True when rest has the
    ''' SHAPE of a time (bare hour, HH:MM[:SS], "HH MM"[:SS], or a time
    ''' followed by a short leftover) regardless of the values, so a
    ''' malformed time ("24:30", "10:30 99") still counts as a time
    ''' tail, not as the start of a phone. Used only by the H1/H2
    ''' date-reading guard (FormsDateTimeReading).
    ''' </summary>
    ''' <summary>
    ''' (Round 40) C-M1: True when text is a time followed by a
    ''' dot-digit tail of any length ("10:30.5", "10:30.123",
    ''' "10.99", "10:30:45.1234") - fractional seconds / junk on
    ''' the time, not the start of a phone. Range-free, like
    ''' IsTimeShaped (a malformed time is still a time tail).
    ''' Colon and bare-hour forms only: a SPACE-separated time
    ''' with a dot tail ("10 30.123456") is NOT covered here -
    ''' it is a guard-only shape (IsTimeShaped's own arm below),
    ''' so the direct path and the CloseCandidate drop never
    ''' mistake a space/dot phone ("01 23.4567890") for a time.
    ''' </summary>
    Private Function IsTimeWithDotTail(text As String) As Boolean
        ' Colon form: HH:MM / HH:MM:SS + a dot tail of any length
        ' ("10:30.5", "10:30:45.1234") - fractional seconds / junk.
        ' (Round 41) C-4: a trailing timezone offset ("+02:00",
        ' "-05:00", "Z") after the dot tail is still the time's own
        ' tail, not the start of a phone ("11-12-2026 10:30.5+02:00").
        ' (Round 42) C-M-1: the offset may also be glued with a space
        ' before it ("10:30.5 +02:00"), a bare +HH ("+02"), or a
        ' +HHMM ("+0200"). Malformed offsets stay out: "+HH:MM:SS"
        ' and 3-digit-hour offsets ("+002:00") are not valid ISO 8601
        ' offsets, so they never match here.
        ' (Round 46) C-H3: the offset sign may be spaced from its
        ' digits ("10:30.5 - 02:00") and a trailing digit group
        ' may follow the offset ("10:30.5 - 02:00 9") - the
        ' strict offset shapes above are kept (sign run + space
        ' only); the D3 carve-out ("10 30.123456" stays valid)
        ' is untouched: this arm needs the colon form. The
        ' trailing groups are BOUNDED to 1-9 digits each: a
        ' trailing group must never itself be a phone (phone
        ' groups are 10+ digits), or the tail swallows a real
        ' phone after the time ("11-12-2026 10:30.123456
        ' 01012345678"), TimeTailGroupCount runs past it and
        ' SuffixTimeTailThenPhone stops finding the phone.
        ' (Round 47) C-H1: the offset atom now allows Z after the
        ' sign run ("10:30.5 - Z 9") - a sign followed by Z is
        ' still the time's own tail, like the numeric offsets.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?\.\d+(?:\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:(?:\d{1,2}:\d{2}|\d{4}|\d{1,2})|[Zz])))?(?:\s+\d{1,9})*\s*$") Then Return True
        ' Bare-hour form: HH + a dot tail of at most 3 digits ("10.99").
        ' A longer tail ("1.1012345678") is a decimal value, not a time:
        ' the 1-3-digit bound is the fractional-seconds bound IsTimeLike
        ' already uses, and a longer bare-hour fraction never needs to
        ' be a time tail (its digits cannot re-slice a date into a
        ' phantom phone - the suffix never completes on them).
        If Regex.IsMatch(text, "^\s*\d{1,2}\.\d{1,3}\s*$") Then Return True
        Return False
    End Function

    ''' <summary>
    ''' (Round 47) C-L1: the BARE dot-tail time only - a colon
    ''' time plus a dot tail and NOTHING after it ("10:30:45.1234",
    ''' "10:30.123456"). Used only by the CloseCandidate Round-40
    ''' drop below: that drop was designed for bare dot-tail times
    ''' whose digits concatenate into a phantom phone, but
    ''' IsTimeWithDotTail now also matches offset/trailing shapes
    ''' ("10:30.5 - 02:00 9") - those must count invalid like
    ''' their sibling families, not vanish entirely.
    ''' </summary>
    Private Function IsBareTimeWithDotTail(text As String) As Boolean
        Return Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?\.\d+\s*$")
    End Function

    ''' <summary>
    ''' Range-free shape version of IsTimeLike: True when rest has the
    ''' SHAPE of a time (bare hour, HH:MM[:SS], "HH MM"[:SS], or a time
    ''' followed by a short leftover) regardless of the values, so a
    ''' malformed time ("24:30", "10:30 99") still counts as a time
    ''' tail, not as the start of a phone. Used only by the H1/H2
    ''' date-reading guard (FormsDateTimeReading). (Round 40) C-M1:
    ''' a time followed by a dot-digit tail of any length
    ''' ("10:30.5", "10:30.123", "10.99") is a time tail too -
    ''' without this the date guard drops and the month/day groups
    ''' re-slice into a phantom valid phone. The space-separated
    ''' form ("10 30.123456") is a guard-only arm here, not part
    ''' of IsTimeWithDotTail (see its comment).
    ''' </summary>
    ''' <summary>
    ''' (Round 44) C-H2/C-H3: True when text is a colon-form time
    ''' (HH:MM[:SS], or a space-hour with colon seconds "10 30:45")
    ''' followed by a timezone offset/Z and/or one or more
    ''' space-separated trailing digit groups ("10:30-02:00",
    ''' "10:30:45 1234", "10:30+02:00 99", "10 30:45-02:00").
    ''' Range-free, like IsTimeShaped: a malformed time is still
    ''' a time tail. At least one of the offset or a trailing
    ''' group is required, so a bare time ("10:30") never matches
    ''' (it is already invalid by digit count). A space-separated
    ''' time with NO colon ("10 30 123456") is NOT covered: that
    ''' shape is a valid space-split phone spelling. The offset
    ''' is deliberately loose ([+-] + digits/colons, or Z): a
    ''' malformed offset is still the time's own tail. The
    ''' comma-fraction direct-valid residual ("10:30:45,1234")
    ''' is untouched: it has no offset and no space-separated
    ''' trailing group, so neither arm matches it. Used by
    ''' NormalizePhone only, so the extractor's date guards
    ''' (which must keep recognising a phone after a time) are
    ''' unaffected.
    ''' </summary>
    Private Function IsTimeWithOffsetOrTrailingTail(text As String) As Boolean
        Dim timePart As String = "^\s*(?:\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?|\d{1,2}\s+\d{2}\s*:\s*\d{2})"
        ' (Round 45) C-H1: the offset sign may be spaced from its digits
        ' ("10:30 - 02:00", "10:30 ++02:00") - [+-]+\s* covers a
        ' doubled sign and any space between sign and digits.
        ' (Round 46) C-H1: the sign run itself may hold spaces
        ' ("- -02:00", "- - 02:00") - the atom is now
        ' [+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]), shared by the
        ' IsTimeShaped / RemainderStartsWithTimeTail /
        ' IsPlainTimeTail arms below.
        ' (Round 47) C-H1: the offset atom now allows Z after the
        ' sign run ("10:30 - Z 99") - a sign followed by Z is
        ' still the time's own tail.
        If Regex.IsMatch(text, timePart & "\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))(?:\s+\d+)*\s*$") Then Return True
        ' (Round 46) C-H2: OFFSET-CARRYING space-form time
        ' ("10 30 - 02:00 99", "10 30-02:00 99") - the offset
        ' sign run is widened like C-H1 above. This arm is
        ' offset-carrying ONLY: a space-form time with trailing
        ' groups but NO offset ("10 30 123456") is a valid
        ' space-split phone and must NOT match here (the
        ' trailing-group-only arm below stays colon-form for
        ' that reason). Z and bare +HH/+HHMM offsets count too.
        ' (Round 47) C-H1: same Z-after-sign-run widening for
        ' the space-form arm ("10 30 - Z 99").
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))(?:\s+\d+)*\s*$") Then Return True
        If Regex.IsMatch(text, timePart & "(?:\s+\d+)+\s*$") Then Return True
        Return False
    End Function

    ''' <summary>
    ''' (Round 41) C-2: True when text is a colon time followed by a
    ''' comma-digit tail of 1-2 digits ("10:30,5", "10:30:45,12") -
    ''' a comma is a decimal separator everywhere else in this module,
    ''' so this is the time's own fractional tail, not the start of a
    ''' phone. Guard-only (like IsTimeWithDotTail's role in
    ''' IsTimeShaped): the direct path keeps its own comma behaviour
    ''' (D-1 accepted residual) and is never routed through here. A
    ''' 3+-digit tail ("10:30,123") is not a fraction and is not
    ''' covered.
    ''' </summary>
    Private Function IsTimeWithCommaTail(text As String) As Boolean
        ' (Round 46) C-H3: spaced-sign offset form only - the
        ' trailing digit group is NOT allowed here: "10:30,5
        ' - 02:00 9" stays the accepted direct-valid residual
        ' (E=0v/1i via the extractor). Strict offset shapes kept.
        ' (Round 47) C-H1: same Z-after-sign-run widening for
        ' the comma tail ("10:30,5 - Z 9").
        Return Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?,\d{1,2}(?:\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:(?:\d{1,2}:\d{2}|\d{4}|\d{1,2})|[Zz])))?\s*$")
    End Function

    ''' <summary>(Round 43) H-1: True when text STARTS with a time
    ''' with a fractional tail and an optional offset tail — canonical
    ''' ("10:30.5+02:00", "10:30:45.123+02", "10:30.5Z") or malformed
    ''' ("10:30.5+02:00:00", "10:30.5+002:00", "10:30.5+020",
    ''' "10:30.5+02000", "10:30.5+02:0"). The offset tail is
    ''' deliberately loose ([+-] + digits/colons): a malformed offset
    ''' is still the datetime's own tail, never a phone part.
    ''' (Round 44) C-H1/C-H2: a PLAIN time with an offset/Z tail
    ''' ("10:30+02:00", "10:30Z") or with a trailing digit group
    ''' ("10:30 123", "10 30 99") counts too. The Round 43
    ''' comment claiming a plain time "is not the shape that
    ''' re-slices a date into a phantom" is WRONG — "11-12-2026
    ''' 10:30+02:00" re-sliced into +201220261030 exactly
    ''' because this function missed it: the "11" closed early
    ''' and the "12-2026 10:30" remainder validated. A bare-hour
    ''' fraction ("10.5") counts too. Prefix match only (no $
    ''' anchor): anything may follow the tail.</summary>
    Private Function RemainderStartsWithTimeTail(text As String) As Boolean
        If IsTimeWithDotTail(text) OrElse IsTimeWithCommaTail(text) Then Return True
        ' (Round 47) C-H1: same Z-after-sign-run widening.
        If Regex.IsMatch(text, "^\s*(?:\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?[.,]\d+|\d{1,2}[.,]\d{1,3})\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))?") Then Return True
        ' (Round 44) plain time + offset/Z, or plain time +
        ' trailing group (colon or space form).
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        If Regex.IsMatch(text, "^\s*(?:\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?|\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?)\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))") Then Return True
        If Regex.IsMatch(text, "^\s*(?:\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?|\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?)\s+\d") Then Return True
        Return False
    End Function

    Private Function IsTimeShaped(rest As String) As Boolean
        Dim text As String = rest
        If Regex.IsMatch(text, "^\s*\d{1,2}(?:\s*:\s*\d{2}(?:\s*:\s*\d{2}(?:\.\d{1,3})?)?)?\s*$") Then Return True
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}(?:\s*:\s*\d{2}(?:\.\d{1,3})?)?\s*$") Then Return True
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2}(?:\.\d{1,3})?)?(?:\s+\d+)+\s*$") Then Return True
        ' (Round 44) C-H1/C-H2: a plain time with a timezone
        ' offset/Z ("10:30+02:00", "10:30Z", "10 30-02:00") or
        ' with any trailing digit group(s) ("10:30 123",
        ' "10 30 99", "10:30+02:00 99") is still a time tail,
        ' not the start of a phone - the fraction-only arms
        ' above left these shapes re-slicing a date into a
        ' phantom phone. The trailing-group arm above takes any
        ' length, so the IsTimeLike pattern-3 <=59 / <=2-digit
        ' cliff ("10:30 99" vs "10:30 123") is gone here. The
        ' offset is loose ([+-] + digits/colons, or Z): a
        ' malformed offset is still the datetime's own tail.
        ' Space-form arms are covered too: this helper is only
        ' consulted on the remainder after a real date, and a
        ' phone after the time is still split off later (see
        ' FormsDateTimeReading's phone-first arms and
        ' SuffixIsTimeWithDotTail, which stays colon-only).
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        ' (Round 47) C-H1: the offset atom now allows Z after the
        ' sign run ("10:30 - Z 99") - this is the arm the date
        ' guard consults, so a signed Z offset must match here
        ' or the month/day groups re-slice into a phantom phone.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))\s*$") Then Return True
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))\s*$") Then Return True
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?(?:\s+\d+)+\s*$") Then Return True
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))(?:\s+\d+)+\s*$") Then Return True
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))(?:\s+\d+)+\s*$") Then Return True
        If IsTimeWithDotTail(text) Then Return True
        ' (Round 41) C-2: comma-fraction time tail (see IsTimeWithCommaTail).
        If IsTimeWithCommaTail(text) Then Return True
        If Regex.IsMatch(text, "^\s*\d{1,2}\s+\d{2}(?:\s*:\s*\d{2})?\.\d+\s*$") Then Return True
        Return False
    End Function

    ''' <summary>
    ''' (Round 40) C-M1, prefix-aware since (Round 43) H-2: True
    ''' when a LEADING run of the digit groups from startIndex onward
    ''' reads as a time with a dot/comma tail ("10:30.123456", "10.99",
    ''' "10:30.5+02:00") — the datetime's own tail, not a phone —
    ''' WHATEVER follows it. The old full-suffix-only design never
    ''' recognised a time tail followed by a phone or junk. Used by
    ''' the complete-date branch of IsDateTimeContinuation: a suffix
    ''' that is only a time tail must not make the date close before
    ''' it (the tail's digits would re-slice the date into a phantom
    ''' phone). A time tail followed by a real phone still lets the
    ''' phone win there — see SuffixTimeTailThenPhone, consulted
    ''' alongside this function at that call site.
    ''' A space-separated time with a dot tail ("10 30.123456") is NOT
    ''' covered here (IsTimeWithDotTail needs the colon form): the D3
    ''' carve-out relies on it, so the date still closes before that
    ''' phone.
    ''' </summary>
    Private Function SuffixIsTimeWithDotTail(matches As MatchCollection, field As String, startIndex As Integer) As Boolean
        Return TimeTailGroupCount(matches, field, startIndex) > 0
    End Function

    ''' <summary>(Round 43) H-2: the number of LEADING digit groups
    ''' from startIndex that form a time with a dot/comma tail
    ''' (fractional seconds and/or a timezone offset — canonical
    ''' only, like IsTimeWithDotTail/IsTimeWithCommaTail), or 0 when
    ''' the suffix does not start with one. The LONGEST such prefix
    ''' wins ("10:30.5+02:00" is the tail, not just "10:30.5"), so a
    ''' phone search after the tail starts past the whole offset. A
    ''' phone's groups can never extend the match: the tail regexes
    ''' are $-anchored, and digits after the fraction/offset (with no
    ''' "+" / "Z" joining them) end it.</summary>
    Private Function TimeTailGroupCount(matches As MatchCollection, field As String, startIndex As Integer) As Integer
        Dim tailGroups As Integer = 0
        Dim sb As New System.Text.StringBuilder()
        Dim prevEnd As Integer = 0
        If startIndex > 0 Then prevEnd = matches(startIndex - 1).Index + matches(startIndex - 1).Length
        For j As Integer = startIndex To matches.Count - 1
            Dim mj As Match = matches(j)
            If mj.Index > prevEnd Then sb.Append(field.Substring(prevEnd, mj.Index - prevEnd))
            sb.Append(mj.Value)
            prevEnd = mj.Index + mj.Length
            ' (Round 41) C-2: a comma-fraction time tail counts the same as a dot tail here.
            ' (Round 44) C-H1/C-H2/C-H3: a colon-form plain-time
            ' tail (offset, trailing group, or both) counts too -
            ' see IsPlainTimeTail below. Space-form tails stay
            ' uncounted here (the D3 carve-out relies on it).
            If IsTimeWithDotTail(sb.ToString()) OrElse IsTimeWithCommaTail(sb.ToString()) OrElse IsPlainTimeTail(sb.ToString()) Then tailGroups = j - startIndex + 1
        Next
        Return tailGroups
    End Function

    ''' <summary>
    ''' (Round 44) C-H1/C-H2/C-H3: True when text is a COLON-form
    ''' plain time with an offset/Z tail and/or trailing digit
    ''' group(s) ("10:30+02:00", "10:30:45 1234",
    ''' "10:30-02:00 99"). Colon-form ONLY: a space-form tail
    ''' ("10 30 99", "10 30.123456") is NOT counted here - the
    ''' D3 carve-out relies on it, so a date still closes before
    ''' that phone ("11-12-2026 10 30.123456" keeps
    ''' +201030123456) and the accepted "10 12 34567.123"
    ''' phantom is untouched. Range-free, like IsTimeShaped.
    ''' The offset is loose ([+-] + digits/colons, or Z).
    ''' Used only by TimeTailGroupCount, so the date guards that
    ''' must keep recognising a space-form phone after a time
    ''' are unaffected.
    ''' </summary>
    Private Function IsPlainTimeTail(text As String) As Boolean
        ' (Round 45) C-H1: spaced/doubled offset sign - see :1069.
        ' (Round 47) C-H1: same Z-after-sign-run widening.
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?\s*(?:[Zz]|[+-]+(?:\s*[+-]+)*\s*(?:[\d:]+|[Zz]))(?:\s+\d+)*\s*$") Then Return True
        If Regex.IsMatch(text, "^\s*\d{1,2}\s*:\s*\d{2}(?:\s*:\s*\d{2})?(?:\s+\d+)+\s*$") Then Return True
        Return False
    End Function

    ''' <summary>(Round 43) H-2: True when the suffix starts with a
    ''' time tail (see TimeTailGroupCount) AND a valid phone follows
    ''' it — the date closes before that phone, exactly as the old
    ''' full-suffix design did. False when the tail is the whole
    ''' suffix or is followed by junk only.</summary>
    Private Function SuffixTimeTailThenPhone(matches As MatchCollection, field As String, startIndex As Integer) As Boolean
        Dim tailGroups As Integer = TimeTailGroupCount(matches, field, startIndex)
        If tailGroups <= 0 Then Return False
        Dim afterTail As Integer = startIndex + tailGroups
        If afterTail >= matches.Count Then Return False
        Return SuffixValidPhone(matches, field, afterTail) IsNot Nothing
    End Function

    ''' <summary>
    ''' (Round 41) C-1: True when a date-regex match is joined by
    ''' spaces only ("3 10 1234", "2026 11 16") - a weakly-marked
    ''' date whose groups may equally be a phone. A slash/dash/dot
    ''' marked date ("11/12/2026") is an explicit date, not a phone.
    ''' </summary>
    Private Function IsSpaceJoinedDate(dateMatch As Match) As Boolean
        If Not dateMatch.Success Then Return False
        Return Regex.IsMatch(dateMatch.Value, "^\d{1,2}\s+\d{1,2}\s+\d{4}$") _
            OrElse Regex.IsMatch(dateMatch.Value, "^\d{4}\s+\d{1,2}\s+\d{1,2}$")
    End Function

    ''' <summary>
    ''' (Round 41b) The year group of a date-regex match: group 1 when
    ''' the year comes first ("2026 11 16"), group 6 when it comes
    ''' last ("11 12 2026"). Only called after
    ''' DatePartsAreRealEitherOrder has accepted the match, so exactly
    ''' one of the two groups holds four digits. The 2000 fallback is
    ''' unreachable there; it falls inside the plausible window on
    ''' purpose, so an unparsed year keeps the date reading instead
    ''' of carving out a phone.
    ''' </summary>
    Private Function DateMatchYear(dateMatch As Match) As Integer
        Dim year As Integer = 0
        If dateMatch.Groups(1).Success Then
            If Integer.TryParse(dateMatch.Groups(1).Value, year) Then Return year
        ElseIf Integer.TryParse(dateMatch.Groups(6).Value, year) Then
            Return year
        End If
        Return 2000
    End Function

    ''' <summary>
    ''' H1/H2 guard for the bare-day branch of IsDateTimeContinuation:
    ''' True when current plus the digit groups from matchIndex onward
    ''' reads as a complete real date followed by nothing, a time, a
    ''' time-shaped tail, a phone, or a time followed by a phone
    ''' ("11-12-2026 10:30", "16/11/2026 10 12 345678",
    ''' "11-12-2026 10:30 01012345678"). The date match must end on a
    ''' group boundary (not in the middle of a digit run), so
    ''' "7 10 12345678" (day "12" split out of "12345678") and
    ''' "3 10 1234 5678" (tail "5678" is neither a time nor a phone)
    ''' are NOT date readings and the phone wins there.
    ''' </summary>
    Private Function FormsDateTimeReading(current As String, matches As MatchCollection, field As String, matchIndex As Integer) As Boolean
        Dim sb As New System.Text.StringBuilder()
        sb.Append(current)
        Dim prevEnd As Integer = 0
        If matchIndex > 0 Then prevEnd = matches(matchIndex - 1).Index + matches(matchIndex - 1).Length
        Dim groupEnds As New List(Of Integer)()
        For j As Integer = matchIndex To matches.Count - 1
            Dim mj As Match = matches(j)
            If mj.Index > prevEnd Then sb.Append(field.Substring(prevEnd, mj.Index - prevEnd))
            sb.Append(mj.Value)
            groupEnds.Add(sb.Length)
            prevEnd = mj.Index + mj.Length
        Next
        Dim text As String = sb.ToString()
        Dim dateMatch As Match = Regex.Match(text, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))")
        If Not dateMatch.Success OrElse Not DatePartsAreRealEitherOrder(dateMatch) Then Return False ' FIX C-D1: same US-date guard fix as IsDateTimeContinuation - a US datetime must be recognised as a date-in-progress here too.
        If dateMatch.Length < text.Length AndAlso Char.IsDigit(text(dateMatch.Length)) Then Return False
        Dim rest As String = text.Substring(dateMatch.Length).Trim().TrimStart(" "c, ":"c, "-"c, "/"c, "."c)
        If rest.Length = 0 Then Return True
        If IsTimeLike(rest) OrElse IsTimeShaped(rest) Then Return True
        ' (Round 40) C-M1: a pure-digit rest reached across a "." gap
        ' ("11/12/2026.1234", "01/02/2026.123") is a dot-digit tail
        ' on the date, not the start of a phone - the date reading
        ' holds, so the day/month groups are not re-sliced into a
        ' phantom phone. A space-separated pure-digit rest
        ' ("3 10 1234 5678") is NOT covered: there the tail may be
        ' the start of a phone, and the existing arms decide.
        If Regex.IsMatch(rest, "^\d+$") AndAlso text.Substring(dateMatch.Length).TrimStart().StartsWith(".") Then
            ' (Round 41) C-1: a space-joined "date" whose groups after
            ' the junk head read as a valid phone ("3 10 1234.5678"
            ' -> +201012345678) is the phone, not a date with a
            ' dot tail: the date reading must not hold there, or the
            ' real phone is lost. Scoped to SPACE-joined dates only:
            ' a slash/dash/dot-marked date ("11/12/2026.1234",
            ' "02/01/2026.12345") is an explicit date and the arm
            ' still holds for it, even though its groups also
            ' concatenate to a valid-looking number.
            ' (Round 41b) The carve-out is gated on the year being
            ' outside the plausible current-era window 1900-2100:
            ' the rescued family has years 1234 / 0234, which no
            ' user means as a real date, while "11 12 2026.1234"
            ' is a real date (11 Dec 2026) whose groups 2-4 also
            ' concatenate to a valid-looking number - without the
            ' gate the arm created a phantom phone +201220261234
            ' there. 1900-2100 covers every date a user of this
            ' tool could plausibly type (current era, ~75 years
            ' either side of 2026, huge margin); when the year is
            ' inside it the date reading holds, exactly like the
            ' slash/dash twins ("11/12/2026.1234",
            ' "11-12-2026.1234" -> 0 valid).
            Dim dateYear As Integer = DateMatchYear(dateMatch)
            If IsSpaceJoinedDate(dateMatch) AndAlso (dateYear < 1900 OrElse dateYear > 2100) AndAlso SuffixValidPhone(matches, field, matchIndex) IsNot Nothing Then Return False
            Return True
        End If
        Dim firstAfter As Integer = matchIndex
        For k As Integer = 0 To groupEnds.Count - 1
            If groupEnds(k) <= dateMatch.Length Then firstAfter = matchIndex + k + 1
        Next
        If firstAfter >= matches.Count Then Return False
        ' A phone directly after the date ("16/11/2026 10 12 345678").
        If SuffixValidPhone(matches, field, firstAfter) IsNot Nothing Then Return True
        ' A time followed by a phone ("11-12-2026 10:30 01012345678"):
        ' the time part is 1-3 groups, then a valid phone starts.
        ' (Round 43) H-2: a time with fractional seconds and/or a
        ' timezone offset spans MORE groups than that ("10:30:45.123"
        ' is 4 groups, "10:30.5+02:00" is 5, "10:30:45.123+02:00" is
        ' 6) — the old +2 bound stopped the scan inside the time, so
        ' the fraction/offset group occupied the t+1 slot, no phone
        ' was found after it, and the date closed into a phantom
        ' phone. The scan is time-shape-aware now: it walks forward
        ' while the groups form a time with fraction/offset, then
        ' checks for a phone after it — and a time with a fraction
        ' followed by ANY trailing group (junk, not only a phone) is
        ' a datetime continuation too: the field merges as one
        ' datetime plus junk (0 valid) instead of re-slicing.
        For t As Integer = firstAfter To Math.Min(firstAfter + 6, matches.Count - 1)
            Dim timeText As String = field.Substring(matches(firstAfter).Index, matches(t).Index + matches(t).Length - matches(firstAfter).Index)
            If (IsTimeLike(timeText) OrElse IsTimeShaped(timeText)) AndAlso t + 1 < matches.Count Then
                If SuffixValidPhone(matches, field, t + 1) IsNot Nothing Then Return True
                ' (Round 43) H-2: fraction time + trailing junk — see
                ' the comment above. The fraction test is a dot/comma
                ' in a time-shaped text (colon form via
                ' IsTimeWithDotTail/IsTimeWithCommaTail, space form
                ' via IsTimeShaped's own dot-tail arm): a plain time
                ' needs no arm here, its junk case is already
                ' recognised by the IsTimeLike/IsTimeShaped rest
                ' check above ("10:30 99").
                If IsTimeShaped(timeText) AndAlso (timeText.Contains(".") OrElse timeText.Contains(",")) Then Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' True when the text between two digit groups is only separators
    ''' allowed inside one number: space, dot, dash, slash, parentheses,
    ''' colon (a colon joins digit groups only as the time separator
    ''' inside one field, e.g. 10:30 in a datetime; the valid-number and
    ''' fragment rules in ExtractFieldCandidates still split two separate
    ''' numbers joined by a colon).
    ''' </summary>
    Private Function GapIsIntraNumberSeparator(gap As String) As Boolean
        If String.IsNullOrEmpty(gap) Then Return True
        For Each ch As Char In gap
            If ch = " "c OrElse ch = "."c OrElse ch = "-"c OrElse ch = "/"c OrElse ch = "("c OrElse ch = ")"c OrElse ch = ":"c OrElse Char.IsWhiteSpace(ch) Then Continue For
            Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' Adds a candidate when its digit count is 10-14. A single unbroken
    ''' run of more than 14 digits, and short fragments under 10 digits,
    ''' are also returned so the caller counts them invalid instead of
    ''' silently dropping them. M3: a date-like head is checked BEFORE
    ''' the 10-14 branch - a candidate starting with a date portion
    ''' (YYYY-MM-DD or DD/MM/YYYY, optionally followed by a time-like
    ''' tail) has that portion stripped and only any remainder is
    ''' processed; when nothing remains it is not counted at all, so
    ''' "2026-11-16 10:30" counts as 0. Only shapes with a 4-digit year
    ''' and a real month (1-12) and day (1-31) count as dates, so
    ''' "12-34-5678" stays an invalid candidate, not a dropped date.
    ''' </summary>
    Private Sub CloseCandidate(candidate As String, result As List(Of String))
        ' FIX C-L2 (Round 17): a whole candidate in scientific notation
        ' with an integral value is rewritten to plain digits (same
        ' rule as ConvertScientificIfIntegral), so "1.012345678E+9"
        ' reaching the extractor directly still validates. A
        ' non-integral value is left as-is: the letter E makes
        ' NormalizePhone reject it, so the whole value counts as ONE
        ' invalid instead of splitting into a phantom valid phone
        ' ("1.012345678") plus a fragment ("9"). "1012345678.5" is
        ' untouched (no exponent marker), so the L7 rule - never round
        ' a non-integral decimal into a phantom candidate - still holds.
        ' FIX C-M3 (Round 18): the pattern also allows spaces around
        ' the exponent marker ("1.012345678 E+9"): the spaces are
        ' stripped before Decimal.TryParse, which rejects an internal
        ' space, so the spaced form converts exactly like the adjacent
        ' one ("1.012345678E+9" -> 1012345678, valid).
        Dim trimmedCandidate As String = candidate.Trim()
        ' FIX C-Low-2 (Round 38): the decimal group allows an EMPTY
        ' fraction ("(?:\.\d*)?") so a merged trailing-dot exponent
        ' ("1.E+9", "01012345678.E9") converts here exactly like the
        ' direct TryParse path does - without this the merged
        ' candidate would be rejected on its letter E and "1.E+9"
        ' would count invalid instead of valid. Plain decimals are
        ' unaffected (no E, no conversion) and "1.012345678E+9"
        ' still converts as before.
        If Regex.IsMatch(trimmedCandidate, "^\+?\d+(?:\.\d*)?\s*[Ee]\s*[+-]?\s*\d+$") Then
            Dim sciValue As Decimal = 0D
            If Decimal.TryParse(Regex.Replace(trimmedCandidate, "\s", String.Empty), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, sciValue) _
                AndAlso sciValue = Decimal.Truncate(sciValue) Then
                candidate = sciValue.ToString("0", Globalization.CultureInfo.InvariantCulture)
            End If
        End If
        Dim digitCount As Integer = CountAsciiDigits(candidate)
        If digitCount = 0 Then Return
        ' H1: a date-like head is a date ONLY when what follows it is
        ' empty or purely time-like ("2011-12-25 10:30" stays dropped).
        ' Otherwise the head is NOT a date: fall through to the normal
        ' 10-14-digit branch below, so "2011-11-223344" (a valid phone,
        ' +201111223344) is not destroyed by recursing on rest "3344".
        ' L1: a space is also a date separator ("2026 11 16 10:30");
        ' the time-like gate above keeps "2011 11 223344" valid, since
        ' rest "3344" is not time-like and falls through.
        Dim dateMatch As Match = Regex.Match(candidate, "^\s*(?:(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})|(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{4}))")
        If dateMatch.Success AndAlso DatePartsAreRealEitherOrder(dateMatch) Then
            ' H1-date: the date regex can match only the first 4 digits
            ' of a longer digit run as the "year" (e.g. "10/12/345678"
            ' matches year 3456, leaving "78") - when the match ends in
            ' the middle of a digit run the head is NOT a date at all:
            ' fall through to the normal 10-14-digit branch below.
            If dateMatch.Length < candidate.Length AndAlso Char.IsDigit(candidate(dateMatch.Length)) Then
                ' Fall through - not a date (see above).
            Else
                ' L2: strip leading separators (e.g. the ":" in
                ' "2026-11-16:30", the "-" in "2026-11-16-10:30")
                ' before matching the time tail.
                Dim rest As String = candidate.Substring(dateMatch.Length).Trim().TrimStart(" "c, ":"c, "-"c, "/"c, "."c)
                If rest.Length = 0 OrElse IsTimeLike(rest) Then Return
                ' H3: a date followed by only a 1-2 digit tail
                ' ("11-12-2026 24") is a date plus a stray short tail,
                ' not a phone: do not fall through to the 10-14-digit
                ' branch below, where the date's digits would re-slice
                ' into a phantom valid phone (+201112202624). Fall
                ' through only when the tail is >= 3 digits, so
                ' "2011-11-223344" (tail "3344") stays a valid phone
                ' and "11-12-2026 123" stays 1 invalid.
                ' (Round 39) C Defect 2: a rest that looks like an
                ' exponent tail (".E<digits>" glued onto the date by
                ' the Round 38 glue) is NOT a <=2-digit time tail:
                ' fall through to normal handling so the digit
                ' fragment still counts invalid ("2026-11-16.E9" ->
                ' 0v/1i) instead of vanishing entirely. The pre-
                ' existing ".5E9" / ".5" vanish is untouched: those
                ' rests start with a digit, so they still Return.
                If Regex.IsMatch(rest, "^[Ee][+-]?\d") Then
                    ' Fall through - exponent tail, not a time tail.
                ElseIf CountAsciiDigits(rest) <= 2 Then
                    Return
                End If
                ' L2 (removed): the old time-tail / hour-tail branches
                ' here were unreachable dead code - a date+time+phone
                ' remainder needs >=17 digits, but a candidate reaching
                ' CloseCandidate has <=14 digits; that case is handled
                ' by the >14 merge-close in ExtractFieldCandidates,
                ' which closes the datetime part before the phone group.
                ' Not a datetime at all: fall through to the normal
                ' 10-14-digit branch below (do not recurse on rest).
            End If
        End If
        ' (Round 40) C-M1: a bare time with a dot-digit tail
        ' ("10:30:45.1234", "10:30.123456") is a time, not a phone -
        ' its digits concatenate to a valid pattern, so without this
        ' it is added as a PHANTOM VALID phone. The 1-3-digit
        ' fractional-second form already drops via IsTimeLike below;
        ' this covers any-length tails, scoped to the 10-14-digit
        ' window where the phantom can arise (shorter forms keep
        ' their existing treatment, longer ones stay invalid).
        ' (Round 47) C-L1: the drop is scoped to BARE dot-tail times
        ' (see IsBareTimeWithDotTail) - an offset/trailing shape
        ' ("10:30.5 - 02:00 9", "10:30.55 - 02:00") falls through
        ' to the 10-14-digit branch below and counts invalid via
        ' NormalizePhone, like the C-H1/C-H2 sibling families.
        If digitCount >= 10 AndAlso digitCount <= 14 AndAlso IsBareTimeWithDotTail(candidate.Trim()) Then Return
        If digitCount >= 10 AndAlso digitCount <= 14 Then
            result.Add(candidate)
            Return
        End If
        If digitCount > 14 Then
            result.Add(candidate)
            Return
        End If
        ' M1: a bare time ("10:30", "10:30:45", "10 :30") is not a
        ' number and is never counted - same treatment as date-like
        ' values. Spaces around the colon are allowed (L3). The
        ' shape check alone is NOT enough: the time is range-checked
        ' via IsTimeLike (hour 0-23, minute/second 0-59), so a
        ' malformed time such as "25:99" falls through and is counted
        ' invalid instead of vanishing. (L1: no 12-hour-suffix handling
        ' here - a letter in a gap closes the candidate upstream, so no
        ' candidate ever contains a 12-hour suffix or a "T".)
        ' FIX C-L1 (Round 17): no shape pre-filter - IsTimeLike alone
        ' decides (it is already anchored and range-checked). The old
        ' colon-only shape left IsTimeLike patterns 2-3 ("10 30",
        ' "10 30:45", "10:30 45") dead code on this path, so bare
        ' space-separated / leftover times were counted invalid
        ' although the date-rest above drops the same strings as
        ' times. The malformed-time behaviour is preserved: "25:99"
        ' is not time-like, so it still falls through and counts
        ' invalid.
        ' FIX C-L1 (Round 18): a BARE hour ("10") is NOT dropped
        ' here - IsTimeLike pattern 1 matches it, so the Round 17
        ' rule silently dropped 0-23 while 24-99 still counted
        ' invalid, against this module's own contract (short digit
        ' fragments are still returned, so the caller counts them as
        ' invalid instead of silently dropping them). The time-drop
        ' now applies only when the candidate contains a colon or a
        ' space/tab, i.e. only IsTimeLike patterns 2-3 ("10 30",
        ' "10 30:45", "10:30 45") can drop it; a bare hour falls
        ' through and counts invalid. The date-rest above is
        ' untouched: a date followed by a bare hour ("2026-11-16 10")
        ' is still a datetime and still drops there. "25:99" is not
        ' time-like, so it still falls through and counts invalid.
        If Regex.IsMatch(candidate.Trim(), "[:\s]") AndAlso IsTimeLike(candidate.Trim()) Then Return
        If IsDateLike(candidate) Then Return
        result.Add(candidate)
    End Sub

    ''' <summary>
    ''' True when the text after a date head is purely time-like: a bare
    ''' hour, HH:MM[:SS], a space-separated time ("10 30"), or a time
    ''' followed by a short (1-2 digit) leftover ("10:30 45", L3).
    ''' FIX C-L3 (Round 14): an optional fractional-seconds suffix
    ''' (".123") after the seconds is accepted in all three patterns
    ''' below (and in the bare-time regex in CloseCandidate and in
    ''' IsTimeShaped) — the hour/minute/second ranges are still
    ''' checked, so "25:99.123" is still not a time.
    ''' Anything longer is not time-like, so the date head is not a
    ''' date at all and the candidate falls through to the normal
    ''' phone branch (H1). L1: ranges are validated - hour 0-23,
    ''' minute/second 0-59 - so "78", "33" or "25:99" are NOT times;
    ''' a bare hour rest is time-like only when it is &lt;= 23, and the
    ''' leftover after HH:MM (pattern 3) only when it is &lt;= 59 (M2).
    ''' Spaces around a colon are allowed ("10 :30", L3).
    ''' L1: no 12-hour-suffix or "T" handling here - it was unreachable
    ''' dead code. Candidates are built only from digit groups and a gap
    ''' containing a letter closes the candidate upstream, so no
    ''' candidate (and therefore no date-rest) ever contains a letter:
    ''' "2026-11-16t10:30" drops via the gap-close and "2026-11-16
    ''' 10:30" followed by a 12-hour suffix drops because the trailing
    ''' suffix is never appended.
    ''' </summary>
    Private Function IsTimeLike(rest As String) As Boolean
        Dim hour As Integer = 0
        Dim minute As Integer = 0
        Dim second As Integer = 0
        Dim m As Match
        Dim text As String = rest
        ' HH, HH:MM or HH:MM:SS (spaces allowed around the colons).
        m = Regex.Match(text, "^\s*(\d{1,2})(?:\s*:\s*(\d{2})(?:\s*:\s*(\d{2})(?:\.\d{1,3})?)?)?\s*$")
        If m.Success Then
            If Not Integer.TryParse(m.Groups(1).Value, hour) OrElse hour > 23 Then Return False
            If m.Groups(2).Success Then
                If Not Integer.TryParse(m.Groups(2).Value, minute) OrElse minute > 59 Then Return False
            End If
            If m.Groups(3).Success Then
                If Not Integer.TryParse(m.Groups(3).Value, second) OrElse second > 59 Then Return False
            End If
            Return True
        End If
        ' "10 30" / "10 30:45": hour and minute separated by a space.
        m = Regex.Match(text, "^\s*(\d{1,2})\s+(\d{2})(?:\s*:\s*(\d{2})(?:\.\d{1,3})?)?\s*$")
        If m.Success Then
            If Not Integer.TryParse(m.Groups(1).Value, hour) OrElse hour > 23 Then Return False
            If Not Integer.TryParse(m.Groups(2).Value, minute) OrElse minute > 59 Then Return False
            If m.Groups(3).Success Then
                If Not Integer.TryParse(m.Groups(3).Value, second) OrElse second > 59 Then Return False
            End If
            Return True
        End If
        ' "10:30 45": a time followed by a short (1-2 digit) leftover.
        ' M2: the leftover (Groups(4)) is range-checked too - it must
        ' be <= 59, otherwise this is not a time ("10:30 99" is not).
        m = Regex.Match(text, "^\s*(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2})(?:\.\d{1,3})?)?\s+(\d{1,2})\s*$")
        If m.Success Then
            If Not Integer.TryParse(m.Groups(1).Value, hour) OrElse hour > 23 Then Return False
            If Not Integer.TryParse(m.Groups(2).Value, minute) OrElse minute > 59 Then Return False
            If m.Groups(3).Success Then
                If Not Integer.TryParse(m.Groups(3).Value, second) OrElse second > 59 Then Return False
            End If
            Dim leftover As Integer = 0
            If Not Integer.TryParse(m.Groups(4).Value, leftover) OrElse leftover > 59 Then Return False
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' True when a date regex match has a real month (1-12) and day
    ''' (1-31). Groups 1-3 are YYYY-MM-DD, groups 4-6 are DD-MM-YYYY.
    ''' </summary>
    Private Function DatePartsAreReal(dateMatch As Match) As Boolean
        Dim yearText As String = String.Empty
        Dim monthText As String = String.Empty
        Dim dayText As String = String.Empty
        If dateMatch.Groups(1).Success Then
            yearText = dateMatch.Groups(1).Value
            monthText = dateMatch.Groups(2).Value
            dayText = dateMatch.Groups(3).Value
        Else
            dayText = dateMatch.Groups(4).Value
            monthText = dateMatch.Groups(5).Value
            yearText = dateMatch.Groups(6).Value
        End If
        Dim year As Integer = 0
        Dim month As Integer = 0
        Dim day As Integer = 0
        If Not Integer.TryParse(yearText, year) Then Return False
        If Not Integer.TryParse(monthText, month) Then Return False
        If Not Integer.TryParse(dayText, day) Then Return False
        Return RealDateParts(year, month, day)
    End Function

    ''' <summary>
    ''' FIX C-D2: True when a date regex match reads as a real date in
    ''' DD/MM/YYYY order OR in MM/DD/YYYY order. The fallback reading
    ''' is attempted whenever the first (DD/MM) reading fails, and
    ''' only for the day-first form (groups 4-6): a US-format date
    ''' such as 11/16/2026 (month 16 in DD/MM) is a real date (Nov 16,
    ''' 2026) and must be dropped as a date, not counted invalid. Only
    ''' the middle-number-&gt;12 case can succeed via the fallback, so no
    ''' ambiguous date changes meaning: 11/12/2026 still drops (both
    ''' readings real) and 13/16/2026 still counts invalid (both
    ''' readings fail). The YYYY-first form has no alternative order.
    ''' FIX C-L1: used by FOUR call sites — IsDateTimeContinuation,
    ''' FormsDateTimeReading, CloseCandidate and IsDateLike — so all
    ''' four stay consistent.
    ''' </summary>
    Private Function DatePartsAreRealEitherOrder(dateMatch As Match) As Boolean
        If DatePartsAreReal(dateMatch) Then Return True
        If dateMatch.Groups(1).Success Then Return False
        Dim year As Integer = 0
        Dim first As Integer = 0
        Dim second As Integer = 0
        If Not Integer.TryParse(dateMatch.Groups(6).Value, year) Then Return False
        If Not Integer.TryParse(dateMatch.Groups(4).Value, first) Then Return False
        If Not Integer.TryParse(dateMatch.Groups(5).Value, second) Then Return False
        Return RealDateParts(year, first, second)
    End Function

    ''' <summary>
    ''' True when year/month/day are a real calendar date (month 1-12,
    ''' day within the month's length - Feb 29 only in a leap year).
    ''' Shared by DatePartsAreReal and the H1/H2 date-prefix guards.
    ''' </summary>
    Private Function RealDateParts(year As Integer, month As Integer, day As Integer) As Boolean
        ' DaysInMonth throws outside 1-9999, and year 0000 is not a
        ' real date anyway. H1: validate the month length - Feb 31 (or
        ' Apr 31) is not a real date, so it must not be treated as one.
        If year < 1 OrElse year > 9999 Then Return False
        If month < 1 OrElse month > 12 Then Return False
        If day < 1 Then Return False
        Return day <= DateTime.DaysInMonth(year, month)
    End Function

    ''' <summary>
    ''' True for date-like values such as 2026-11-16 or 16/11/2026, which
    ''' are excluded from candidates entirely and never counted. Only
    ''' shapes with a 4-digit year and a real month (1-12) and day (1-31)
    ''' count: "12-34-5678" is not a date and stays an invalid candidate.
    ''' </summary>
    Private Function IsDateLike(candidate As String) As Boolean
        Dim trimmed As String = candidate.Trim()
        Dim m As Match = Regex.Match(trimmed, "^(?:(\d{4})[\s\-/.](\d{1,2})[\s\-/.](\d{1,2})|(\d{1,2})[\s\-/.](\d{1,2})[\s\-/.](\d{4}))$")
        If Not m.Success Then Return False
        Return DatePartsAreRealEitherOrder(m)
    End Function

    Private Function CountAsciiDigits(text As String) As Integer
        Dim count As Integer = 0
        For Each ch As Char In text
            If ch >= "0"c AndAlso ch <= "9"c Then count += 1
        Next
        Return count
    End Function

    ''' <summary>
    ''' True when the text is already a complete valid Egyptian mobile
    ''' number (NormalizePhone accepts it). Used while merging digit
    ''' groups so a finished number is closed before the next group.
    ''' </summary>
    Private Function IsValidCandidate(candidate As String) As Boolean
        Dim normalized As String = String.Empty
        Return NormalizePhone(candidate, normalized)
    End Function

    ' ============================================================
    ' GSM 03.38 alphabet (default + extension tables).
    ' A character needs UCS2 when it is in neither table - AND also
    ' when it IS in the extension table (^ { } \ [ ~ ] | Euro): 0x1B
    ' is the modem's abort byte, so extension characters can never be
    ' sent in GSM text mode. The same holds for Xi (U+039E) and for
    ' characters that would map to 0x1A / 0x1B. Extension characters
    ' therefore force UCS2 at 70 chars per segment instead of GSM 160,
    ' increasing the segment count and cost.
    ' Latin-1 accented characters that ARE in the GSM alphabet
    ' (e.g. e-acute, u-umlaut, C-cedilla) stay GSM, not UCS2.
    ' ============================================================

    ' Lazily built once, thread-safe by construction: no double-checked
    ' locking (which needs Volatile to be correct) on these sets.
    ' (VB has no Volatile field modifier, so Lazy(Of T) is used - the
    ' same pattern ModemSender uses for its GSM maps. Module members
    ' are implicitly shared, so no Shared keyword is needed here.)
    Private ReadOnly _gsmDefaultLazy As New Lazy(Of HashSet(Of Char))(AddressOf BuildGsmDefaultSet)
    Private ReadOnly _gsmExtensionLazy As New Lazy(Of HashSet(Of Char))(AddressOf BuildGsmExtensionSet)

    Private Function GsmDefaultSet() As HashSet(Of Char)
        Return _gsmDefaultLazy.Value
    End Function

    Private Function BuildGsmDefaultSet() As HashSet(Of Char)
        Dim built As New HashSet(Of Char)()
        Dim codes As Integer() = {
            &HA, &HD, &H20, &H21, &H22, &H23, &H24, &H25, &H26, &H27,
            &H28, &H29, &H2A, &H2B, &H2C, &H2D, &H2E, &H2F,
            &H3A, &H3B, &H3C, &H3D, &H3E, &H3F, &H40,
            &HA1, &HA3, &HA4, &HA5, &HA7,
            &HBF, &HC4, &HC5, &HC6, &HC7, &HC9, &HD1, &HD6, &HD8, &HDC, &HDF,
            &HE0, &HE4, &HE5, &HE6, &HE8, &HE9, &HEC, &HF1, &HF2, &HF6, &HF8, &HF9, &HFC,
            &H394, &H393, &H398, &H39B, &H3A0, &H3A3, &H3A6, &H3A8, &H3A9, &H5F}
        ' L-phone-2: &H39E (Xi) is NOT in the GSM default alphabet -
        ' it was removed from ModemSender's map; it stays UCS2-only
        ' via the NeedsUcs2 intercept below, not via this set.
        For Each code As Integer In codes
            built.Add(ChrW(code))
        Next
        Dim code2 As Integer
        For code2 = AscW("0"c) To AscW("9"c)
            built.Add(ChrW(code2))
        Next
        For code2 = AscW("A"c) To AscW("Z"c)
            built.Add(ChrW(code2))
        Next
        For code2 = AscW("a"c) To AscW("z"c)
            built.Add(ChrW(code2))
        Next
        Return built
    End Function

    Private Function GsmExtensionSet() As HashSet(Of Char)
        Return _gsmExtensionLazy.Value
    End Function

    Private Function BuildGsmExtensionSet() As HashSet(Of Char)
        Dim built As New HashSet(Of Char)()
        Dim codes As Integer() = {&HC, &H5B, &H5C, &H5D, &H5E, &H7B, &H7C, &H7D, &H7E, &H20AC}
        For Each code As Integer In codes
            built.Add(ChrW(code))
        Next
        Return built
    End Function

    ''' <summary>
    ''' True when the message needs UCS2 encoding: any character outside
    ''' the GSM 03.38 default alphabet (e.g. Arabic), ANY character from
    ''' the GSM extension table (^ { } \ [ ~ ] | Euro - 0x1B is the
    ''' modem's abort byte, so they can never be sent in GSM text mode),
    ''' Xi (U+039E), or any character that would map to 0x1A / 0x1B.
    ''' Extension characters therefore force UCS2 at 70 chars per segment
    ''' instead of GSM 160, increasing the segment count and cost.
    ''' SplitMessage routes through this function, so such messages always
    ''' take the 70-char UCS2 branch.
    ''' </summary>
    Public Function NeedsUcs2(message As String) As Boolean
        If String.IsNullOrEmpty(message) Then Return False
        Dim defaultSet As HashSet(Of Char) = GsmDefaultSet()
        Dim extensionSet As HashSet(Of Char) = GsmExtensionSet()
        For Each ch As Char In message
            If extensionSet.Contains(ch) Then Return True
            If ch = ChrW(&H39E) Then Return True
            Dim code As Integer = AscW(ch)
            If code = &H1A OrElse code = &H1B Then Return True
            If defaultSet.Contains(ch) Then Continue For
            Return True
        Next
        Return False
    End Function

    ''' <summary>
    ''' Number of SMS parts. A message that fits in one part
    ''' (GSM 160 septets, UCS2 70 chars) is 1; a longer message
    ''' is split for concatenation (GSM 153 septets, UCS2 67
    ''' chars per part) and the phone joins the parts into ONE
    ''' displayed message. This returns exactly the number of
    ''' parts SplitMessage produces, so the displayed count and
    ''' the number actually sent always agree.
    ''' </summary>
    Public Function SegmentCount(message As String) As Integer
        Return SplitMessage(message).Count
    End Function

    ''' <summary>
    ''' Splits a message into SMS-sized parts. A message that fits
    ''' in one part (GSM 160 septets, UCS2 70 chars) returns a
    ''' single part, sent exactly as before (no UDH). A longer
    ''' message is split for CONCATENATION (Round 56): GSM 153
    ''' septets, UCS2 67 chars per part — the 6-octet UDH takes
    ''' 6 octets (140-6=134=67 UCS2 chars) / 7 septets
    ''' (160-7=153 GSM septets) — and the phone joins the parts
    ''' into ONE displayed message. Surrogate pairs are never
    ''' split across parts.
    ''' </summary>
    Public Function SplitMessage(message As String) As List(Of String)
        Dim parts As New List(Of String)()
        If String.IsNullOrEmpty(message) Then Return parts
        If NeedsUcs2(message) Then
            ' (Round 56) Concatenation: a message longer than
            ' one part (70 chars) is split at 67 chars so the
            ' 6-octet UDH fits in the 140-octet user data.
            Dim perPart As Integer = If(message.Length <= 70, 70, 67)
            Dim i As Integer = 0
            While i < message.Length
                Dim len As Integer = Math.Min(perPart, message.Length - i)
                If len > 1 AndAlso i + len < message.Length AndAlso Char.IsHighSurrogate(message(i + len - 1)) Then
                    len -= 1
                End If
                parts.Add(message.Substring(i, len))
                i += len
            End While
        Else
            ' GSM branch: every character here costs 1 septet.
            ' Extension characters never reach this branch (NeedsUcs2
            ' is True for them), so there is no cost-2 case to handle.
            ' (Round 56) Concatenation: a message longer than
            ' one part (160 septets) is split at 153 septets so
            ' the 6-octet UDH (7 septets) fits in 160 septets.
            Dim perPart As Integer = If(message.Length <= 160, 160, 153)
            Dim sb As New System.Text.StringBuilder()
            Dim used As Integer = 0
            For Each ch As Char In message
                If used + 1 > perPart AndAlso sb.Length > 0 Then
                    parts.Add(sb.ToString())
                    sb.Clear()
                    used = 0
                End If
                sb.Append(ch)
                used += 1
            Next
            If sb.Length > 0 Then parts.Add(sb.ToString())
        End If
        Return parts
    End Function

    ''' <summary>
    ''' Encodes text as UTF-16BE hex (UCS2) for AT+CMGS in UCS2 mode.
    ''' </summary>
    Public Function ToUcs2Hex(text As String) As String
        Dim bytes As Byte() = System.Text.Encoding.BigEndianUnicode.GetBytes(text)
        Dim sb As New System.Text.StringBuilder(bytes.Length * 2)
        For Each b As Byte In bytes
            sb.Append(b.ToString("X2"))
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' (Round 56) Destination-address field for an SMS-SUBMIT
    ''' PDU: address length (digits, hex) + type 91
    ''' (international) + the digits as semi-octets (odd digit
    ''' count pads with F). Returns String.Empty unless the
    ''' digits are all numeric and start with "20" (Round 51
    ''' guard — see BuildSmsSubmitPdu). Shared by the UCS2
    ''' builder here and the GSM builder in ModemSender, so
    ''' the international-number rule lives in one place.
    ''' </summary>
    Public Function BuildDestinationSemiOctets(normalizedNumber As String) As String
        If String.IsNullOrEmpty(normalizedNumber) Then Return String.Empty
        Dim digits As String = normalizedNumber
        If digits.StartsWith("+", System.StringComparison.Ordinal) Then digits = digits.Substring(1)
        If digits.Length = 0 OrElse digits.Length > 255 Then Return String.Empty
        For Each ch As Char In digits
            If ch < "0"c OrElse ch > "9"c Then Return String.Empty
        Next
        ' (Round 51) The DA type below is 91 (international): only
        ' accept digits that are actually international for this
        ' app's Egyptian domain, i.e. starting with "20"
        ' (NormalizePhone only ever outputs +20 numbers). A national
        ' number (010...) or a 00-prefixed number would otherwise be
        ' mislabelled international and rejected or misrouted by the
        ' network; returning Empty fails the row loudly via the
        ' PDU-build-failure path in ModemSender.
        If Not digits.StartsWith("20", System.StringComparison.Ordinal) Then Return String.Empty
        Dim swapped As New System.Text.StringBuilder(digits.Length + 1)
        Dim i As Integer = 0
        While i < digits.Length
            Dim second As Char = If(i + 1 < digits.Length, digits(i + 1), "F"c)
            swapped.Append(second)
            swapped.Append(digits(i))
            i += 2
        End While
        Return digits.Length.ToString("X2") & "91" & swapped.ToString()
    End Function

    ''' <summary>
    ''' (Round 50) Builds a single-segment SMS-SUBMIT PDU (hex) for
    ''' sending a UCS2 message in PDU mode (AT+CMGF=0): the Huawei
    ''' E1550 firmware rejects the UCS2 TEXT-mode AT+CMGS (live:
    ''' Arabic failed in 1s while English worked), so Arabic is sent
    ''' this way instead. SMSC = 00 (SIM default), first octet 01
    ''' (SMS-SUBMIT, no validity period, no status report), message
    ''' reference 00; destination is the international number as
    ''' semi-octets (odd digit count pads with F); PID 00, DCS 08
    ''' (UCS2); user data is the segment as UTF-16BE hex, its length
    ''' counted in octets. (Round 56) When udhTotal > 1 the segment
    ''' is one part of a CONCATENATED message: first octet 41
    ''' (UDHI set) and a 6-octet header 05 00 03 <ref> <total>
    ''' <seq> (8-bit reference, same for every part of the
    ''' message) prepended to the user data; UDL counts the
    ''' header octets too. A call without UDH arguments builds
    ''' exactly the old single-segment PDU (first octet 01).
    ''' Returns String.Empty when the number is not all digits
    ''' (after one leading '+'), the digits do not start with
    ''' "20" (Round 51: only Egyptian international numbers are
    ''' accepted - a national 010... or 00-prefixed number must
    ''' never be labelled type 91), the segment is empty, a UDH
    ''' parameter is out of range (ref 1..255, seq 1..total),
    ''' or either length would not fit its one-octet field.
    ''' </summary>
    Public Function BuildSmsSubmitPdu(normalizedNumber As String, segment As String, Optional udhRef As Integer = 0, Optional udhTotal As Integer = 0, Optional udhSeq As Integer = 0) As String
        If String.IsNullOrEmpty(segment) OrElse String.IsNullOrEmpty(normalizedNumber) Then Return String.Empty
        Dim useUdh As Boolean = (udhTotal > 1)
        If useUdh AndAlso (udhRef < 1 OrElse udhRef > 255 OrElse udhSeq < 1 OrElse udhSeq > udhTotal OrElse udhTotal > 255) Then Return String.Empty
        Dim daField As String = BuildDestinationSemiOctets(normalizedNumber)
        If daField = String.Empty Then Return String.Empty
        Dim udHex As String = ToUcs2Hex(segment)
        Dim udOctets As Integer = If(useUdh, 6, 0) + (udHex.Length \ 2)
        If udOctets = 0 OrElse udOctets > 255 Then Return String.Empty
        Dim sb As New System.Text.StringBuilder()
        sb.Append("00") ' SMSC length 0: use the SIM's default SMSC.
        sb.Append(If(useUdh, "41", "01")) ' SMS-SUBMIT (41 = UDHI set: concatenated part).
        sb.Append("00") ' Message reference (the modem assigns one).
        sb.Append(daField)
        sb.Append("00") ' Protocol identifier.
        sb.Append("08") ' Data coding scheme: UCS2.
        sb.Append(udOctets.ToString("X2")) ' User data length, in octets (UDH included).
        If useUdh Then
            sb.Append("050003")
            sb.Append(udhRef.ToString("X2"))
            sb.Append(udhTotal.ToString("X2"))
            sb.Append(udhSeq.ToString("X2"))
        End If
        sb.Append(udHex)
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' (Round 50) The AT+CMGS length for a PDU: the number of TPDU
    ''' octets, i.e. all PDU octets except the SMSC length octet.
    ''' Returns -1 for an empty or odd-length hex string.
    ''' </summary>
    Public Function SmsSubmitTpduOctets(pduHex As String) As Integer
        If String.IsNullOrEmpty(pduHex) OrElse (pduHex.Length Mod 2) <> 0 Then Return -1
        Return (pduHex.Length \ 2) - 1
    End Function

End Module
