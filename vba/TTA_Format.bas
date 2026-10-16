Attribute VB_Name = "TTA_Format"
' TTA-Format cho Word 2007 tro ve sau. Sinh tu dong tu ma add-in (UPDATE_VBA=1 npx jest src/vba), khong sua tay.
' Macro: TTA_ChonChuan, TTA_ThietLapTrang, TTA_CaiStyle, TTA_DinhDangNoiDung, TTA_DanhSoTrang,
' TTA_KhoiDau, TTA_KhoiCuoi, TTA_KiemTra.
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

Private Const TTA_APP As String = "TTA-Format"
Private Const MAX_ISSUES As Long = 500

' Giai ma chuoi \uXXXX (ma nguon chi co ky tu ASCII).
Public Function U(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            out = out & ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

' Hop thong bao hien duoc chu co dau.
Private Function Msg(ByVal text As String, Optional ByVal flags As Long = 64) As Long
    Msg = MessageBoxW(0, StrPtr(text), StrPtr(TTA_APP), flags Or &H12000)
End Function

Private Function CurrentStd() As Integer
    Dim v As Integer
    v = Val(GetSetting(TTA_APP, "Options", "Standard", "1"))
    If v < 1 Or v > TTA_STANDARD_COUNT Then v = 1
    CurrentStd = v
End Function

Private Function NoDocument() As Boolean
    If Documents.Count = 0 Then
        Msg U("H\u00E3y m\u1EDF m\u1ED9t v\u0103n b\u1EA3n tr\u01B0\u1EDBc.")
        NoDocument = True
    End If
End Function

' ---------- Chon chuan ----------

Public Sub TTA_ChonChuan()
    Dim answer As Long
    answer = Msg(U("Ch\u1ECDn chu\u1EA9n tr\u00ECnh b\u00E0y:") & vbCrLf & vbCrLf & U("C\u00F3 (Yes): ") & CfgName(1) & vbCrLf & U("Kh\u00F4ng (No): ") & CfgName(2) & vbCrLf & vbCrLf & U("\u0110ang d\u00F9ng: ") & CfgName(CurrentStd()), 3 + 32)
    If answer = 6 Then SaveSetting TTA_APP, "Options", "Standard", "1"
    If answer = 7 Then SaveSetting TTA_APP, "Options", "Standard", "2"
    If answer = 6 Or answer = 7 Then Msg U("\u0110\u00E3 ch\u1ECDn: ") & CfgName(CurrentStd())
End Sub

' ---------- Thiet lap trang ----------

Public Sub TTA_ThietLapTrang()
    Dim std As Integer, s As Section, n As Long
    Dim t As Double, b As Double, l As Double, r As Double, h As Double, f As Double
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    CfgPage std, t, b, l, r, h, f
    For Each s In ActiveDocument.Sections
        With s.PageSetup
            .PaperSize = wdPaperA4
            .TopMargin = MillimetersToPoints(t)
            .BottomMargin = MillimetersToPoints(b)
            .LeftMargin = MillimetersToPoints(l)
            .RightMargin = MillimetersToPoints(r)
            .HeaderDistance = MillimetersToPoints(h)
            .FooterDistance = MillimetersToPoints(f)
        End With
        n = n + 1
    Next s
    Msg U("\u0110\u00E3 \u0111\u1EB7t kh\u1ED5 A4 cho ") & n & U(" ph\u1EA7n, l\u1EC1 tr\u00EAn ") & t & U(" mm, d\u01B0\u1EDBi ") & b & U(" mm, tr\u00E1i ") & l & U(" mm, ph\u1EA3i ") & r & U(" mm; \u0111\u1EA7u trang ") & h & U(" mm, ch\u00E2n trang ") & f & " mm (" & CfgName(std) & ")."
End Sub

' ---------- Style ----------

Public Sub TTA_CaiStyle()
    Dim std As Integer, n As Integer
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    n = InstallStyles(std)
    Msg U("\u0110\u00E3 c\u00E0i ") & n & " style " & CfgPrefix(std) & "... theo " & CfgName(std) & U(" v\u00E0o v\u0103n b\u1EA3n \u0111ang m\u1EDF.")
End Sub

Private Function InstallStyles(ByVal std As Integer) As Integer
    Dim i As Integer, st As Style
    Dim nm As String, sz As Single, bo As Boolean, it As Boolean, un As Boolean, al As Integer
    Dim fi As Single, sb As Single, sa As Single, lk As Integer, lv As Single, kn As Boolean
    For i = 1 To CfgStyleCount(std)
        CfgStyle std, i, nm, sz, bo, it, un, al, fi, sb, sa, lk, lv, kn
        Set st = GetOrAddStyle(nm)
        With st.Font
            .Name = CfgFont(std)
            .Size = sz
            .Bold = bo
            .Italic = it
            If un Then .Underline = wdUnderlineSingle Else .Underline = wdUnderlineNone
        End With
        With st.ParagraphFormat
            .Alignment = al
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = fi
            .SpaceBefore = sb
            .SpaceAfter = sa
            If lk = 2 Then
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = lv
            Else
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = LinesToPoints(lv)
            End If
            .KeepWithNext = kn
        End With
    Next i
    InstallStyles = CfgStyleCount(std)
End Function

Private Function GetOrAddStyle(ByVal styleName As String) As Style
    On Error Resume Next
    Set GetOrAddStyle = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    If GetOrAddStyle Is Nothing Then Set GetOrAddStyle = ActiveDocument.Styles.Add(styleName, wdStyleTypeParagraph)
End Function

Private Function StyleExists(ByVal styleName As String) As Boolean
    Dim st As Style
    On Error Resume Next
    Set st = ActiveDocument.Styles(styleName)
    StyleExists = Not st Is Nothing
End Function

' ---------- Dinh dang noi dung theo vai tro doan ----------

Private Function ParagraphText(ByVal p As Paragraph) As String
    ParagraphText = Trim$(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr$(7), ""), vbTab, " "))
End Function

Private Function Rx(ByVal pattern As String, ByVal s As String) As Boolean
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = False
    Rx = re.Test(s)
End Function

Private Function IsUpperLine(ByVal t As String) As Boolean
    IsUpperLine = (t = UCase$(t)) And (LCase$(t) <> t)
End Function

' Chu cai dau tien trong doan co viet hoa khong.
Private Function FirstLetterUpper(ByVal t As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If LCase$(c) <> UCase$(c) Then
            FirstLetterUpper = (c = UCase$(c))
            Exit Function
        End If
    Next i
End Function

' Vai tro doan, cung quy tac voi add-in: chuong, ten chuong, muc, ten muc, dieu, de muc La Ma, khoan, diem, y, noi dung.
Private Function DetectRole(ByVal t As String, ByVal prev As String) As String
    Dim L As String, roman As String, nb As String
    L = U("[A-Za-z\u00C0-\u1EF9]")
    roman = "[IVXLC]+"
    nb = U("(?![A-Za-z\u00C0-\u1EF90-9])")
    If Rx(U("^(Ch\u01B0\u01A1ng|CH\u01AF\u01A0NG)\s+(") & roman & "|\d+)" & nb, t) Or Rx(U("^(Ph\u1EA7n|PH\u1EA6N)\s+(th\u1EE9\s+") & L & U("+|TH\u1EE8\s+") & L & "+|" & roman & "|\d+)" & nb, t) Then
        DetectRole = "chuong"
    ElseIf Rx(U("^(M\u1EE5c|M\u1EE4C)\s+(") & roman & "|\d+)" & nb, t) Then
        DetectRole = "muc"
    ElseIf Rx(U("^(\u0110i\u1EC1u|\u0110I\u1EC0U)\s+\d+[a-z\u0111]?\."), t) Then
        DetectRole = "dieu"
    ElseIf Rx("^" & roman & "\.\s+" & L, t) Then
        DetectRole = "deMuc"
    ElseIf Rx("^\d{1,2}\.\d{1,2}\.\s+" & L, t) Then
        DetectRole = "diem"
    ElseIf Rx("^\d{1,2}\.\d{1,2}\s+" & L, t) And FirstLetterUpper(t) Then
        DetectRole = "diem"
    ElseIf Rx("^\d{1,2}\.\s+" & L, t) Then
        DetectRole = "khoan"
    ElseIf Rx(U("^[a-z\u0111]\)\s"), t) Then
        DetectRole = "diem"
    ElseIf Rx(U("^[-+\u2013\u2022]\s"), t) Then
        DetectRole = "y"
    ElseIf IsUpperLine(t) And prev = "chuong" Then
        DetectRole = "tenChuong"
    ElseIf IsUpperLine(t) And prev = "muc" Then
        DetectRole = "tenMuc"
    Else
        DetectRole = "noiDung"
    End If
End Function

' Doan duoc ap style: chua dung style cua chuan nao, khong phai Heading, khong phai danh sach tu dong.
Private Function CanApply(ByVal p As Paragraph) As Boolean
    Dim nm As String, i As Integer
    nm = p.Style.NameLocal
    For i = 1 To TTA_STANDARD_COUNT
        If Left$(nm, Len(CfgPrefix(i))) = CfgPrefix(i) Then Exit Function
    Next i
    If p.OutlineLevel <> wdOutlineLevelBodyText Then Exit Function
    If p.Range.ListFormat.ListType <> wdListNoNumbering Then Exit Function
    CanApply = True
End Function

Public Sub TTA_DinhDangNoiDung()
    Dim std As Integer, rng As Range, p As Paragraph, t As String, prev As String, role As String, nm As String
    Dim done As Long, skipped As Long
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    If Selection.Type = wdSelectionNormal And Len(Selection.Range.Text) > 1 Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If
    Application.ScreenUpdating = False
    For Each p In rng.Paragraphs
        t = ParagraphText(p)
        If t <> "" And Not p.Range.Information(wdWithInTable) Then
            role = DetectRole(t, prev)
            prev = role
            If CanApply(p) Then
                nm = CfgRoleStyle(std, role)
                If Not StyleExists(nm) Then InstallStyles std
                p.Style = ActiveDocument.Styles(nm)
                If CfgRoleBold(std, role) Then p.Range.Font.Bold = True
                done = done + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next p
    Application.ScreenUpdating = True
    Msg U("\u0110\u00E3 \u00E1p style ") & CfgName(std) & " cho " & done & U(" \u0111o\u1EA1n (ch\u01B0\u01A1ng, m\u1EE5c, \u0111i\u1EC1u, kho\u1EA3n, \u0111i\u1EC3m, th\u00E2n b\u00E0i).") & vbCrLf & U("B\u1ECF qua ") & skipped & U(" \u0111o\u1EA1n \u0111\u00E3 c\u00F3 style c\u1EE7a chu\u1EA9n, Heading ho\u1EB7c danh s\u00E1ch t\u1EF1 \u0111\u1ED9ng; \u0111o\u1EA1n trong b\u1EA3ng kh\u00F4ng \u0111\u1ED5i.") & vbCrLf & U("Ho\u00E0n t\u00E1c: Ctrl+Z.")
End Sub

' ---------- So trang ----------

Public Sub TTA_DanhSoTrang()
    Dim std As Integer, s As Section, r As Range, n As Long
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    For Each s In ActiveDocument.Sections
        n = n + 1
        s.PageSetup.DifferentFirstPageHeaderFooter = (n = 1 And CfgHideFirst(std))
        If n = 1 Or Not s.Headers(wdHeaderFooterPrimary).LinkToPrevious Then
            Set r = s.Headers(wdHeaderFooterPrimary).Range
            r.Text = ""
            r.Collapse wdCollapseStart
            r.Fields.Add r, wdFieldPage, "", False
            Set r = s.Headers(wdHeaderFooterPrimary).Range
            r.Font.Name = CfgFont(std)
            r.Font.Size = CfgPageNumberSize(std)
            With r.ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = 0
                .FirstLineIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 0
            End With
        End If
        If n = 1 And CfgHideFirst(std) Then s.Headers(wdHeaderFooterFirstPage).Range.Text = ""
    Next s
    Msg U("\u0110\u00E3 \u0111\u00E1nh s\u1ED1 trang gi\u1EEFa l\u1EC1 tr\u00EAn, c\u1EE1 ") & CfgPageNumberSize(std) & IIf(CfgHideFirst(std), U(", kh\u00F4ng hi\u1EC7n s\u1ED1 \u1EDF trang 1."), ".")
End Sub

' ---------- Khoi dau, khoi cuoi ----------

Private Sub Fmt(ByVal p As Paragraph, ByVal sizePt As Single, ByVal isBold As Boolean, Optional ByVal isItalic As Boolean = False, Optional ByVal align As Long = 1)
    With p.Range.Font
        .Name = CfgFont(CurrentStd())
        .Size = sizePt
        .Bold = isBold
        .Italic = isItalic
        .Underline = wdUnderlineNone
    End With
    With p.Format
        .Alignment = align
        .LeftIndent = 0
        .RightIndent = 0
        .FirstLineIndent = 0
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = True
    End With
End Sub

' Bang hai cot khong vien rong bang vung chu, o dau hoac cuoi van ban.
Private Function NewTable(ByVal atStart As Boolean, ByVal leftRatio As Single) As Table
    Dim r As Range, t As Table, w As Single
    If atStart Then
        ActiveDocument.Range(0, 0).InsertParagraphBefore
        Set r = ActiveDocument.Paragraphs(1).Range
        With ActiveDocument.Sections(1).PageSetup
            w = .PageWidth - .LeftMargin - .RightMargin
        End With
    Else
        ActiveDocument.Content.InsertParagraphAfter
        Set r = ActiveDocument.Paragraphs.Last.Range
        With ActiveDocument.Sections.Last.PageSetup
            w = .PageWidth - .LeftMargin - .RightMargin
        End With
    End If
    Set t = ActiveDocument.Tables.Add(r, 1, 2)
    t.Borders.Enable = False
    t.Rows.LeftIndent = 0
    t.LeftPadding = 0
    t.RightPadding = 0
    t.Columns(1).Width = w * leftRatio
    t.Columns(2).Width = w * (1 - leftRatio)
    Set NewTable = t
End Function

Public Sub TTA_KhoiDau()
    Dim std As Integer, t As Table, c As Cell, r As Range
    If NoDocument() Then Exit Sub
    If ActiveDocument.Range(0, 0).Information(wdWithInTable) Then
        Msg U("\u0110\u1EA7u v\u0103n b\u1EA3n \u0111ang l\u00E0 b\u1EA3ng. H\u00E3y th\u00EAm m\u1ED9t d\u00F2ng tr\u1ED1ng \u1EDF \u0111\u1EA7u v\u0103n b\u1EA3n r\u1ED3i ch\u1EA1y l\u1EA1i.")
        Exit Sub
    End If
    std = CurrentStd()
    Set t = NewTable(True, 0.38)
    Set c = t.Cell(1, 1)
    If CfgKind(std) = 1 Then
        c.Range.Text = U("T\u00CAN C\u01A0 QUAN, T\u1ED4 CH\u1EE8C CH\u1EE6 QU\u1EA2N") & vbCr & U("T\u00CAN C\u01A0 QUAN, T\u1ED4 CH\u1EE8C") & vbCr & "" & vbCr & U("S\u1ED1:       /       ")
        Fmt c.Range.Paragraphs(1), CfgSize(std, "parent"), False
        Fmt c.Range.Paragraphs(2), CfgSize(std, "issuer"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "so"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "so"), False
    Else
        c.Range.Text = U("\u0110\u1EA2NG B\u1ED8 ...") & vbCr & U("CHI B\u1ED8 ...") & vbCr & "*" & vbCr & U("S\u1ED1    -.../...")
        Fmt c.Range.Paragraphs(1), CfgSize(std, "parent"), False
        Fmt c.Range.Paragraphs(2), CfgSize(std, "issuer"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "issuer"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "so"), False
    End If
    Set c = t.Cell(1, 2)
    If CfgMotto(std) <> "" Then
        c.Range.Text = CfgTitle(std) & vbCr & CfgMotto(std) & vbCr & "" & vbCr & U("......., ng\u00E0y ... th\u00E1ng ... n\u0103m ......")
        Fmt c.Range.Paragraphs(1), CfgSize(std, "title"), True
        Fmt c.Range.Paragraphs(2), CfgSize(std, "motto"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "date"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "date"), False, True
    Else
        c.Range.Text = CfgTitle(std) & vbCr & "" & vbCr & U("......., ng\u00E0y ... th\u00E1ng ... n\u0103m ......")
        Fmt c.Range.Paragraphs(1), CfgSize(std, "title"), True
        Fmt c.Range.Paragraphs(2), CfgSize(std, "date"), False
        Fmt c.Range.Paragraphs(3), CfgSize(std, "date"), False, True
    End If
    Set r = t.Range
    r.Collapse wdCollapseEnd
    r.InsertBefore vbCr & U("T\u00CAN LO\u1EA0I V\u0102N B\u1EA2N") & vbCr & U("Tr\u00EDch y\u1EBFu n\u1ED9i dung v\u0103n b\u1EA3n") & vbCr
    Fmt r.Paragraphs(1), CfgSize(std, "docType"), False
    Fmt r.Paragraphs(2), CfgSize(std, "docType"), True
    Fmt r.Paragraphs(3), CfgSize(std, "subject"), True
    Msg U("\u0110\u00E3 ch\u00E8n kh\u1ED1i \u0111\u1EA7u v\u0103n b\u1EA3n (") & CfgName(std) & U("). S\u1EEDa ch\u1EEF \u0111\u1EC3 s\u1EB5n cho \u0111\u00FAng; \u0111\u01B0\u1EDDng k\u1EBB d\u01B0\u1EDBi t\u00EAn c\u01A1 quan, ti\u00EAu ng\u1EEF, tr\u00EDch y\u1EBFu h\u00E3y v\u1EBD b\u1EB1ng Insert > Shapes > Line (ho\u1EB7c d\u00F9ng add-in tr\u00EAn Word 2016 tr\u1EDF l\u00EAn).")
End Sub

Public Sub TTA_KhoiCuoi()
    Dim std As Integer, t As Table, c As Cell
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    Set t = NewTable(False, 0.45)
    Set c = t.Cell(1, 1)
    c.Range.Text = U("N\u01A1i nh\u1EADn:") & vbCr & U("- Nh\u01B0 tr\u00EAn;") & vbCr & "- ...;" & vbCr & U("- L\u01B0u: VT, ...")
    Fmt c.Range.Paragraphs(1), CfgSize(std, "recipientsLabel"), True, True, 0
    Fmt c.Range.Paragraphs(2), CfgSize(std, "recipientsList"), False, False, 0
    Fmt c.Range.Paragraphs(3), CfgSize(std, "recipientsList"), False, False, 0
    Fmt c.Range.Paragraphs(4), CfgSize(std, "recipientsList"), False, False, 0
    Set c = t.Cell(1, 2)
    c.Range.Text = U("QUY\u1EC0N H\u1EA0N, CH\u1EE8C V\u1EE4 C\u1EE6A NG\u01AF\u1EDCI K\u00DD") & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & U("H\u1ECD v\u00E0 t\u00EAn")
    Fmt c.Range.Paragraphs(1), CfgSize(std, "authority"), True
    Fmt c.Range.Paragraphs(2), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(3), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(4), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(5), CfgSize(std, "signName"), True
    Msg U("\u0110\u00E3 ch\u00E8n kh\u1ED1i cu\u1ED1i v\u0103n b\u1EA3n (") & CfgName(std) & U(") \u1EDF cu\u1ED1i t\u00E0i li\u1EC7u. S\u1EEDa ch\u1EEF \u0111\u1EC3 s\u1EB5n cho \u0111\u00FAng.")
End Sub

' ---------- Kiem tra ----------

Private Sub AddIssue(ByRef out As String, ByRef issues As Long, ByVal line As String)
    issues = issues + 1
    If issues <= MAX_ISSUES Then out = out & "- " & line & vbCr
End Sub

Private Function Where(ByVal i As Long, ByVal t As String) As String
    If Len(t) > 40 Then t = Left$(t, 40) & "..."
    Where = U("\u0110o\u1EA1n ") & i & " (" & t & "): "
End Function

Private Function Between(ByVal mn As Double, ByVal mx As Double) As String
    If mn = mx Then Between = CStr(mn) Else Between = mn & "-" & mx
End Function

Private Function IsA4(ByVal w As Single, ByVal h As Single) As Boolean
    Dim a As Single, b As Single
    a = MillimetersToPoints(210)
    b = MillimetersToPoints(297)
    IsA4 = (Abs(w - a) < 3 And Abs(h - b) < 3) Or (Abs(w - b) < 3 And Abs(h - a) < 3)
End Function

Private Function MarginOf(ByVal ps As PageSetup, ByVal k As Integer) As Single
    Select Case k
        Case 0: MarginOf = ps.TopMargin
        Case 1: MarginOf = ps.BottomMargin
        Case 2: MarginOf = ps.LeftMargin
        Case Else: MarginOf = ps.RightMargin
    End Select
End Function

Private Function HasPageField(ByVal hf As HeaderFooter) As Boolean
    Dim f As Field
    For Each f In hf.Range.Fields
        If InStr(1, UCase$(f.Code.Text), "PAGE") > 0 Then HasPageField = True
    Next f
End Function

Private Function IsBody(ByVal p As Paragraph) As Boolean
    If p.Range.Information(wdWithInTable) Then Exit Function
    If p.OutlineLevel <> wdOutlineLevelBodyText Then Exit Function
    IsBody = (p.Alignment = wdAlignParagraphJustify)
End Function

Public Sub TTA_KiemTra()
    Dim std As Integer, src As Document, rep As Document, out As String, issues As Long
    Dim s As Section, n As Long, k As Integer, sides As Variant, labels As Variant
    Dim mn As Double, mx As Double, mm As Double
    Dim sMin As Single, sMax As Single, iMin As Double, iMax As Double, lk As Integer, lMin As Double, lMax As Double
    Dim p As Paragraph, t As String, i As Long, sz As Single, fi As Double, lines As Double, fontName As String
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    Set src = ActiveDocument
    sides = Array("top", "bottom", "left", "right")
    labels = Array(U("tr\u00EAn"), U("d\u01B0\u1EDBi"), U("tr\u00E1i"), U("ph\u1EA3i"))
    For Each s In src.Sections
        n = n + 1
        If Not IsA4(s.PageSetup.PageWidth, s.PageSetup.PageHeight) Then AddIssue out, issues, U("Ph\u1EA7n ") & n & U(": kh\u1ED5 gi\u1EA5y kh\u00F4ng ph\u1EA3i A4.")
        For k = 0 To 3
            CfgMarginRange std, sides(k), mn, mx
            mm = Round(PointsToMillimeters(MarginOf(s.PageSetup, k)), 1)
            If mm < mn - 0.5 Or mm > mx + 0.5 Then AddIssue out, issues, U("Ph\u1EA7n ") & n & U(": l\u1EC1 ") & labels(k) & " " & mm & U(" mm, c\u1EA7n ") & Between(mn, mx) & " mm."
        Next k
        If src.ComputeStatistics(wdStatisticPages) > 1 And Not HasPageField(s.Headers(wdHeaderFooterPrimary)) Then AddIssue out, issues, U("Ph\u1EA7n ") & n & U(": ch\u01B0a c\u00F3 s\u1ED1 trang \u1EDF \u0111\u1EA7u trang (ch\u1EA1y TTA_DanhSoTrang).")
    Next s

    CfgBody std, sMin, sMax, iMin, iMax, lk, lMin, lMax
    For Each p In src.Paragraphs
        i = i + 1
        t = ParagraphText(p)
        If t <> "" Then
            fontName = p.Range.Font.Name
            If fontName = "" Then
                AddIssue out, issues, Where(i, t) & U("c\u00F3 nhi\u1EC1u ph\u00F4ng ch\u1EEF, c\u1EA7n ") & CfgFont(std) & "."
            ElseIf fontName <> CfgFont(std) Then
                AddIssue out, issues, Where(i, t) & U("ph\u00F4ng ") & fontName & U(", c\u1EA7n ") & CfgFont(std) & "."
            End If
            If IsBody(p) Then
                sz = p.Range.Font.Size
                If sz = wdUndefined Then
                    AddIssue out, issues, Where(i, t) & U("c\u1EE1 ch\u1EEF kh\u00F4ng \u0111\u1ED3ng nh\u1EA5t, c\u1EA7n ") & Between(sMin, sMax) & "."
                ElseIf sz < sMin Or sz > sMax Then
                    AddIssue out, issues, Where(i, t) & U("c\u1EE1 ch\u1EEF ") & sz & U(", c\u1EA7n ") & Between(sMin, sMax) & "."
                End If
                fi = Round(PointsToMillimeters(p.FirstLineIndent), 1)
                If fi < iMin Or fi > iMax Then AddIssue out, issues, Where(i, t) & U("th\u1EE5t d\u00F2ng \u0111\u1EA7u ") & fi & U(" mm, c\u1EA7n ") & Between(iMin, iMax) & " mm."
                If lk = 1 Then
                    If p.LineSpacingRule = wdLineSpaceExactly Or p.LineSpacingRule = wdLineSpaceAtLeast Then
                        AddIssue out, issues, Where(i, t) & U("gi\u00E3n d\u00F2ng c\u1ED1 \u0111\u1ECBnh, c\u1EA7n ") & Between(lMin, lMax) & U(" d\u00F2ng.")
                    Else
                        lines = Round(p.LineSpacing / 12, 2)
                        If lines < lMin - 0.01 Or lines > lMax + 0.01 Then AddIssue out, issues, Where(i, t) & U("gi\u00E3n d\u00F2ng ") & lines & U(" d\u00F2ng, c\u1EA7n ") & Between(lMin, lMax) & U(" d\u00F2ng.")
                    End If
                ElseIf p.LineSpacingRule <> wdLineSpaceExactly Or p.LineSpacing < lMin - 0.1 Or p.LineSpacing > lMax + 0.1 Then
                    AddIssue out, issues, Where(i, t) & U("gi\u00E3n d\u00F2ng c\u1EA7n Exactly ") & Between(lMin, lMax) & " pt."
                End If
            End If
        End If
    Next p

    Set rep = Documents.Add
    rep.Content.Text = U("K\u1EBET QU\u1EA2 KI\u1EC2M TRA TH\u1EC2 TH\u1EE8C") & vbCr & U("Chu\u1EA9n: ") & CfgName(std) & vbCr & U("V\u0103n b\u1EA3n: ") & src.Name & vbCr & IIf(issues = 0, U("Kh\u00F4ng th\u1EA5y l\u1ED7i v\u1EC1 kh\u1ED5 gi\u1EA5y, l\u1EC1, s\u1ED1 trang, ph\u00F4ng, c\u1EE1 ch\u1EEF, th\u1EE5t d\u00F2ng, gi\u00E3n d\u00F2ng."), issues & U(" l\u1ED7i") & IIf(issues > MAX_ISSUES, U(" (ch\u1EC9 li\u1EC7t k\u00EA ") & MAX_ISSUES & U(" l\u1ED7i \u0111\u1EA7u)"), "") & ":") & vbCr & out & vbCr & U("B\u1ED9 macro ch\u1EC9 ki\u1EC3m tra tr\u00ECnh b\u00E0y c\u01A1 b\u1EA3n; ki\u1EC3m tra \u0111\u1EE7 th\u1EC3 th\u1EE9c (s\u1ED1, k\u00FD hi\u1EC7u, vi\u1EBFt hoa, n\u01A1i nh\u1EADn...) d\u00F9ng add-in TTA-Format tr\u00EAn Word 2016 tr\u1EDF l\u00EAn.")
    rep.Content.Font.Name = "Times New Roman"
    rep.Content.Font.Size = 13
    rep.Paragraphs(1).Range.Font.Bold = True
End Sub
