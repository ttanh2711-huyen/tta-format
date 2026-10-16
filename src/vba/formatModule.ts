/**
 * Mô-đun macro chính của bộ VBA (T29) cho Word 2007 trở về sau, viết bằng tiếng Việt có dấu; `toAsciiVba` đổi
 * thành mã ASCII khi sinh tệp .bas. Giá trị theo chuẩn lấy từ mô-đun TTA_Config (sinh từ hồ sơ chuẩn).
 * MsgBox, InputBox của VBA không hiện được chữ có dấu nên dùng MessageBoxW; chữ cần điền để sẵn trong văn bản.
 */
export const FORMAT_MODULE = String.raw`Attribute VB_Name = "TTA_Format"
' TTA-Format cho Word 2007 trở về sau. Sinh tự động từ mã add-in (UPDATE_VBA=1 npx jest src/vba), không sửa tay.
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

' Giải mã chuỗi \uXXXX (mã nguồn chỉ có ký tự ASCII).
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

' Hộp thông báo hiện được chữ có dấu.
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
        Msg "Hãy mở một văn bản trước."
        NoDocument = True
    End If
End Function

' ---------- Chọn chuẩn ----------

Public Sub TTA_ChonChuan()
    Dim answer As Long
    answer = Msg("Chọn chuẩn trình bày:" & vbCrLf & vbCrLf & "Có (Yes): " & CfgName(1) & vbCrLf & "Không (No): " & CfgName(2) & vbCrLf & vbCrLf & "Đang dùng: " & CfgName(CurrentStd()), 3 + 32)
    If answer = 6 Then SaveSetting TTA_APP, "Options", "Standard", "1"
    If answer = 7 Then SaveSetting TTA_APP, "Options", "Standard", "2"
    If answer = 6 Or answer = 7 Then Msg "Đã chọn: " & CfgName(CurrentStd())
End Sub

' ---------- Thiết lập trang ----------

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
    Msg "Đã đặt khổ A4 cho " & n & " phần, lề trên " & t & " mm, dưới " & b & " mm, trái " & l & " mm, phải " & r & " mm; đầu trang " & h & " mm, chân trang " & f & " mm (" & CfgName(std) & ")."
End Sub

' ---------- Style ----------

Public Sub TTA_CaiStyle()
    Dim std As Integer, n As Integer
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    n = InstallStyles(std)
    Msg "Đã cài " & n & " style " & CfgPrefix(std) & "... theo " & CfgName(std) & " vào văn bản đang mở."
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

' ---------- Định dạng nội dung theo vai trò đoạn ----------

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

' Chữ cái đầu tiên trong đoạn có viết hoa không.
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

' Vai trò đoạn, cùng quy tắc với add-in: chương, tên chương, mục, tên mục, điều, đề mục La Mã, khoản, điểm, ý, nội dung.
Private Function DetectRole(ByVal t As String, ByVal prev As String) As String
    Dim L As String, roman As String, nb As String
    L = "[A-Za-zÀ-ỹ]"
    roman = "[IVXLC]+"
    nb = "(?![A-Za-zÀ-ỹ0-9])"
    If Rx("^(Chương|CHƯƠNG)\s+(" & roman & "|\d+)" & nb, t) Or Rx("^(Phần|PHẦN)\s+(thứ\s+" & L & "+|THỨ\s+" & L & "+|" & roman & "|\d+)" & nb, t) Then
        DetectRole = "chuong"
    ElseIf Rx("^(Mục|MỤC)\s+(" & roman & "|\d+)" & nb, t) Then
        DetectRole = "muc"
    ElseIf Rx("^(Điều|ĐIỀU)\s+\d+[a-zđ]?\.", t) Then
        DetectRole = "dieu"
    ElseIf Rx("^" & roman & "\.\s+" & L, t) Then
        DetectRole = "deMuc"
    ElseIf Rx("^\d{1,2}\.\d{1,2}\.\s+" & L, t) Then
        DetectRole = "diem"
    ElseIf Rx("^\d{1,2}\.\d{1,2}\s+" & L, t) And FirstLetterUpper(t) Then
        DetectRole = "diem"
    ElseIf Rx("^\d{1,2}\.\s+" & L, t) Then
        DetectRole = "khoan"
    ElseIf Rx("^[a-zđ]\)\s", t) Then
        DetectRole = "diem"
    ElseIf Rx("^[-+–•]\s", t) Then
        DetectRole = "y"
    ElseIf IsUpperLine(t) And prev = "chuong" Then
        DetectRole = "tenChuong"
    ElseIf IsUpperLine(t) And prev = "muc" Then
        DetectRole = "tenMuc"
    Else
        DetectRole = "noiDung"
    End If
End Function

' Đoạn được áp style: chưa dùng style của chuẩn nào, không phải Heading, không phải danh sách tự động.
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
    Msg "Đã áp style " & CfgName(std) & " cho " & done & " đoạn (chương, mục, điều, khoản, điểm, thân bài)." & vbCrLf & "Bỏ qua " & skipped & " đoạn đã có style của chuẩn, Heading hoặc danh sách tự động; đoạn trong bảng không đổi." & vbCrLf & "Hoàn tác: Ctrl+Z."
End Sub

' ---------- Số trang ----------

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
    Msg "Đã đánh số trang giữa lề trên, cỡ " & CfgPageNumberSize(std) & IIf(CfgHideFirst(std), ", không hiện số ở trang 1.", ".")
End Sub

' ---------- Khối đầu, khối cuối ----------

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

' Bảng hai cột không viền rộng bằng vùng chữ, ở đầu hoặc cuối văn bản.
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
        Msg "Đầu văn bản đang là bảng. Hãy thêm một dòng trống ở đầu văn bản rồi chạy lại."
        Exit Sub
    End If
    std = CurrentStd()
    Set t = NewTable(True, 0.38)
    Set c = t.Cell(1, 1)
    If CfgKind(std) = 1 Then
        c.Range.Text = "TÊN CƠ QUAN, TỔ CHỨC CHỦ QUẢN" & vbCr & "TÊN CƠ QUAN, TỔ CHỨC" & vbCr & "" & vbCr & "Số:       /       "
        Fmt c.Range.Paragraphs(1), CfgSize(std, "parent"), False
        Fmt c.Range.Paragraphs(2), CfgSize(std, "issuer"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "so"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "so"), False
    Else
        c.Range.Text = "ĐẢNG BỘ ..." & vbCr & "CHI BỘ ..." & vbCr & "*" & vbCr & "Số    -.../..."
        Fmt c.Range.Paragraphs(1), CfgSize(std, "parent"), False
        Fmt c.Range.Paragraphs(2), CfgSize(std, "issuer"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "issuer"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "so"), False
    End If
    Set c = t.Cell(1, 2)
    If CfgMotto(std) <> "" Then
        c.Range.Text = CfgTitle(std) & vbCr & CfgMotto(std) & vbCr & "" & vbCr & "......., ngày ... tháng ... năm ......"
        Fmt c.Range.Paragraphs(1), CfgSize(std, "title"), True
        Fmt c.Range.Paragraphs(2), CfgSize(std, "motto"), True
        Fmt c.Range.Paragraphs(3), CfgSize(std, "date"), False
        Fmt c.Range.Paragraphs(4), CfgSize(std, "date"), False, True
    Else
        c.Range.Text = CfgTitle(std) & vbCr & "" & vbCr & "......., ngày ... tháng ... năm ......"
        Fmt c.Range.Paragraphs(1), CfgSize(std, "title"), True
        Fmt c.Range.Paragraphs(2), CfgSize(std, "date"), False
        Fmt c.Range.Paragraphs(3), CfgSize(std, "date"), False, True
    End If
    Set r = t.Range
    r.Collapse wdCollapseEnd
    r.InsertBefore vbCr & "TÊN LOẠI VĂN BẢN" & vbCr & "Trích yếu nội dung văn bản" & vbCr
    Fmt r.Paragraphs(1), CfgSize(std, "docType"), False
    Fmt r.Paragraphs(2), CfgSize(std, "docType"), True
    Fmt r.Paragraphs(3), CfgSize(std, "subject"), True
    Msg "Đã chèn khối đầu văn bản (" & CfgName(std) & "). Sửa chữ để sẵn cho đúng; đường kẻ dưới tên cơ quan, tiêu ngữ, trích yếu hãy vẽ bằng Insert > Shapes > Line (hoặc dùng add-in trên Word 2016 trở lên)."
End Sub

Public Sub TTA_KhoiCuoi()
    Dim std As Integer, t As Table, c As Cell
    If NoDocument() Then Exit Sub
    std = CurrentStd()
    Set t = NewTable(False, 0.45)
    Set c = t.Cell(1, 1)
    c.Range.Text = "Nơi nhận:" & vbCr & "- Như trên;" & vbCr & "- ...;" & vbCr & "- Lưu: VT, ..."
    Fmt c.Range.Paragraphs(1), CfgSize(std, "recipientsLabel"), True, True, 0
    Fmt c.Range.Paragraphs(2), CfgSize(std, "recipientsList"), False, False, 0
    Fmt c.Range.Paragraphs(3), CfgSize(std, "recipientsList"), False, False, 0
    Fmt c.Range.Paragraphs(4), CfgSize(std, "recipientsList"), False, False, 0
    Set c = t.Cell(1, 2)
    c.Range.Text = "QUYỀN HẠN, CHỨC VỤ CỦA NGƯỜI KÝ" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & "Họ và tên"
    Fmt c.Range.Paragraphs(1), CfgSize(std, "authority"), True
    Fmt c.Range.Paragraphs(2), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(3), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(4), CfgSize(std, "signTitle"), True
    Fmt c.Range.Paragraphs(5), CfgSize(std, "signName"), True
    Msg "Đã chèn khối cuối văn bản (" & CfgName(std) & ") ở cuối tài liệu. Sửa chữ để sẵn cho đúng."
End Sub

' ---------- Kiểm tra ----------

Private Sub AddIssue(ByRef out As String, ByRef issues As Long, ByVal line As String)
    issues = issues + 1
    If issues <= MAX_ISSUES Then out = out & "- " & line & vbCr
End Sub

Private Function Where(ByVal i As Long, ByVal t As String) As String
    If Len(t) > 40 Then t = Left$(t, 40) & "..."
    Where = "Đoạn " & i & " (" & t & "): "
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
    labels = Array("trên", "dưới", "trái", "phải")
    For Each s In src.Sections
        n = n + 1
        If Not IsA4(s.PageSetup.PageWidth, s.PageSetup.PageHeight) Then AddIssue out, issues, "Phần " & n & ": khổ giấy không phải A4."
        For k = 0 To 3
            CfgMarginRange std, sides(k), mn, mx
            mm = Round(PointsToMillimeters(MarginOf(s.PageSetup, k)), 1)
            If mm < mn - 0.5 Or mm > mx + 0.5 Then AddIssue out, issues, "Phần " & n & ": lề " & labels(k) & " " & mm & " mm, cần " & Between(mn, mx) & " mm."
        Next k
        If src.ComputeStatistics(wdStatisticPages) > 1 And Not HasPageField(s.Headers(wdHeaderFooterPrimary)) Then AddIssue out, issues, "Phần " & n & ": chưa có số trang ở đầu trang (chạy TTA_DanhSoTrang)."
    Next s

    CfgBody std, sMin, sMax, iMin, iMax, lk, lMin, lMax
    For Each p In src.Paragraphs
        i = i + 1
        t = ParagraphText(p)
        If t <> "" Then
            fontName = p.Range.Font.Name
            If fontName = "" Then
                AddIssue out, issues, Where(i, t) & "có nhiều phông chữ, cần " & CfgFont(std) & "."
            ElseIf fontName <> CfgFont(std) Then
                AddIssue out, issues, Where(i, t) & "phông " & fontName & ", cần " & CfgFont(std) & "."
            End If
            If IsBody(p) Then
                sz = p.Range.Font.Size
                If sz = wdUndefined Then
                    AddIssue out, issues, Where(i, t) & "cỡ chữ không đồng nhất, cần " & Between(sMin, sMax) & "."
                ElseIf sz < sMin Or sz > sMax Then
                    AddIssue out, issues, Where(i, t) & "cỡ chữ " & sz & ", cần " & Between(sMin, sMax) & "."
                End If
                fi = Round(PointsToMillimeters(p.FirstLineIndent), 1)
                If fi < iMin Or fi > iMax Then AddIssue out, issues, Where(i, t) & "thụt dòng đầu " & fi & " mm, cần " & Between(iMin, iMax) & " mm."
                If lk = 1 Then
                    If p.LineSpacingRule = wdLineSpaceExactly Or p.LineSpacingRule = wdLineSpaceAtLeast Then
                        AddIssue out, issues, Where(i, t) & "giãn dòng cố định, cần " & Between(lMin, lMax) & " dòng."
                    Else
                        lines = Round(p.LineSpacing / 12, 2)
                        If lines < lMin - 0.01 Or lines > lMax + 0.01 Then AddIssue out, issues, Where(i, t) & "giãn dòng " & lines & " dòng, cần " & Between(lMin, lMax) & " dòng."
                    End If
                ElseIf p.LineSpacingRule <> wdLineSpaceExactly Or p.LineSpacing < lMin - 0.1 Or p.LineSpacing > lMax + 0.1 Then
                    AddIssue out, issues, Where(i, t) & "giãn dòng cần Exactly " & Between(lMin, lMax) & " pt."
                End If
            End If
        End If
    Next p

    Set rep = Documents.Add
    rep.Content.Text = "KẾT QUẢ KIỂM TRA THỂ THỨC" & vbCr & "Chuẩn: " & CfgName(std) & vbCr & "Văn bản: " & src.Name & vbCr & IIf(issues = 0, "Không thấy lỗi về khổ giấy, lề, số trang, phông, cỡ chữ, thụt dòng, giãn dòng.", issues & " lỗi" & IIf(issues > MAX_ISSUES, " (chỉ liệt kê " & MAX_ISSUES & " lỗi đầu)", "") & ":") & vbCr & out & vbCr & "Bộ macro chỉ kiểm tra trình bày cơ bản; kiểm tra đủ thể thức (số, ký hiệu, viết hoa, nơi nhận...) dùng add-in TTA-Format trên Word 2016 trở lên."
    rep.Content.Font.Name = "Times New Roman"
    rep.Content.Font.Size = 13
    rep.Paragraphs(1).Range.Font.Bold = True
End Sub
`;
