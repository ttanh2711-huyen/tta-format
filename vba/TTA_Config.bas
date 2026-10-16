Attribute VB_Name = "TTA_Config"
' TTA-Format cho Word 2007 tro ve sau: cau hinh sinh tu dong tu ho so chuan cua add-in. KHONG sua tay;
' chay lai: UPDATE_VBA=1 npx jest src/vba
Option Explicit

Public Const TTA_STANDARD_COUNT As Integer = 2

Public Function CfgName(ByVal std As Integer) As String
    Select Case std
        Case 1: CfgName = U("Ch\u00EDnh quy\u1EC1n (N\u0110 30)")
        Case 2: CfgName = U("\u0110\u1EA3ng (HD 05)")
        Case Else: CfgName = ""
    End Select
End Function

Public Function CfgPrefix(ByVal std As Integer) As String
    Select Case std
        Case 1: CfgPrefix = "ND30_"
        Case 2: CfgPrefix = "HD05_"
        Case Else: CfgPrefix = ""
    End Select
End Function

Public Function CfgKind(ByVal std As Integer) As Integer
    Select Case std
        Case 1: CfgKind = 1
        Case 2: CfgKind = 2
        Case Else: CfgKind = 1
    End Select
End Function

Public Function CfgTitle(ByVal std As Integer) As String
    Select Case std
        Case 1: CfgTitle = U("C\u1ED8NG H\u00D2A X\u00C3 H\u1ED8I CH\u1EE6 NGH\u0128A VI\u1EC6T NAM")
        Case 2: CfgTitle = U("\u0110\u1EA2NG C\u1ED8NG S\u1EA2N VI\u1EC6T NAM")
        Case Else: CfgTitle = ""
    End Select
End Function

Public Function CfgMotto(ByVal std As Integer) As String
    Select Case std
        Case 1: CfgMotto = U("\u0110\u1ED9c l\u1EADp - T\u1EF1 do - H\u1EA1nh ph\u00FAc")
        Case 2: CfgMotto = ""
        Case Else: CfgMotto = ""
    End Select
End Function

Public Function CfgFont(ByVal std As Integer) As String
    Select Case std
        Case 1: CfgFont = "Times New Roman"
        Case 2: CfgFont = "Times New Roman"
        Case Else: CfgFont = "Times New Roman"
    End Select
End Function

Public Function CfgHideFirst(ByVal std As Integer) As Boolean
    Select Case std
        Case 1: CfgHideFirst = True
        Case 2: CfgHideFirst = True
        Case Else: CfgHideFirst = True
    End Select
End Function

Public Function CfgPageNumberSize(ByVal std As Integer) As Single
    Select Case std
        Case 1: CfgPageNumberSize = 14
        Case 2: CfgPageNumberSize = 14
        Case Else: CfgPageNumberSize = 13
    End Select
End Function

Public Sub CfgPage(ByVal std As Integer, ByRef topMm As Double, ByRef bottomMm As Double, ByRef leftMm As Double, ByRef rightMm As Double, ByRef headerMm As Double, ByRef footerMm As Double)
    Select Case std
        Case 1: topMm = 20: bottomMm = 20: leftMm = 30: rightMm = 20: headerMm = 12.5: footerMm = 5.3
        Case 2: topMm = 20: bottomMm = 20: leftMm = 30: rightMm = 15: headerMm = 10: footerMm = 12.5
    End Select
End Sub

Public Sub CfgMarginRange(ByVal std As Integer, ByVal side As String, ByRef minMm As Double, ByRef maxMm As Double)
    Select Case std & "|" & side
        Case "1|top": minMm = 20: maxMm = 25
        Case "1|bottom": minMm = 20: maxMm = 25
        Case "1|left": minMm = 30: maxMm = 35
        Case "1|right": minMm = 15: maxMm = 20
        Case "2|top": minMm = 20: maxMm = 20
        Case "2|bottom": minMm = 20: maxMm = 20
        Case "2|left": minMm = 30: maxMm = 30
        Case "2|right": minMm = 15: maxMm = 15
    End Select
End Sub

Public Sub CfgBody(ByVal std As Integer, ByRef sizeMin As Single, ByRef sizeMax As Single, ByRef indentMin As Double, ByRef indentMax As Double, ByRef lineKind As Integer, ByRef lineMin As Double, ByRef lineMax As Double)
    Select Case std
        Case 1: sizeMin = 13: sizeMax = 14: indentMin = 9.9: indentMax = 12.8: lineKind = 1: lineMin = 1: lineMax = 1.5
        Case 2: sizeMin = 14: sizeMax = 15: indentMin = 9: indentMax = 11: lineKind = 2: lineMin = 18: lineMax = 22
    End Select
End Sub

Public Function CfgStyleCount(ByVal std As Integer) As Integer
    Select Case std
        Case 1: CfgStyleCount = 15
        Case 2: CfgStyleCount = 11
        Case Else: CfgStyleCount = 0
    End Select
End Function

Public Sub CfgStyle(ByVal std As Integer, ByVal i As Integer, ByRef styleName As String, ByRef sizePt As Single, ByRef isBold As Boolean, ByRef isItalic As Boolean, ByRef isUnderline As Boolean, ByRef align As Integer, ByRef firstIndentPt As Single, ByRef beforePt As Single, ByRef afterPt As Single, ByRef lineKind As Integer, ByRef lineValue As Single, ByRef keepNext As Boolean)
    Select Case std * 100 + i
        Case 101: styleName = "ND30_NoiDung": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 102: styleName = "ND30_KinhGui": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 12: afterPt = 12: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 103: styleName = "ND30_CanCu": sizePt = 14: isBold = False: isItalic = True: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 104: styleName = "ND30_PhanChuong": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 105: styleName = "ND30_TieuDePhanChuong": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 106: styleName = "ND30_MucTieuMuc": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 107: styleName = "ND30_TieuDeMucTieuMuc": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 108: styleName = "ND30_Dieu": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 109: styleName = "ND30_Khoan": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 110: styleName = "ND30_KhoanCoTieuDe": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 111: styleName = "ND30_Diem": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 36: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 112: styleName = "ND30_PhuLucSo": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 113: styleName = "ND30_PhuLucTieuDe": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = True
        Case 114: styleName = "ND30_PhuLucChiDan": sizePt = 14: isBold = False: isItalic = True: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 1: lineValue = 1.17: keepNext = False
        Case 115: styleName = "ND30_BangDuLieu": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 0: firstIndentPt = 0: beforePt = 3: afterPt = 3: lineKind = 1: lineValue = 1: keepNext = False
        Case 201: styleName = "HD05_NoiDung": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 28.35: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = False
        Case 202: styleName = "HD05_KinhGuiTu": sizePt = 14: isBold = False: isItalic = True: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = False
        Case 203: styleName = "HD05_KinhGuiDanhSach": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 0: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = False
        Case 204: styleName = "HD05_PhanChuong": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = True
        Case 205: styleName = "HD05_TenPhanChuong": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = True
        Case 206: styleName = "HD05_MucSo": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = True
        Case 207: styleName = "HD05_TenMuc": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 1: firstIndentPt = 0: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = True
        Case 208: styleName = "HD05_Dieu": sizePt = 14: isBold = True: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 28.35: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = True
        Case 209: styleName = "HD05_Khoan": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 28.35: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = False
        Case 210: styleName = "HD05_Diem": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 3: firstIndentPt = 28.35: beforePt = 6: afterPt = 6: lineKind = 2: lineValue = 20: keepNext = False
        Case 211: styleName = "HD05_BangDuLieu": sizePt = 14: isBold = False: isItalic = False: isUnderline = False: align = 0: firstIndentPt = 0: beforePt = 3: afterPt = 3: lineKind = 1: lineValue = 1: keepNext = False
    End Select
End Sub

Public Function CfgRoleStyle(ByVal std As Integer, ByVal role As String) As String
    Select Case std & "|" & role
        Case "1|chuong": CfgRoleStyle = "ND30_PhanChuong"
        Case "1|tenChuong": CfgRoleStyle = "ND30_TieuDePhanChuong"
        Case "1|muc": CfgRoleStyle = "ND30_MucTieuMuc"
        Case "1|tenMuc": CfgRoleStyle = "ND30_TieuDeMucTieuMuc"
        Case "1|dieu": CfgRoleStyle = "ND30_Dieu"
        Case "1|deMuc": CfgRoleStyle = "ND30_NoiDung"
        Case "1|khoan": CfgRoleStyle = "ND30_Khoan"
        Case "1|diem": CfgRoleStyle = "ND30_Diem"
        Case "1|y": CfgRoleStyle = "ND30_NoiDung"
        Case "1|noiDung": CfgRoleStyle = "ND30_NoiDung"
        Case "2|chuong": CfgRoleStyle = "HD05_PhanChuong"
        Case "2|tenChuong": CfgRoleStyle = "HD05_TenPhanChuong"
        Case "2|muc": CfgRoleStyle = "HD05_MucSo"
        Case "2|tenMuc": CfgRoleStyle = "HD05_TenMuc"
        Case "2|dieu": CfgRoleStyle = "HD05_Dieu"
        Case "2|deMuc": CfgRoleStyle = "HD05_NoiDung"
        Case "2|khoan": CfgRoleStyle = "HD05_Khoan"
        Case "2|diem": CfgRoleStyle = "HD05_Diem"
        Case "2|y": CfgRoleStyle = "HD05_NoiDung"
        Case "2|noiDung": CfgRoleStyle = "HD05_NoiDung"
        Case Else: CfgRoleStyle = CfgPrefix(std) & "NoiDung"
    End Select
End Function

Public Function CfgRoleBold(ByVal std As Integer, ByVal role As String) As Boolean
    Select Case std & "|" & role
        Case "1|deMuc": CfgRoleBold = True
        Case "2|deMuc": CfgRoleBold = True
        Case Else: CfgRoleBold = False
    End Select
End Function

Public Function CfgSize(ByVal std As Integer, ByVal key As String) As Single
    Select Case std & "|" & key
        Case "1|title": CfgSize = 13
        Case "1|motto": CfgSize = 14
        Case "1|parent": CfgSize = 13
        Case "1|issuer": CfgSize = 13
        Case "1|so": CfgSize = 13
        Case "1|date": CfgSize = 13
        Case "1|vv": CfgSize = 13
        Case "1|docType": CfgSize = 14
        Case "1|subject": CfgSize = 14
        Case "1|authority": CfgSize = 14
        Case "1|signTitle": CfgSize = 14
        Case "1|signName": CfgSize = 14
        Case "1|recipientsLabel": CfgSize = 12
        Case "1|recipientsList": CfgSize = 11
        Case "2|title": CfgSize = 15
        Case "2|motto": CfgSize = 0
        Case "2|parent": CfgSize = 14
        Case "2|issuer": CfgSize = 14
        Case "2|so": CfgSize = 14
        Case "2|date": CfgSize = 14
        Case "2|vv": CfgSize = 12
        Case "2|docType": CfgSize = 16
        Case "2|subject": CfgSize = 14
        Case "2|authority": CfgSize = 14
        Case "2|signTitle": CfgSize = 14
        Case "2|signName": CfgSize = 14
        Case "2|recipientsLabel": CfgSize = 14
        Case "2|recipientsList": CfgSize = 12
        Case Else: CfgSize = 13
    End Select
End Function
