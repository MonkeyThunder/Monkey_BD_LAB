Attribute VB_Name = "modURLEncode"
'#######################################################################
'#
'#  모듈 코딩자 : 바람개비(cjh9217) in Naver
'#  Blog        : http://gaibee.tistory.com
'#  작성 의도   : URL Encoding, Decoding을 좀더 빠르고 쉽게 하기 위함
'#
'#######################################################################

Option Explicit

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codepage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
) As Long
     
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codepage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
) As Long

Private Const CP_UTF8 As Long = 65001
Public Function ConvertUiToANSI(ByVal ConvStr As String) As String
    Dim arrTmp() As String
    arrTmp = Split(ConvStr, "\u")

    Dim Index As Long
    For Index = 1 To UBound(arrTmp)
        arrTmp(Index) = ChrW$(CLng("&H" & Left$(arrTmp(Index), 4))) & Mid$(arrTmp(Index), 5)
    Next Index

    ConvertUiToANSI = Join(arrTmp, "")
End Function


'### 한글 -> EUC-KR(인코딩)
Public Function Change(Str As String) As String
On Error GoTo ErrLbl
Dim I
For I = 1 To Len(Str)
If Len(Hex(Asc(Mid$(Str, I, 1)))) = 4 Then
Change = Change & "%" & Mid$(Hex(Asc(Mid$(Str, I, 1))), 1, 2) & "%" & Mid$(Hex(Asc(Mid$(Str, I, 1))), 3, 2)
Else
Change = Change & "%" & Hex(Asc(Mid$(Str, I, 1)))
End If
Next
ErrLbl:
End Function

'### 흔히 고구마s에서 쓰이는 Change 함수(위의 것)와 같은 종류 입니다.
'### 속도는 이게 더 빠를겁니다...
'### 한글 -> EUC-KR(인코딩)
Public Function URLEncodeAnsi(Str As String) As String
On Error GoTo ErrLbl

    Dim AnsiArr() As Byte, I As Long, Buf As String
    AnsiArr = StrConv(Str, vbFromUnicode)

    For I = 0 To UBound(AnsiArr)
        Buf = Buf & "%" & Hex$(AnsiArr(I))
    Next I
    
    URLEncodeAnsi = Buf

ErrLbl:
End Function

'### EUC-KR -> 한글(디코딩)
Public Function URLDecodeAnsi(Str As String) As String
On Error GoTo ErrLbl

    Dim AnsiArr() As Byte, I As Long, Buf() As String
    Buf = Split(Str, "%")
    ReDim AnsiArr(UBound(Buf) - 1)
    
    For I = 1 To UBound(Buf)
        AnsiArr(I - 1) = Val("&H" & Buf(I))
    Next I
    
    URLDecodeAnsi = StrConv(AnsiArr, vbUnicode)

ErrLbl:
End Function
Public Function URLEncode(ByRef sStr As String) As String
    Dim I As Long
    Dim sAsc As Integer
    Dim sBuf As String

    For I = 1 To Len(sStr)
        sAsc = Asc(Mid$(sStr, I, 1))

        If Len(Hex(sAsc)) = 4 Then
            URLEncode = URLEncode & "%" & Mid$(Hex(sAsc), 1, 2) & "%" & Mid$(Hex(sAsc), 3, 2)
        Else
            sBuf = Hex(sAsc)
            If Len(sBuf) = 1 Then sBuf = "0" & sBuf
            URLEncode = URLEncode & "%" & sBuf
        End If

        DoEvents
    Next
End Function
'### 한글 -> UTF-8(인코딩)
Public Function URLEncodeUTF8(Str As String) As String
On Error GoTo ErrLbl

    Dim BufSize As Long, MultiArr() As Byte, Buf As String, I As Long
    Dim UniArr() As Byte
    UniArr = Str
    
    BufSize = WideCharToMultiByte(CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, 0&, 0&, 0&, 0&)
    
    If BufSize > 0 Then
        ReDim MultiArr(BufSize - 1&)
        WideCharToMultiByte CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, VarPtr(MultiArr(0)), BufSize, 0&, 0&
    End If
    
    For I = 0 To UBound(MultiArr)
        Buf = Buf & "%" & Hex$(MultiArr(I))
    Next I
    
    URLEncodeUTF8 = Buf

ErrLbl:
End Function

'### UTF-8 -> 한글(디코딩)
Public Function URLDecodeUTF8(Str As String) As String
On Error GoTo ErrLbl

    Dim MultiArr() As Byte, strSplit() As String, I As Long, Converted() As Byte
    strSplit = Split(Str, "%")
    ReDim MultiArr(UBound(strSplit) - 1)
    
    For I = 1 To UBound(strSplit)
        MultiArr(I - 1) = Val("&H" & strSplit(I))
    Next I
    
    Dim BufSize As Long
    BufSize = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(MultiArr(0)), UBound(MultiArr) + 1&, 0&, 0&)
    
    If BufSize > 0 Then
        ReDim Converted(BufSize * 2 + 1)
        MultiByteToWideChar CP_UTF8, 0&, VarPtr(MultiArr(0)), UBound(MultiArr) + 1&, VarPtr(Converted(0)), BufSize
    End If
    
    URLDecodeUTF8 = Converted
   
ErrLbl:
End Function

'### 한글 -> UTF-16LE(인코딩)
Public Function URLEncodeUTF16LE(Str As String) As String
On Error GoTo ErrLbl

    Dim UniArr() As Byte, I As Long, Buf As String
    UniArr = Str
    
    For I = 0 To UBound(UniArr)
        Buf = Buf & "%" & Hex$(UniArr(I))
    Next I
    
    URLEncodeUTF16LE = Buf

ErrLbl:
End Function

'### UTF-16LE -> 한글(디코딩)
Public Function URLDecodeUTF16LE(Str As String) As String
On Error GoTo ErrLbl

    Dim UniArr() As Byte, I As Long, Buf As String, strSplit() As String
    strSplit = Split(Str, "%")
    ReDim UniArr(UBound(strSplit) - 1)

    For I = 1 To UBound(strSplit)
        UniArr(I - 1) = Val("&H" & strSplit(I))
    Next I
    
    URLDecodeUTF16LE = UniArr

ErrLbl:
End Function
