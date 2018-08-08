Attribute VB_Name = "modInven"
Dim W As New WinHttpRequest
Public Sub 팁과노하우(Lv As ListView)
On Error Resume Next
    Dim RE$, i&
    Dim Url$, Subject$, Nick$, sDate$, sHit$, sComent$
    With W
        .Open "GET", "http://www.inven.co.kr/board/powerbbs.php?come_idx=3584"
        .Send
        RE = .ResponseText
        Lv.ListItems.Clear
        For i = 1 To 50
            DoEvents
            Url = "http://www.inven.co.kr/board/powerbbs.php?come_idx=3584" & Split(Split(RE, "http://www.inven.co.kr/board/powerbbs.php?come_idx=3584")(i + 34), """")(0)
            Subject = Replace$(Split(Split(RE, "text-decoration:none;'>")(i + 11), "</A>")(0), "&nbsp;", vbNullString)
            Subject = 태그삭제(Subject)
            sComent = Split(Split(RE, "font-size:11px;color:#25710A;font-weight:bold;letter-spacing:-1px;'>")(i + 11), "</span>")(0)
            sComent = Replace$(sComent, "&nbsp;&nbsp;&nbsp;&nbsp;", "[0]")
            Nick = Split(Split(RE, "layerNickName('")(i + 11), "'")(0)
            sDate = Split(Split(RE, "<TD nowrap class='date'>")(i + 11), "&nbsp;")(0)
            sHit = Split(Split(RE, "<TD nowrap class='hit'>")(i + 11), "</TD>")(0)
            Lv.ListItems.Add , , sComent & Subject
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(1) = Nick
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(2) = sDate
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(3) = sHit
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(4) = Url
        Next i
    End With
End Sub
Public Sub 직업(Lv As ListView, idx As Long)
On Error Resume Next
    Dim RE$, i&
    Dim Url$, Subject$, Nick$, sDate$, sHit$
    With W
        .Open "GET", "http://www.inven.co.kr/board/powerbbs.php?come_idx=" & idx
        .Send
        RE = .ResponseText
        Lv.ListItems.Clear
        For i = 1 To 50
            DoEvents
            Url = "http://www.inven.co.kr/board/powerbbs.php?come_idx=" & idx & Split(Split(RE, "http://www.inven.co.kr/board/powerbbs.php?come_idx=" & idx)(i + 16), """")(0)
            Subject = Replace$(Split(Split(RE, "text-decoration:none;'>")(i + 2), "</A>")(0), "&nbsp;", vbNullString)
            Subject = 태그삭제(Subject)
            sComent = Split(Split(RE, "font-size:11px;color:#25710A;font-weight:bold;letter-spacing:-1px;'>")(i + 2), "</span>")(0)
            sComent = Replace$(sComent, "&nbsp;&nbsp;&nbsp;&nbsp;", "[0]")
            Nick = Split(Split(RE, "layerNickName('")(i + 2), "'")(0)
            sDate = Split(Split(RE, "<TD nowrap class='date'>")(i + 2), "&nbsp;")(0)
            sHit = Split(Split(RE, "<TD nowrap class='hit'>")(i + 2), "</TD>")(0)
            Lv.ListItems.Add , , sComent & Subject
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(1) = Nick
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(2) = sDate
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(3) = sHit
            Lv.ListItems.Item(Lv.ListItems.Count).SubItems(4) = Url
        Next i
    End With
End Sub
Public Sub 레인저(Lv As ListView)

End Sub
Public Sub 자이언트(Lv As ListView)

End Sub
Public Sub 금수랑(Lv As ListView)

End Sub

Function 태그삭제(Str As String) As String
If InStr(Str, "<img") Then
Str = Replace$(Replace$(Str, Split(Split(Str, "<")(1), ">")(0), vbNullString), "<>", vbNullString)
End If
Str = Replace$(Replace$(Str, Split(Split(Str, "[")(1), "]")(0), vbNullString), "[]", vbNullString)
태그삭제 = Replace$(Replace$(Replace$(Str, "<strong>", vbNullString), "</strong>", vbNullString), "&gt;", vbNullString)
End Function
