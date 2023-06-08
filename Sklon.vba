Private Function RussianStringToURLEncode(ByVal txt As String) As String
    Dim i As Long
    Dim L As String
    Dim T As String
    For i = 1 To Len(txt)
        L = Mid(txt, i, 1)
        Select Case AscW(L)
            Case Is > 4095: T = "%" & Hex(AscW(L) \ 64 \ 64 + 224) & "%" & Hex(AscW(L) \ 64) & "%" & Hex(8 * 16 + AscW(L) Mod 64)
            Case Is > 127: T = "%" & Hex(AscW(L) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(L) Mod 64)
            Case 32: T = "%20"
            Case Else: T = L
        End Select
        RussianStringToURLEncode = RussianStringToURLEncode & T
    Next
End Function
Private Function GetHtml(url As String)
    Dim Http As Object
    On Error Resume Next
    Set Http = CreateObject("MSXML2.XMLHTTP.6.0")
    If Err.Number <> 0 Then
        Set Http = CreateObject("MSXML.XMLHTTPRequest.6.0")
    End If
    On Error GoTo 0
    If Http Is Nothing Then
        GetHtml = "Нет подключения"
        Exit Function
    End If
    Http.Open "GET", url, False
    Http.Send
    GetHtml = Http.ResponseText
    Set Http = Nothing
End Function
Private Function СКЛОНЕНИЕ_ФРАЗ(ByVal ФРАЗА As String, Optional ByVal ПАДЕЖ As Integer = 1) As String
    Dim sURL As String
    Dim HTMLText As String
    Dim StartText As Long, EndText As Long
    Dim TegStart As String, TegEnd As String
    Dim PAD As String
    If ПАДЕЖ < 1 Or ПАДЕЖ > 6 Then
        СКЛОНЕНИЕ_ФРАЗ = "Падеж задан неверно"
        Exit Function
    End If
    Select Case ПАДЕЖ
        Case 1: 'именительный
            СКЛОНЕНИЕ_ФРАЗ = ФРАЗА: Exit Function
        Case 2: PAD = "Р"    'родительный
        Case 3: PAD = "Д"    'дательный
        Case 4: PAD = "В"    'винительный
        Case 5: PAD = "Т"    'творительный
        Case 6: PAD = "П"    'предложный
    End Select
    sURL = "https://micro-solution.ru/api/sklon_phrase.php?s=" & RussianStringToURLEncode(ФРАЗА)
    HTMLText = GetHtml(sURL)
    TegStart = "<" & PAD & ">"
    TegEnd = "</" & PAD & ">"
    If InStr(1, HTMLText, "error") > 0 Then
        TegStart = "<message>"
        TegEnd = "</message>"
    End If
    StartText = InStr(1, HTMLText, TegStart)
    EndText = InStr(StartText, HTMLText, TegEnd)
    СКЛОНЕНИЕ_ФРАЗ = Mid(HTMLText, StartText + Len(TegStart), EndText - StartText - Len(TegStart))
End Function
