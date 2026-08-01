Attribute VB_Name = "modGovNotice"
Option Explicit

Private Const MAX_POSTS As Long = 5
Private Const HTTP_TIMEOUT_MS As Long = 20000
Private Const BODY_COL_WIDTH As Long = 40
Private Const BODY_TEXT_WIDTH As Long = 50   ' column width 40 + 10 chars
Private Const CARD_HEIGHT As Long = 7   ' header + 5 items + spacer
Private Const CARD_COLUMNS As Long = 4
Public gSilent As Boolean

Public Sub FetchNoticesSilent()
    gSilent = True
    On Error GoTo Clean
    FetchNotices
Clean:
    gSilent = False
End Sub

Public Sub SetupResultButtons()
    Dim wsRes As Worksheet
    Set wsRes = SheetByName(ChrW(&HACB0&) & ChrW(&HACFC&))
    If wsRes Is Nothing Then Set wsRes = SheetByName("Result")
    If wsRes Is Nothing Then Exit Sub
    EnsureFetchButton wsRes
End Sub

Public Sub FetchNotices()
    Dim wsCfg As Worksheet, wsRes As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, r As Long
    Dim deptName As String, url As String, useYn As String
    Dim includeHint As String, excludeHint As String
    Dim html As String, errMsg As String
    Dim posts As Collection
    Dim okCnt As Long, failCnt As Long, total As Long, done As Long
    Dim collectTime As String
    Dim cardIndex As Long
    Dim startCol As Long, startRow As Long

    On Error GoTo Fatal

    Set wsCfg = SheetByName(ChrW(&HC124&) & ChrW(&HC815&))
    Set wsRes = SheetByName(ChrW(&HACB0&) & ChrW(&HACFC&))
    Set wsLog = SheetByName(ChrW(&HB85C&) & ChrW(&HADF8&))
    If wsCfg Is Nothing Then Set wsCfg = SheetByName("Config")
    If wsRes Is Nothing Then Set wsRes = SheetByName("Result")
    If wsLog Is Nothing Then Set wsLog = SheetByName("Log")
    If wsCfg Is Nothing Or wsRes Is Nothing Then
        If Not gSilent Then MsgBox "Required sheets not found.", vbExclamation
        Exit Sub
    End If

    ClearResultBoard wsRes
    If Not wsLog Is Nothing Then ClearSheetData wsLog, 2

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
    total = 0
    For r = 2 To lastRow
        If Len(Trim$(CStr(wsCfg.Cells(r, 1).Value))) > 0 Then total = total + 1
    Next r

    If total = 0 Then
        If Not gSilent Then MsgBox "No ministries in settings sheet.", vbExclamation
        Exit Sub
    End If

    collectTime = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    okCnt = 0
    failCnt = 0
    done = 0
    cardIndex = 0

    Application.ScreenUpdating = False
    SetupResultBoardHeader wsRes, collectTime
    EnsureFetchButton wsRes

    For r = 2 To lastRow
        deptName = Trim$(CStr(wsCfg.Cells(r, 1).Value))
        url = Trim$(CStr(wsCfg.Cells(r, 2).Value))
        useYn = UCase$(Trim$(CStr(wsCfg.Cells(r, 3).Value)))
        includeHint = Trim$(CStr(wsCfg.Cells(r, 5).Value))
        excludeHint = Trim$(CStr(wsCfg.Cells(r, 6).Value))

        If Len(deptName) = 0 Then GoTo NextMinistry

        done = done + 1
        Application.StatusBar = "Collecting " & CStr(done) & "/" & CStr(total) & " - " & deptName
        DoEvents

        startCol = CardStartCol(cardIndex)
        startRow = 3 + (cardIndex \ CARD_COLUMNS) * CARD_HEIGHT

        errMsg = ""
        html = ""
        Set posts = Nothing

        If useYn <> "Y" Then
            errMsg = "UseYn=N (change to Y and verify URL)"
        ElseIf Len(url) = 0 Or InStr(1, url, "http", vbTextCompare) <> 1 Then
            errMsg = "Invalid list URL"
        Else
            html = GetHtml(url, errMsg)
        End If

        If Len(errMsg) = 0 Then
            Set posts = ParseAnnouncementLinks(html, url, includeHint, excludeHint, MAX_POSTS)
            If (posts Is Nothing Or posts.Count = 0) And Len(includeHint) > 0 Then
                Set posts = ParseAnnouncementLinks(html, url, "", excludeHint, MAX_POSTS)
            End If
            If posts Is Nothing Or posts.Count = 0 Then
                errMsg = "No post links found (check HTML / include hint)"
            End If
        End If

        If Len(errMsg) > 0 Then
            failCnt = failCnt + 1
            WriteDeptFailCard wsRes, startRow, startCol, deptName
            WriteLog wsLog, deptName, "FAIL", 0, errMsg, collectTime
        Else
            okCnt = okCnt + 1
            WriteDeptSuccessCard wsRes, startRow, startCol, deptName, posts
            WriteLog wsLog, deptName, "OK", posts.Count, "", collectTime
        End If

        cardIndex = cardIndex + 1

NextMinistry:
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Not gSilent Then MsgBox "Done: OK " & CStr(okCnt) & " / FAIL " & CStr(failCnt) & " / TOTAL " & CStr(cardIndex), vbInformation
    Exit Sub

Fatal:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Not gSilent Then MsgBox "Fatal: " & Err.Description, vbCritical
End Sub
Private Function GetHtml(ByVal url As String, ByRef errMsg As String) As String
    Dim html As String
    Dim err2 As String
    Dim attempt As Long
    errMsg = ""
    html = ""
    For attempt = 1 To 6
        errMsg = ""
        html = GetHtmlWinHttp(url, errMsg)
        If Len(html) > 0 Then Exit For
        err2 = ""
        html = GetHtmlXmlHttp(url, err2)
        If Len(html) > 0 Then
            errMsg = ""
            Exit For
        End If
        If Len(err2) > 0 Then errMsg = errMsg & " / " & err2
        If attempt < 6 Then
            DoEvents
            On Error Resume Next
            Application.Wait Now + TimeSerial(0, 0, attempt)
            On Error GoTo 0
        End If
    Next attempt
    GetHtml = html
End Function

Private Function GetHtmlWinHttp(ByVal url As String, ByRef errMsg As String) As String
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    On Error Resume Next
    ' WinHttpRequestOption_SecureProtocols = 9 ; enable TLS1.0/1.1/1.2
    http.Option(9) = 128 + 512 + 2048
    On Error GoTo Fail
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    http.setRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en;q=0.8"
    http.Send
    If http.status < 200 Or http.status >= 400 Then
        errMsg = "HTTP " & CStr(http.status)
        GetHtmlWinHttp = ""
        Exit Function
    End If
    GetHtmlWinHttp = CStr(http.responseText)
    If Len(Trim$(GetHtmlWinHttp)) = 0 Then errMsg = "Empty HTML"
    Exit Function
Fail:
    errMsg = "WinHTTP: " & Err.Description
    GetHtmlWinHttp = ""
End Function

Private Function GetHtmlXmlHttp(ByVal url As String, ByRef errMsg As String) As String
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    http.Send
    If http.status < 200 Or http.status >= 400 Then
        errMsg = "XMLHTTP " & CStr(http.status)
        GetHtmlXmlHttp = ""
        Exit Function
    End If
    GetHtmlXmlHttp = CStr(http.responseText)
    If Len(Trim$(GetHtmlXmlHttp)) = 0 Then errMsg = "Empty HTML"
    Exit Function
Fail:
    errMsg = "XMLHTTP: " & Err.Description
    GetHtmlXmlHttp = ""
End Function

Private Function ParseAnnouncementLinks(ByVal html As String, ByVal baseUrl As String, _
                                       ByVal includeHint As String, ByVal excludeHint As String, _
                                       ByVal maxCount As Long) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim pos As Long, hrefPos As Long, endTag As Long
    Dim href As String, title As String, absUrl As String, dateText As String
    Dim lowerHtml As String
    Dim aOpen As Long
    Dim jsAbs As String
    Dim onclickValue As String
    Dim titleAttr As String
    Dim loopCount As Long
    Dim accepted As Boolean

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    lowerHtml = LCase$(html)
    pos = 1
    loopCount = 0

    Set result = TrySpecializedParse(html, baseUrl, maxCount)
    If Not result Is Nothing Then
        If result.Count > 0 Then
            Set ParseAnnouncementLinks = result
            Exit Function
        End If
    End If

    On Error GoTo ParseDone

    Do While result.Count < maxCount
        loopCount = loopCount + 1
        If loopCount > 8000 Then Exit Do
        aOpen = InStr(pos, lowerHtml, "<a ")
        If aOpen = 0 Then aOpen = InStr(pos, lowerHtml, "<a>")
        If aOpen = 0 Then Exit Do

        hrefPos = InStr(aOpen, lowerHtml, "href=")
        endTag = InStr(aOpen, lowerHtml, ">")
        If hrefPos = 0 Or endTag = 0 Or hrefPos > endTag Then
            pos = aOpen + 2
            GoTo ContParse
        End If

        href = ExtractAttrValue(html, hrefPos + 5)
        onclickValue = GetTagAttrValue(html, aOpen, endTag, "onclick")
        titleAttr = GetTagAttrValue(html, aOpen, endTag, "title")
        accepted = False
        jsAbs = ""

        If IsLikelyPostLink(href, includeHint, excludeHint) Then
            accepted = True
        ElseIf InStr(1, href, "fn_egov_select", vbTextCompare) > 0 Or _
               InStr(1, href, "fn_detail", vbTextCompare) > 0 Or _
               InStr(1, href, "goView", vbTextCompare) > 0 Then
            jsAbs = ConvertJsToUrl(href, onclickValue, baseUrl, html)
            If Len(jsAbs) > 0 Then
                href = jsAbs
                accepted = True
            End If
        ElseIf Len(onclickValue) > 8 Then
            If InStr(1, onclickValue, "fn_", vbTextCompare) > 0 Or _
               InStr(1, onclickValue, "goView", vbTextCompare) > 0 Or _
               InStr(1, onclickValue, "fnView", vbTextCompare) > 0 Then
                jsAbs = ConvertJsToUrl(href, onclickValue, baseUrl, html)
                If Len(jsAbs) > 0 Then
                    href = jsAbs
                    accepted = True
                End If
            End If
        End If

        If Not accepted Then
            pos = endTag + 1
            GoTo ContParse
        End If

        title = ExtractLinkText(html, endTag)
        title = CleanText(title)
        If Len(title) < 2 Then
            title = CleanText(titleAttr)
            ' strip "[post link]" style prefixes
            If Left$(title, 1) = "[" Then
                Dim br As Long
                br = InStr(title, "]")
                If br > 0 And br < 30 Then title = Trim$(mId$(title, br + 1))

            End If
        End If
        If Len(title) < 2 Or Len(title) > 300 Then
            pos = endTag + 1
            GoTo ContParse
        End If
        If IsJunkTitle(title) Then
            pos = endTag + 1
            GoTo ContParse
        End If

        absUrl = CleanUrl(ToAbsoluteUrl(baseUrl, href))
        If Len(absUrl) = 0 Or InStr(1, absUrl, "http", vbTextCompare) <> 1 Then
            pos = endTag + 1
            GoTo ContParse
        End If
        If seen.Exists(LCase$(absUrl)) Then
            pos = endTag + 1
            GoTo ContParse
        End If

        dateText = GuessDateNearLink(html, aOpen)
        seen.Add LCase$(absUrl), True
        result.Add Array(title, absUrl, dateText)
        pos = endTag + 1
ContParse:
    Loop

ParseDone:
    On Error GoTo 0
    Set ParseAnnouncementLinks = result
End Function

Private Function TrySpecializedParse(ByVal html As String, ByVal baseUrl As String, ByVal maxCount As Long) As Collection
    Dim result As Collection
    Dim lowerUrl As String

    lowerUrl = LCase$(baseUrl)
    Set result = New Collection

    If InStr(lowerUrl, "msit.go.kr") > 0 And InStr(lowerUrl, "/bbs/list.do") > 0 Then
        ExtractMsitLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mofa.go.kr") > 0 And (InStr(lowerUrl, "/www/brd/m_4075/list.do") > 0 Or InStr(lowerUrl, "/www/brd/m_4080/list.do") > 0) Then
        ExtractByRegex html, baseUrl, "href=""(\./view\.do\?seq=\d+[^""]*)""[^>]*>\s*(?:<span[^>]*>[^<]*</span>\s*)?([^<]{2,})", result, maxCount
    ElseIf InStr(lowerUrl, "mois.go.kr") > 0 And InStr(lowerUrl, "bbsmstr_000000000006") > 0 Then
        ExtractMoisLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "moe.go.kr") > 0 And InStr(lowerUrl, "boardcnts/listrenew.do") > 0 Then
        ExtractMoeLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mpm.go.kr") > 0 And (InStr(lowerUrl, "newsnoitice") > 0 Or InStr(lowerUrl, "noticelist") > 0) Then
        ExtractMpmLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "moj.go.kr") > 0 And InStr(lowerUrl, "/moj/223/subview.do") > 0 Then
        ExtractByRegexTitleHref html, baseUrl, "title=""([^""]+)""[\s\S]{0,120}?<a href=""(/bbs/moj/184/\d+/artclView\.do)""", result, maxCount
    ElseIf InStr(lowerUrl, "mnd.go.kr") > 0 And InStr(lowerUrl, "/mnd/154/subview.do") > 0 Then
        ExtractMndLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mafra.go.kr") > 0 And InStr(lowerUrl, "/home/5108/subview.do") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(/bbs/home/791/\d+/artclView\.do)""[\s\S]{0,140}?\r?\n\s*([^<\r\n]{2,})\s*\r?\n\s*<span class=""new"">", result, maxCount
    ElseIf InStr(lowerUrl, "motir.go.kr") > 0 And InStr(lowerUrl, "/kor/article/") > 0 Then
        ExtractMotirLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mss.go.kr") > 0 And InStr(lowerUrl, "cbidx=81") > 0 Then
        ExtractMssLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mcst.go.kr") > 0 And InStr(lowerUrl, "noticelist.jsp") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(noticeView\.jsp\?pSeq=[^""]+)""[^>]*title=""([^""]+)""", result, maxCount
    ElseIf InStr(lowerUrl, "unikorea.go.kr") > 0 And InStr(lowerUrl, "bbs_0000000000000001") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(/web/unikorea/bbs/bbs_0000000000000001/\d+[^""]*)""[^>]*>\s*([^<]+?)\s*</a>", result, maxCount
    ElseIf InStr(lowerUrl, "molit.go.kr") > 0 And InStr(lowerUrl, "/usr/bord0201/m_69/brd.jsp") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(\./DTL\.jsp[^""]+)""[^>]*>([\s\S]*?)</a>", result, maxCount
    ElseIf InStr(lowerUrl, "mpva.go.kr") > 0 And InStr(lowerUrl, "selectbbsnttlist.do") > 0 Then
        ExtractMpvaLinks html, baseUrl, result, maxCount
    ElseIf InStr(lowerUrl, "mfds.go.kr") > 0 And InStr(lowerUrl, "/brd/m_689/") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(\./view\.do\?seq=\d+[^""]*)""[^>]*class=""title""[^>]*>\s*([^<]{2,})", result, maxCount
    ElseIf InStr(lowerUrl, "mohw.go.kr") > 0 And InStr(lowerUrl, "bid=0003") > 0 Then
        ExtractByRegex html, baseUrl, "href=""(/board\.es\?[^""]*act=view[^""]*list_no=\d+[^""]*)""[^>]*class=""txt_title""[^>]*>[\s\S]*?(?:</span>\s*)?([^<]{2,})</a>", result, maxCount
    End If

    Set TrySpecializedParse = result
End Function

Private Sub ExtractMoisLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim nttId As String, title As String, href As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' 1) onclick: fn_egov_inqire_notice('nttId','BBSMSTR_...')
    re.pattern = "inqire_notice\('(\d+)',\s*'BBSMSTR_000000000006'\)[^>]*>\s*([^<]{2,})"
    Set matches = re.Execute(html)
    For Each m In matches
        nttId = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        If Len(title) < 2 Or IsJunkTitle(title) Then GoTo NextMois1
        href = "https://www.mois.go.kr/frt/bbs/type013/commonSelectBoardArticle.do?bbsId=BBSMSTR_000000000006&nttId=" & nttId
        If Not seen.Exists(LCase$(href)) Then
            seen.Add LCase$(href), True
            result.Add Array(title, href, "")
            If result.Count >= maxCount Then Exit Sub
        End If
NextMois1:
    Next m

    If result.Count >= maxCount Then Exit Sub

    ' 2) href fallback (jsessionid stripped later via clean URL rebuild)
    re.pattern = "commonSelectBoardArticle\.do(?:;jsessionid=[^?\""']+)?\?bbsId=BBSMSTR_000000000006(?:&amp;|&)nttId=(\d+)[^\""']*[\""'][^>]*>\s*([^<]{2,})"
    Set matches = re.Execute(html)
    For Each m In matches
        nttId = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        If Len(title) < 2 Or IsJunkTitle(title) Then GoTo NextMois2
        href = "https://www.mois.go.kr/frt/bbs/type013/commonSelectBoardArticle.do?bbsId=BBSMSTR_000000000006&nttId=" & nttId
        If Not seen.Exists(LCase$(href)) Then
            seen.Add LCase$(href), True
            result.Add Array(title, href, "")
            If result.Count >= maxCount Then Exit Sub
        End If
NextMois2:
    Next m
End Sub

Private Sub ExtractMoeLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim boardSeq As String, title As String, href As String
    Dim menuVal As String, seen As Object
    menuVal = ExtractQueryValue(baseUrl, "m")
    If Len(menuVal) = 0 Then menuVal = "020501"
    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "goView\('333',\s*'(\d+)'[^""]*""[^>]*title=""([^""]+)"""
    re.IgnoreCase = True
    re.Global = True
    Set matches = re.Execute(html)
    For Each m In matches
        boardSeq = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        If Len(title) < 2 Or IsJunkTitle(title) Then GoTo NextMoe
        href = "https://www.moe.go.kr/boardCnts/viewRenew.do?boardID=333&boardSeq=" & boardSeq & _
               "&lev=0&searchType=null&statusYN=W&page=1&s=moe&m=" & menuVal & "&opType=N"
        If Not seen.Exists(LCase$(href)) Then
            seen.Add LCase$(href), True
            result.Add Array(title, href, "")
            If result.Count >= maxCount Then Exit For
        End If
NextMoe:
    Next m
End Sub

Private Sub ExtractMpmLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object

    Dim href As String, title As String, absUrl As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "href=""(\?boardId=[^""]*mode=view[^""]*cntId=\d+[^""]*)""\s*>\s*([^<]{2,})\s*<"
    re.IgnoreCase = True
    re.Global = True
    Set matches = re.Execute(html)
    For Each m In matches
        href = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        absUrl = CleanUrl(ToAbsoluteUrl(baseUrl, href))
        If Len(title) >= 2 And Len(absUrl) > 0 And Not IsJunkTitle(title) Then
            If Not seen.Exists(LCase$(absUrl)) Then
                seen.Add LCase$(absUrl), True
                result.Add Array(title, absUrl, "")
                If result.Count >= maxCount Then Exit For
            End If
        End If
    Next m
End Sub

Private Sub ExtractMpvaLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim nttNo As String, title As String, href As String
    Dim keyVal As String, bbsNo As String
    keyVal = ExtractQueryValue(baseUrl, "key")
    bbsNo = ExtractQueryValue(baseUrl, "bbsNo")
    If Len(keyVal) = 0 Then keyVal = "76"
    If Len(bbsNo) = 0 Then bbsNo = "15"

    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "href=""[^""]*selectBbsNttView\.do[^""]*nttNo=(\d+)[^""]*""[^>]*>\s*([^<]{2,})</a>"
    re.IgnoreCase = True
    re.Global = True
    Set matches = re.Execute(html)
    For Each m In matches
        nttNo = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        If Len(title) < 2 Or IsJunkTitle(title) Then GoTo NextMpva
        href = "https://www.mpva.go.kr/mpva/selectBbsNttView.do?key=" & keyVal & "&bbsNo=" & bbsNo & "&nttNo=" & nttNo
        result.Add Array(title, href, "")
        If result.Count >= maxCount Then Exit For
NextMpva:
    Next m
End Sub

Private Sub ExtractMsitLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim reKey As Object, reTitle As Object, keyMatches As Object, titleMatches As Object
    Dim i As Long, href As String, title As String
    Set reKey = CreateObject("VBScript.RegExp")
    reKey.pattern = "onclick=""fn_detail\((\d+)\);""[^>]*title=""[^""]+"""
    reKey.IgnoreCase = True
    reKey.Global = True
    Set keyMatches = reKey.Execute(html)

    Set reTitle = CreateObject("VBScript.RegExp")
    reTitle.pattern = "sHtml\+= unescape\('([^']+)'\);"
    reTitle.IgnoreCase = True
    reTitle.Global = True
    Set titleMatches = reTitle.Execute(html)

    For i = 0 To keyMatches.Count - 1
        If i > titleMatches.Count - 1 Then Exit For
        href = "fn_detail(" & CStr(keyMatches.Item(i).SubMatches.Item(0)) & ")"
        title = CleanText(CStr(titleMatches.Item(i).SubMatches.Item(0)))
        result.Add Array(title, ConvertFnDetailToUrl(href, baseUrl, html), "")
        If result.Count >= maxCount Then Exit For
    Next i
End Sub

Private Sub ExtractByRegexTitleHref(ByVal html As String, ByVal baseUrl As String, ByVal pattern As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim href As String, title As String, absUrl As String
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True

    Set matches = re.Execute(html)
    For Each m In matches
        title = CleanText(CStr(m.SubMatches.Item(0)))
        href = CStr(m.SubMatches.Item(1))
        absUrl = CleanUrl(ToAbsoluteUrl(baseUrl, href))
        If Len(title) >= 2 And Len(absUrl) > 0 Then
            If Not seen.Exists(LCase$(absUrl)) Then
                seen.Add LCase$(absUrl), True
                result.Add Array(title, absUrl, "")
                If result.Count >= maxCount Then Exit For
            End If
        End If
    Next m
End Sub

Private Sub ExtractMotirLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim articleId As String, title As String, href As String
    Dim boardID As String, p As Long, q As Long
    boardID = "ATCL6e90bb9de"
    p = InStr(1, LCase$(baseUrl), "/kor/article/", vbTextCompare)
    If p > 0 Then
        q = p + Len("/kor/article/")
        boardID = mId$(baseUrl, q)
        If InStr(boardID, "?") > 0 Then boardID = Left$(boardID, InStr(boardID, "?") - 1)
        If InStr(boardID, "/") > 0 Then boardID = Left$(boardID, InStr(boardID, "/") - 1)
        boardID = Trim$(boardID)
        If Len(boardID) = 0 Then boardID = "ATCL6e90bb9de"
    End If
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "article\.view\('(\d+)'\);""\s*>\s*<i>([^<]+)</i>"
    re.IgnoreCase = True
    re.Global = True
    Set matches = re.Execute(html)
    For Each m In matches
        articleId = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        href = "https://www.motir.go.kr/kor/article/" & boardID & "/" & articleId & "/view?mno=&pageIndex=1&rowPageC=0&displayAuthor=&searchCategory=0&schClear=on&startDtD=&endDtD=&searchCondition=1&searchKeyword="
        result.Add Array(title, href, "")
        If result.Count >= maxCount Then Exit For
    Next m
End Sub

Private Sub ExtractMndLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim reHref As Object, reTitle As Object, hrefMatches As Object, titleMatches As Object
    Dim i As Long, href As String, title As String
    Set reHref = CreateObject("VBScript.RegExp")
    reHref.pattern = "href=""(/bbs/mnd/11066/[^""]+/artclView\.do)"""
    reHref.IgnoreCase = True
    reHref.Global = True
    Set hrefMatches = reHref.Execute(html)

    Set reTitle = CreateObject("VBScript.RegExp")
    reTitle.pattern = "<strong><span>([^<]+)</span></strong>"
    reTitle.IgnoreCase = True
    reTitle.Global = True
    Set titleMatches = reTitle.Execute(html)

    For i = 0 To hrefMatches.Count - 1
        If i > titleMatches.Count - 1 Then Exit For
        href = CStr(hrefMatches.Item(i).SubMatches.Item(0))
        title = CleanText(CStr(titleMatches.Item(i).SubMatches.Item(0)))
        result.Add Array(title, CleanUrl(ToAbsoluteUrl(baseUrl, href)), "")
        If result.Count >= maxCount Then Exit For
    Next i
End Sub
Private Sub ExtractMssLinks(ByVal html As String, ByVal baseUrl As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim bcIdx As String, parentSeq As String, title As String, href As String
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "doBbsFView\('81','(\d+)','[^']*','(\d+)'\);""\s+title=""([^""]+)"""
    re.IgnoreCase = True
    re.Global = True
    Set matches = re.Execute(html)
    For Each m In matches
        bcIdx = CStr(m.SubMatches.Item(0))
        parentSeq = CStr(m.SubMatches.Item(1))
        title = CleanText(CStr(m.SubMatches.Item(2)))
        href = "https://www.mss.go.kr/site/smba/ex/bbs/View.do?cbIdx=81&bcIdx=" & bcIdx & "&parentSeq=" & parentSeq
        result.Add Array(title, href, "")
        If result.Count >= maxCount Then Exit For
    Next m
End Sub

Private Sub ExtractByRegex(ByVal html As String, ByVal baseUrl As String, ByVal pattern As String, ByRef result As Collection, ByVal maxCount As Long)
    Dim re As Object, matches As Object, m As Object
    Dim href As String, title As String, absUrl As String
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = True
    re.MultiLine = True

    Set matches = re.Execute(html)
    For Each m In matches
        href = CStr(m.SubMatches.Item(0))
        title = CleanText(CStr(m.SubMatches.Item(1)))
        absUrl = CleanUrl(ToAbsoluteUrl(baseUrl, href))
        If Len(title) >= 2 And Len(absUrl) > 0 Then
            If Not seen.Exists(LCase$(absUrl)) Then
                seen.Add LCase$(absUrl), True
                result.Add Array(title, absUrl, "")
                If result.Count >= maxCount Then Exit For
            End If
        End If
    Next m
End Sub

Private Function ConvertJsToUrl(ByVal href As String, ByVal onclk As String, ByVal baseUrl As String, ByVal html As String) As String
    Dim result As String
    ConvertJsToUrl = ""
    On Error GoTo JsUrlFail
    
    result = ConvertFnEgovSelectToUrl(href, baseUrl)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertFnDetailToUrl(href, baseUrl, html)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertFnEgovSelectToUrl(onclk, baseUrl)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertFnDetailToUrl(onclk, baseUrl, html)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertGoViewToUrl(onclk, baseUrl)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertMogefToUrl(onclk, baseUrl)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone
    result = ConvertFnSelectDocToUrl(onclk, baseUrl)
    If Len(result) > 0 Then ConvertJsToUrl = result: GoTo JsUrlDone

JsUrlDone:
    On Error GoTo 0
    Exit Function
JsUrlFail:
    ConvertJsToUrl = ""
    Resume JsUrlDone
End Function

Private Function ConvertGoViewToUrl(ByVal onclk As String, ByVal baseUrl As String) As String
    Dim h As String, p As Long, args As String
    Dim parts() As String, boardID As String, boardSeq As String
    Dim statusYN As String, currPage As String, menuVal As String
    Dim schemeHost As String, i As Long, slash As Long, host As String
    
    ConvertGoViewToUrl = ""
    h = Trim$(onclk)
    p = InStr(1, h, "goView(", vbTextCompare)
    If p = 0 Then Exit Function
    
    p = InStr(p, h, "(")
    Dim p2 As Long
    p2 = InStr(p + 1, h, ")")
    If p = 0 Or p2 = 0 Then Exit Function
    args = mId$(h, p + 1, p2 - p - 1)
    args = Replace(Replace(Replace(args, "'", ""), """", ""), " ", "")
    parts = Split(args, ",")
    If UBound(parts) < 1 Then Exit Function
    boardID = parts(0)
    boardSeq = parts(1)
    statusYN = "W"
    currPage = "1"
    If UBound(parts) >= 4 Then statusYN = parts(4)
    If UBound(parts) >= 5 Then currPage = parts(5)
    
    menuVal = ExtractQueryValue(baseUrl, "m")
    If Len(menuVal) = 0 Then menuVal = "020501"
    
    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then schemeHost = Left$(baseUrl, i + 2) & Left$(host, slash - 1) Else schemeHost = Left$(baseUrl, i + 2) & host
    
    ConvertGoViewToUrl = schemeHost & "/boardCnts/viewRenew.do?boardID=" & boardID & _
                         "&boardSeq=" & boardSeq & "&lev=0&searchType=null&statusYN=" & statusYN & _
                         "&page=" & currPage & "&s=moe&m=" & menuVal & "&opType=N"
End Function

Private Function ConvertMogefToUrl(ByVal onclk As String, ByVal baseUrl As String) As String
    Dim h As String, p As Long, p2 As Long, bbtSn As String
    Dim schemeHost As String, i As Long, slash As Long, host As String
    Dim midVal As String
    
    ConvertMogefToUrl = ""
    h = Trim$(onclk)
    If InStr(1, h, "fn_selectView(", vbTextCompare) = 0 Then Exit Function
    
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "(\d{5,})"
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(h)
    If matches.Count = 0 Then Exit Function
    bbtSn = CStr(matches.Item(0).SubMatches.Item(0))
    
    midVal = ExtractQueryValue(baseUrl, "mid")
    If Len(midVal) = 0 Then midVal = "news400"
    
    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then schemeHost = Left$(baseUrl, i + 2) & Left$(host, slash - 1) Else schemeHost = Left$(baseUrl, i + 2) & host
    
    ConvertMogefToUrl = schemeHost & "/nw/ntc/nw_ntc_s001d.do?mid=" & midVal & "&bbtSn=" & bbtSn
End Function

Private Function ConvertFnSelectDocToUrl(ByVal onclk As String, ByVal baseUrl As String) As String
    Dim h As String, p As Long, p2 As Long, docSeq As String
    Dim schemeHost As String, i As Long, slash As Long, host As String
    Dim menuSeq As String, bbsSeq As String
    
    ConvertFnSelectDocToUrl = ""
    h = Trim$(onclk)

    p = InStr(1, h, "fn_selectDoc(", vbTextCompare)
    If p = 0 Then Exit Function
    
    p = InStr(p, h, "(")
    p2 = InStr(p + 1, h, ")")
    If p = 0 Or p2 = 0 Then Exit Function
    docSeq = mId$(h, p + 1, p2 - p - 1)
    docSeq = Replace(Replace(Replace(Trim$(docSeq), "'", ""), """", ""), " ", "")
    If Len(docSeq) = 0 Then Exit Function
    
    menuSeq = ExtractQueryValue(baseUrl, "menuSeq")
    bbsSeq = ExtractQueryValue(baseUrl, "bbsSeq")
    If Len(menuSeq) = 0 Then menuSeq = "375"
    If Len(bbsSeq) = 0 Then bbsSeq = "9"
    
    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then schemeHost = Left$(baseUrl, i + 2) & Left$(host, slash - 1) Else schemeHost = Left$(baseUrl, i + 2) & host
    
    ConvertFnSelectDocToUrl = schemeHost & "/doc/ko/selectDoc.do?docSeq=" & docSeq & "&menuSeq=" & menuSeq & "&bbsSeq=" & bbsSeq
End Function

Private Function ConvertFnEgovSelectToUrl(ByVal href As String, ByVal baseUrl As String) As String
    Dim h As String, p As Long, q As Long, r As Long
    Dim nttId As String, bbsId As String, menuNo As String
    Dim schemeHost As String, i As Long, slash As Long, host As String

    ConvertFnEgovSelectToUrl = ""
    h = Trim$(href)
    p = InStr(1, h, "fn_egov_select(", vbTextCompare)
    If p = 0 Then Exit Function

    p = InStr(p, h, "(")
    If p = 0 Then Exit Function
    q = InStr(p + 1, h, ",")
    r = InStr(q + 1, h, ")")
    If q = 0 Or r = 0 Then Exit Function

    nttId = mId$(h, p + 1, q - p - 1)
    bbsId = mId$(h, q + 1, r - q - 1)
    nttId = Replace(Replace(Replace(Trim$(nttId), "'", ""), """", ""), " ", "")
    bbsId = Replace(Replace(Replace(Trim$(bbsId), "'", ""), """", ""), " ", "")
    If Len(nttId) = 0 Or Len(bbsId) = 0 Then Exit Function

    menuNo = ExtractQueryValue(baseUrl, "menuNo")
    If Len(menuNo) = 0 Then menuNo = "4050100"

    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then
        schemeHost = Left$(baseUrl, i + 2) & Left$(host, slash - 1)
    Else
        schemeHost = Left$(baseUrl, i + 2) & host
    End If

    ConvertFnEgovSelectToUrl = schemeHost & "/nw/nes/detailNesDtaView.do?searchBbsId1=" & bbsId & _
                               "&searchNttId1=" & nttId & "&menuNo=" & menuNo
End Function

Private Function ConvertFnDetailToUrl(ByVal href As String, ByVal baseUrl As String, ByVal html As String) As String
    Dim h As String, p1 As Long, p2 As Long
    Dim nttSeqNo As String, bbsSeqNo As String
    Dim sCode As String, menuId As String, menuPid As String
    Dim schemeHost As String, i As Long, slash As Long, host As String
    Dim basePath As String

    ConvertFnDetailToUrl = ""
    h = Trim$(href)
    If InStr(1, h, "fn_detail(", vbTextCompare) = 0 Then Exit Function

    p1 = InStr(1, h, "(")
    p2 = InStr(p1 + 1, h, ")")
    If p1 = 0 Or p2 = 0 Then Exit Function
    nttSeqNo = mId$(h, p1 + 1, p2 - p1 - 1)
    nttSeqNo = Replace(Replace(Replace(Trim$(nttSeqNo), "'", ""), """", ""), " ", "")
    If Len(nttSeqNo) = 0 Then Exit Function

    bbsSeqNo = ExtractHiddenValue(html, "bbsSeqNo")
    sCode = ExtractHiddenValue(html, "sCode")
    menuId = ExtractHiddenValue(html, "mId")
    menuPid = ExtractHiddenValue(html, "mPid")
    If Len(bbsSeqNo) = 0 Then bbsSeqNo = ExtractQueryValue(baseUrl, "bbsSeqNo")
    If Len(sCode) = 0 Then sCode = ExtractQueryValue(baseUrl, "sCode")
    If Len(menuId) = 0 Then menuId = ExtractQueryValue(baseUrl, "mId")
    If Len(menuPid) = 0 Then menuPid = ExtractQueryValue(baseUrl, "mPid")
    If Len(sCode) = 0 Then sCode = "user"

    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then
        schemeHost = Left$(baseUrl, i + 2) & Left$(host, slash - 1)
    Else
        schemeHost = Left$(baseUrl, i + 2) & host
    End If

    basePath = "/bbs/view.do?sCode=" & sCode
    If Len(menuId) > 0 Then basePath = basePath & "&mId=" & menuId
    If Len(menuPid) > 0 Then basePath = basePath & "&mPid=" & menuPid
    basePath = basePath & "&pageIndex=&bbsSeqNo=" & bbsSeqNo & "&nttSeqNo=" & nttSeqNo & "&searchOpt=ALL&searchTxt="
    ConvertFnDetailToUrl = schemeHost & basePath
End Function

Private Function ExtractHiddenValue(ByVal html As String, ByVal fieldName As String) As String
    Dim pattern As String
    Dim re As Object, matches As Object
    Dim m As Object
    ExtractHiddenValue = ""
    pattern = "<input[^>]*name\s*=\s*[""']" & fieldName & "[""'][^>]*value\s*=\s*[""']([^""']*)[""']"
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    Set matches = re.Execute(html)
    If matches.Count > 0 Then
        Set m = matches.Item(0)
        ExtractHiddenValue = CStr(m.SubMatches.Item(0))
    End If
End Function

Private Function ExtractQueryValue(ByVal url As String, ByVal key As String) As String
    Dim q As Long, p As Long, amp As Long, chunk As String
    ExtractQueryValue = ""
    q = InStr(1, url, "?")
    If q = 0 Then Exit Function
    chunk = mId$(url, q + 1)
    p = InStr(1, chunk, key & "=", vbTextCompare)
    If p = 0 Then Exit Function
    chunk = mId$(chunk, p + Len(key) + 1)
    amp = InStr(chunk, "&")
    If amp > 0 Then chunk = Left$(chunk, amp - 1)
    ExtractQueryValue = chunk
End Function

Private Function CleanUrl(ByVal u As String) As String
    Dim s As String, p As Long, q As Long
    s = Trim$(u)
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&quot;", """")
    s = Replace(s, "&#034;", """")
    s = Replace(s, "/./", "/")
    ' remove ;jsessionid=XXX before ? or end
    p = InStr(1, s, ";jsessionid=", vbTextCompare)
    If p > 0 Then
        q = InStr(p, s, "?")

        If q > 0 Then
            s = Left$(s, p - 1) & mId$(s, q)
        Else
            s = Left$(s, p - 1)
        End If
    End If
    CleanUrl = s
End Function

' ".jsp" must not be treated as a ".js" asset, so extensions are matched at a path boundary.
Private Function IsStaticAsset(ByVal lowerHref As String) As Boolean
    Dim exts As Variant, i As Long
    exts = Array(".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".ico", ".woff", ".ttf")
    For i = LBound(exts) To UBound(exts)
        If HasExtension(lowerHref, CStr(exts(i))) Then
            IsStaticAsset = True
            Exit Function
        End If
    Next i
    IsStaticAsset = False
End Function

Private Function HasExtension(ByVal lowerHref As String, ByVal ext As String) As Boolean
    Dim p As Long, nextCh As String
    p = InStr(1, lowerHref, ext)
    Do While p > 0
        ' do not treat ".jsp" as ".js"
        If ext = ".js" And mId$(lowerHref, p, 4) = ".jsp" Then
            p = InStr(p + 1, lowerHref, ext)
            GoTo NextExtMatch
        End If
        If p + Len(ext) > Len(lowerHref) Then
            HasExtension = True
            Exit Function
        End If
        nextCh = mId$(lowerHref, p + Len(ext), 1)
        If nextCh = "?" Or nextCh = "#" Or nextCh = "&" Or nextCh = "/" Or nextCh = ";" Then
            HasExtension = True
            Exit Function
        End If
        p = InStr(p + 1, lowerHref, ext)
NextExtMatch:
    Loop
    HasExtension = False
End Function

Private Function IsLikelyPostLink(ByVal href As String, ByVal includeHint As String, ByVal excludeHint As String) As Boolean
    Dim h As String
    h = LCase$(Trim$(href))
    If Len(h) = 0 Then Exit Function
    If Left$(h, 11) = "javascript:" Then Exit Function
    If Left$(h, 7) = "mailto:" Then Exit Function
    If h = "#" Or Left$(h, 1) = "#" Or Left$(h, 5) = "#none" Then Exit Function
    If InStr(h, "logout") > 0 Or InStr(h, "login") > 0 Then Exit Function
    If InStr(h, "filedown") > 0 Or InStr(h, "download") > 0 Then Exit Function
    If IsStaticAsset(h) Then Exit Function

    If Len(excludeHint) > 0 Then
        If InStr(1, h, LCase$(excludeHint), vbTextCompare) > 0 Then Exit Function
    End If

    If Len(includeHint) > 0 Then
        IsLikelyPostLink = (InStr(1, h, includeHint, vbTextCompare) > 0)
        Exit Function
    End If

    If InStr(h, "nttid=") > 0 Or InStr(h, "nttno=") > 0 Or InStr(h, "article") > 0 Then IsLikelyPostLink = True: Exit Function
    ' avoid matching mainView.do / commonSelectBoardList etc.
    If InStr(h, "mainview.do") > 0 Then Exit Function
    If InStr(h, "/view.do") > 0 Or InStr(h, "view.do?") > 0 Or InStr(h, "view.jsp") > 0 Or InStr(h, "boardview") > 0 Or InStr(h, "read.do") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "selectboardarticle") > 0 Or InStr(h, "commonselectboardarticle") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "boardcnts/view") > 0 Or InStr(h, "dtl.jsp") > 0 Or InStr(h, "idx=") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "mode=view") > 0 Or InStr(h, "act=view") > 0 Or InStr(h, "list_no=") > 0 Or InStr(h, "cntid=") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "bbs/") > 0 And (InStr(h, "view") > 0 Or InStr(h, "cntid=") > 0) Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "brd/") > 0 And InStr(h, "view") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "noticedetail") > 0 Or InStr(h, "noticedtl") > 0 Or InStr(h, "noticeview") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "subview.do") > 0 And InStr(h, "enc=") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "selectbbsnttview") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "viewrenew.do") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "selectdoc.do") > 0 Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "nw_ntc_s001d") > 0 Then IsLikelyPostLink = True: Exit Function
    ' unikorea / mpb style: /bbs/<boardId>/<numericId>?
    If InStr(h, "/bbs/") > 0 And IsBbsNumericDetail(h) Then IsLikelyPostLink = True: Exit Function
    If InStr(h, "bause=true") > 0 And InStr(h, "/bbs/") > 0 Then IsLikelyPostLink = True: Exit Function

    If InStr(h, "list.do") > 0 Or InStr(h, "list.jsp") > 0 Or InStr(h, "listrenew") > 0 Then Exit Function
    If InStr(h, "main.do") > 0 Or InStr(h, "index.do") > 0 Or InStr(h, "index.jsp") > 0 Then Exit Function

    IsLikelyPostLink = False
End Function

Private Function IsBbsNumericDetail(ByVal lowerHref As String) As Boolean
    Dim p As Long, q As Long, chunk As String, i As Long, c As String, digits As Long
    p = InStrRev(lowerHref, "/")
    If p = 0 Then Exit Function
    chunk = mId$(lowerHref, p + 1)
    q = InStr(chunk, "?")
    If q > 0 Then chunk = Left$(chunk, q - 1)
    If Len(chunk) = 0 Then Exit Function
    digits = 0
    For i = 1 To Len(chunk)
        c = mId$(chunk, i, 1)
        If c < "0" Or c > "9" Then Exit Function
        digits = digits + 1
    Next i
    IsBbsNumericDetail = (digits >= 3)
End Function

Private Function IsJunkTitle(ByVal title As String) As Boolean
    Dim t As String
    t = LCase$(Replace(title, " ", ""))
    If t = "more" Or t = "list" Or t = "prev" Or t = "next" Then IsJunkTitle = True: Exit Function
    If t = "home" Or t = "login" Or t = "print" Or t = "share" Then IsJunkTitle = True: Exit Function
    If Len(t) <= 1 Then IsJunkTitle = True: Exit Function
    IsJunkTitle = False
End Function

Private Function ExtractAttrValue(ByVal html As String, ByVal startPos As Long) As String
    Dim ch As String, i As Long, quote As String, buf As String
    i = startPos
    Do While i <= Len(html)
        ch = mId$(html, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        i = i + 1
    Loop
    If i > Len(html) Then Exit Function
    ch = mId$(html, i, 1)
    If ch = """" Or ch = "'" Then
        quote = ch
        i = i + 1
        Do While i <= Len(html)
            ch = mId$(html, i, 1)
            If ch = quote Then Exit Do
            buf = buf & ch
            i = i + 1
        Loop
    Else
        Do While i <= Len(html)
            ch = mId$(html, i, 1)
            If ch = " " Or ch = ">" Or ch = vbCr Or ch = vbLf Or ch = vbTab Then Exit Do
            buf = buf & ch
            i = i + 1
        Loop
    End If
    ExtractAttrValue = buf
End Function

Private Function GetTagAttrValue(ByVal html As String, ByVal tagStart As Long, ByVal tagEnd As Long, ByVal attrName As String) As String
    Dim tagHtml As String
    Dim lowerTag As String
    Dim attrPos As Long
    
    GetTagAttrValue = ""
    If tagStart <= 0 Or tagEnd <= tagStart Then Exit Function
    
    tagHtml = mId$(html, tagStart, tagEnd - tagStart + 1)
    lowerTag = LCase$(tagHtml)
    attrPos = InStr(1, lowerTag, LCase$(attrName) & "=", vbTextCompare)
    If attrPos = 0 Then Exit Function
    
    GetTagAttrValue = ExtractAttrValue(tagHtml, attrPos + Len(attrName) + 1)
End Function

Private Function ExtractLinkText(ByVal html As String, ByVal afterOpenTag As Long) As String
    Dim closePos As Long, inner As String, s As Long, e As Long
    closePos = InStr(afterOpenTag + 1, LCase$(html), "</a>")
    If closePos = 0 Then Exit Function
    inner = mId$(html, afterOpenTag + 1, closePos - afterOpenTag - 1)
    Do While InStr(inner, "<") > 0 And InStr(inner, ">") > 0
        s = InStr(inner, "<")
        e = InStr(s, inner, ">")
        If e = 0 Then Exit Do
        inner = Left$(inner, s - 1) & mId$(inner, e + 1)
    Loop
    ExtractLinkText = inner
End Function

Private Function CleanText(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, "&nbsp;", " ")
    t = Replace(t, "&amp;", "&")
    t = Replace(t, "&lt;", "<")
    t = Replace(t, "&gt;", ">")
    t = Replace(t, "&quot;", """")
    t = Replace(t, "&#39;", "'")
    t = Replace(t, "&#034;", """")
    t = Replace(t, "&apos;", "'")
    t = DecodeNumericEntities(t)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Trim$(t)
    ' strip leading "????" badge text from board lists
    If Left$(t, 2) = ChrW(&HC0C8&) & ChrW(&HAE00&) Then
        t = Trim$(mId$(t, 3))
    End If
    CleanText = t
End Function

Private Function DecodeNumericEntities(ByVal s As String) As String
    Dim result As String, i As Long, j As Long, code As Long, numStr As String
    result = ""
    i = 1
    Do While i <= Len(s)
        If mId$(s, i, 2) = "&#" Then
            j = InStr(i + 2, s, ";")
            If j > 0 And j - i < 10 Then
                numStr = mId$(s, i + 2, j - i - 2)
                On Error Resume Next
                code = CLng(numStr)
                If Err.Number <> 0 Then code = 0
                Err.Clear
                On Error GoTo 0
                If code > 0 And code < 65536 Then
                    result = result & ChrW(code)
                    i = j + 1
                Else
                    result = result & mId$(s, i, 1)
                    i = i + 1
                End If
            Else
                result = result & mId$(s, i, 1)
                i = i + 1
            End If
        Else
            result = result & mId$(s, i, 1)
            i = i + 1
        End If
    Loop
    DecodeNumericEntities = result
End Function

Private Function ToAbsoluteUrl(ByVal baseUrl As String, ByVal href As String) As String
    Dim h As String, base As String, scheme As String, host As String, slash As Long, i As Long
    Dim path As String, lastSlash As Long
    h = Trim$(href)
    If Len(h) = 0 Then Exit Function
    If InStr(1, h, "http://", vbTextCompare) = 1 Or InStr(1, h, "https://", vbTextCompare) = 1 Then
        ToAbsoluteUrl = h
        Exit Function
    End If
    If Left$(h, 2) = "//" Then
        If InStr(1, baseUrl, "https://", vbTextCompare) = 1 Then
            ToAbsoluteUrl = "https:" & h
        Else
            ToAbsoluteUrl = "http:" & h
        End If
        Exit Function
    End If

    i = InStr(1, baseUrl, "://")
    If i = 0 Then Exit Function
    scheme = Left$(baseUrl, i + 2)
    host = mId$(baseUrl, i + 3)
    slash = InStr(host, "/")
    If slash > 0 Then
        base = scheme & Left$(host, slash - 1)
    Else
        base = scheme & host
    End If

    If Left$(h, 1) = "?" Then
        Dim qpos As Long, bare As String
        qpos = InStr(baseUrl, "?")
        If qpos > 0 Then bare = Left$(baseUrl, qpos - 1) Else bare = baseUrl
        ToAbsoluteUrl = bare & h
        Exit Function
    End If

    If Left$(h, 1) = "/" Then
        ToAbsoluteUrl = base & h
    Else
        path = mId$(baseUrl, Len(base) + 1)
        If Len(path) = 0 Then path = "/"
        lastSlash = InStrRev(path, "/")
        If lastSlash > 0 Then path = Left$(path, lastSlash) Else path = "/"
        ToAbsoluteUrl = base & path & h
    End If
End Function

Private Function GuessDateNearLink(ByVal html As String, ByVal linkPos As Long) As String
    Dim startPos As Long, chunk As String, i As Long, buf As String, found As String
    startPos = linkPos - 120
    If startPos < 1 Then startPos = 1
    chunk = mId$(html, startPos, 400)
    For i = 1 To Len(chunk) - 9
        buf = mId$(chunk, i, 10)
        If mId$(buf, 5, 1) = "-" And mId$(buf, 8, 1) = "-" Then
            If IsDigits(Left$(buf, 4)) And IsDigits(mId$(buf, 6, 2)) And IsDigits(Right$(buf, 2)) Then
                found = buf
                Exit For
            End If
        ElseIf mId$(buf, 5, 1) = "." And mId$(buf, 8, 1) = "." Then
            If IsDigits(Left$(buf, 4)) And IsDigits(mId$(buf, 6, 2)) And IsDigits(Right$(buf, 2)) Then
                found = Left$(buf, 4) & "-" & mId$(buf, 6, 2) & "-" & Right$(buf, 2)
                Exit For
            End If
        ElseIf mId$(buf, 5, 1) = "/" And mId$(buf, 8, 1) = "/" Then
            If IsDigits(Left$(buf, 4)) And IsDigits(mId$(buf, 6, 2)) And IsDigits(Right$(buf, 2)) Then
                found = Left$(buf, 4) & "-" & mId$(buf, 6, 2) & "-" & Right$(buf, 2)
                Exit For
            End If
        End If
    Next i
    GuessDateNearLink = found
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = mId$(s, i, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next i
    IsDigits = True
End Function

Private Function SheetByName(ByVal n As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
End Function

Private Sub SetupResultBoardHeader(ByVal ws As Worksheet, ByVal collectTime As String)
    Dim rng As Range
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F").ColumnWidth = 2
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H").ColumnWidth = 40
    ws.Columns("I").ColumnWidth = 2
    ws.Columns("J").ColumnWidth = 3
    ws.Columns("K").ColumnWidth = 40

    Set rng = ws.Range("A1:K1")
    rng.Merge
    rng.Value = "Government Ministry Notices (Top 5) - 4 Columns"
    rng.Font.Bold = True
    rng.Font.Size = 16
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = RGB(31, 78, 121)
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    ws.Rows(1).RowHeight = 28

    Set rng = ws.Range("A2:K2")
    rng.Merge
    rng.Value = "Collected: " & collectTime & "   |   Title fits column width with ...   |   Click title to open in browser"
    rng.Font.Size = 9
    rng.Font.Color = RGB(89, 89, 89)
    rng.Interior.Color = RGB(242, 242, 242)
End Sub

Private Function Ellipsize(ByVal s As String, ByVal maxWidth As Long) As String
    Dim t As String, i As Long, w As Long, cw As Long, ch As Long
    Dim ellipsisW As Long
    t = Trim$(s)
    If maxWidth <= 0 Then
        Ellipsize = t
        Exit Function
    End If
    ellipsisW = 3
    w = 0
    For i = 1 To Len(t)
        ch = AscW(mId$(t, i, 1))
        If ch < 0 Then ch = ch + 65536
        If ch > 127 Then cw = 2 Else cw = 1
        If w + cw > maxWidth Then
            ' leave room for "..."
            Do While i > 1 And w + ellipsisW > maxWidth
                i = i - 1
                ch = AscW(mId$(t, i, 1))
                If ch < 0 Then ch = ch + 65536
                If ch > 127 Then w = w - 2 Else w = w - 1
            Loop
            Ellipsize = Left$(t, i - 1) & "..."
            Exit Function
        End If
        w = w + cw
    Next i
    Ellipsize = t
End Function

Private Sub WriteDeptSuccessCard(ByVal ws As Worksheet, ByVal startRow As Long, _
                                 ByVal startCol As Long, ByVal deptName As String, _
                                 ByVal posts As Collection)
    Dim r As Long, i As Long
    Dim p As Variant
    Dim title As String, link As String, shortTitle As String
    Dim rng As Range
    Dim endCol As Long
    Dim headColor As Long, lineColor As Long, bodyColor As Long

    endCol = startCol + 1
    r = startRow

    headColor = DeptHeaderColor(deptName)
    lineColor = DeptBorderColor(deptName)
    bodyColor = DeptBodyColor(deptName)

    Set rng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
    rng.Merge
    rng.Value = "[" & deptName & "]"
    rng.Font.Bold = True
    rng.Font.Size = 11
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = headColor
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = lineColor
    ws.Rows(r).RowHeight = 20
    r = r + 1

    For i = 1 To MAX_POSTS
        ws.Cells(r, startCol).Value = CStr(i) & "."
        ws.Cells(r, startCol).Font.Bold = True
        ws.Cells(r, startCol).HorizontalAlignment = xlCenter

        Set rng = ws.Range(ws.Cells(r, startCol + 1), ws.Cells(r, endCol))
        rng.Interior.Color = bodyColor
        rng.WrapText = False
        rng.ShrinkToFit = False

        If Not posts Is Nothing And i <= posts.Count Then
            p = posts(i)
            title = CStr(p(0))
            link = CStr(p(1))
            shortTitle = Ellipsize(title, BODY_TEXT_WIDTH)
            rng.Value = shortTitle
            If Len(link) > 0 Then
                On Error Resume Next
                ws.Hyperlinks.Add Anchor:=rng, Address:=link, TextToDisplay:=shortTitle
                On Error GoTo 0
            End If
            rng.Font.Color = RGB(5, 99, 193)
            rng.Font.Size = 9
            rng.WrapText = False
        Else
            rng.Value = ""
        End If

        ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.Color = lineColor
        ws.Rows(r).RowHeight = 15

        r = r + 1
    Next i
End Sub

Private Sub WriteDeptFailCard(ByVal ws As Worksheet, ByVal startRow As Long, _
                              ByVal startCol As Long, ByVal deptName As String)
    Dim r As Long, i As Long
    Dim rng As Range
    Dim endCol As Long
    Dim headColor As Long, lineColor As Long

    endCol = startCol + 1
    r = startRow

    headColor = DeptHeaderColor(deptName)
    lineColor = DeptBorderColor(deptName)

    Set rng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
    rng.Merge
    rng.Value = "[" & deptName & "]"
    rng.Font.Bold = True
    rng.Font.Size = 11
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = headColor
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = lineColor
    ws.Rows(r).RowHeight = 20
    r = r + 1

    Set rng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
    rng.Merge
    rng.Value = KoreanFixSettingsText()
    rng.Font.Bold = True
    rng.Font.Color = RGB(192, 0, 0)
    rng.Interior.Color = RGB(252, 228, 214)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(244, 177, 131)
    r = r + 1

    For i = 2 To MAX_POSTS
        ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.LineStyle = xlContinuous
        ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol)).Borders.Color = lineColor
        ws.Rows(r).RowHeight = 15
        r = r + 1
    Next i
End Sub

Private Sub ClearResultBoard(ByVal ws As Worksheet)
    Dim lastRow As Long
    On Error Resume Next
    Do While ws.Hyperlinks.Count > 0
        ws.Hyperlinks(1).Delete
    Loop
    ws.Cells.UnMerge
    On Error GoTo 0

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 20
    If lastRow < 50 Then lastRow = 50
    ws.Range(ws.Rows(1), ws.Rows(lastRow)).Clear
End Sub

Private Sub EnsureFetchButton(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim b As Object
    Dim leftPos As Single
    On Error Resume Next
    ws.Shapes("btnFetchNotices").Delete
    ws.Shapes("btnExportHtml").Delete
    For Each b In ws.Buttons()
        b.Delete
    Next b
    On Error GoTo 0

    leftPos = CSng(ws.Columns("K").Left)

    ' existing refresh button (right, above column K)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, 4, 150, 28)
    shp.Name = "btnFetchNotices"
    shp.OnAction = "'" & ThisWorkbook.Name & "'!FetchNotices"
    shp.TextFrame2.TextRange.Text = KoreanRefreshButtonText()
    shp.TextFrame2.TextRange.Font.Size = 11
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(39, 110, 241)
    shp.Line.ForeColor.RGB = RGB(27, 78, 176)
    shp.Adjustments.Item(1) = 0.18

    ' new dashboard export button (left of refresh)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos - 190, 4, 180, 28)
    shp.Name = "btnExportHtml"
    shp.OnAction = "'" & ThisWorkbook.Name & "'!ExportHtmlDashboard"
    shp.TextFrame2.TextRange.Text = KoreanDashboardButtonText()
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(0, 137, 123)
    shp.Line.ForeColor.RGB = RGB(0, 105, 92)
    shp.Adjustments.Item(1) = 0.18
End Sub

Public Sub ExportHtmlDashboardSilent()
    gSilent = True
    On Error GoTo Clean
    ExportHtmlDashboard
Clean:
    gSilent = False
End Sub

Public Sub ExportHtmlDashboard()
    Dim wsRes As Worksheet, wsCfg As Worksheet
    Dim outPathPc As String, outPathMobile As String
    Dim outPathPressPc As String, outPathPressMobile As String
    Dim templatePc As String, templateMobile As String
    Dim templatePressPc As String, templatePressMobile As String
    Dim jsonText As String, majorJson As String, outHtml As String
    Dim token As String, tokenMajor As String

    On Error GoTo Fatal
    Set wsRes = SheetByName(ChrW(&HACB0&) & ChrW(&HACFC&))
    Set wsCfg = SheetByName(ChrW(&HC124&) & ChrW(&HC815&))
    If wsRes Is Nothing Then Set wsRes = SheetByName("Result")
    If wsCfg Is Nothing Then Set wsCfg = SheetByName("Config")
    If wsRes Is Nothing Or wsCfg Is Nothing Then
        If Not gSilent Then MsgBox "Required sheets not found.", vbExclamation
        Exit Sub
    End If

    outPathPc = ThisWorkbook.path & Application.PathSeparator & "gov_notice_board.html"
    outPathMobile = ThisWorkbook.path & Application.PathSeparator & "gov_notice_board_mobile.html"
    outPathPressPc = ThisWorkbook.path & Application.PathSeparator & "gov_major_press_board.html"
    outPathPressMobile = ThisWorkbook.path & Application.PathSeparator & "gov_major_press_board_mobile.html"

    templatePc = LoadHtmlTemplateFile("gov_notice_board_pc_template.html", EmbeddedHtmlTemplatePc())
    templateMobile = LoadHtmlTemplateFile("gov_notice_board_mobile_template.html", EmbeddedHtmlTemplateMobile())
    templatePressPc = LoadHtmlTemplateFile("gov_major_press_pc_template.html", "")
    templatePressMobile = LoadHtmlTemplateFile("gov_major_press_mobile_template.html", "")

    If Len(templatePc) = 0 Or Len(templateMobile) = 0 Then
        If Not gSilent Then MsgBox "HTML notice template missing.", vbExclamation
        Exit Sub
    End If
    If Len(templatePressPc) = 0 Or Len(templatePressMobile) = 0 Then
        If Not gSilent Then MsgBox "Major press template missing (gov_major_press_*_template.html).", vbExclamation
        Exit Sub
    End If

    jsonText = BuildNoticeJsonFromSheets(wsCfg, wsRes)
    Application.StatusBar = "Fetching major press releases..."
    majorJson = BuildMajorPressJson(wsCfg, wsRes)
    Application.StatusBar = False


    token = "%%NOTICE_JSON%%"
    tokenMajor = "%%MAJOR_JSON%%"
    If InStr(1, templatePc, token, vbBinaryCompare) = 0 Or InStr(1, templateMobile, token, vbBinaryCompare) = 0 Then
        If Not gSilent Then MsgBox "NOTICE_JSON token missing.", vbExclamation
        Exit Sub
    End If
    If InStr(1, templatePressPc, tokenMajor, vbBinaryCompare) = 0 Or InStr(1, templatePressMobile, tokenMajor, vbBinaryCompare) = 0 Then
        If Not gSilent Then MsgBox "MAJOR_JSON token missing.", vbExclamation
        Exit Sub
    End If

    outHtml = Replace(templatePc, token, jsonText)
    WriteTextUtf8 outPathPc, outHtml
    outHtml = Replace(templateMobile, token, jsonText)
    WriteTextUtf8 outPathMobile, outHtml
    outHtml = Replace(templatePressPc, tokenMajor, majorJson)
    WriteTextUtf8 outPathPressPc, outHtml
    outHtml = Replace(templatePressMobile, tokenMajor, majorJson)
    WriteTextUtf8 outPathPressMobile, outHtml

    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=outPathPc
    On Error GoTo Fatal

    If Not gSilent Then MsgBox "HTML created: notice(PC/Mobile) + major press(PC/Mobile).", vbInformation
    Exit Sub
Fatal:
    Application.StatusBar = False
    If Not gSilent Then MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Private Function LoadHtmlTemplateFile(ByVal fileName As String, ByVal embeddedFallback As String) As String
    Dim p As String, t As String
    p = ThisWorkbook.path & Application.PathSeparator & fileName
    t = ""
    On Error Resume Next
    If Dir(p) <> "" Then t = ReadTextUtf8(p)
    On Error GoTo 0
    If Len(t) > 0 Then
        LoadHtmlTemplateFile = t
    Else
        LoadHtmlTemplateFile = embeddedFallback
    End If
End Function

Private Function BuildMajorPressJson(ByVal wsCfg As Worksheet, ByVal wsRes As Worksheet) As String
    Dim i As Long, j As Long
    Dim deptName As String, pressUrl As String, includeHint As String
    Dim html As String, errMsg As String
    Dim posts As Collection
    Dim noticesJson As String, pressJson As String
    Dim json As String
    Dim collected As String
    Dim p As Variant

    collected = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    json = "{""collectedAt"":""" & JsonEscape(collected) & """,""departments"":["

    For i = 1 To 4
        deptName = MajorDeptNameByIndex(i)
        pressUrl = MajorPressUrlByIndex(i)
        includeHint = MajorPressIncludeByIndex(i)
        noticesJson = BuildPostsJsonForDept(wsCfg, wsRes, deptName)
        pressJson = ""
        errMsg = ""
        html = ""
        Set posts = Nothing

        If Len(pressUrl) > 0 Then
            html = GetHtml(pressUrl, errMsg)
            If Len(errMsg) = 0 And Len(html) > 0 Then
                Set posts = ParseAnnouncementLinks(html, pressUrl, includeHint, "", MAX_POSTS)
                If (posts Is Nothing Or posts.Count = 0) And Len(includeHint) > 0 Then
                    Set posts = ParseAnnouncementLinks(html, pressUrl, "", "", MAX_POSTS)
                End If
            End If
        End If

        If Not posts Is Nothing Then
            For j = 1 To posts.Count
                If j > MAX_POSTS Then Exit For
                p = posts(j)
                If Len(pressJson) > 0 Then pressJson = pressJson & ","
                pressJson = pressJson & "{""title"":""" & JsonEscape(CStr(p(0))) & """,""url"":""" & JsonEscape(CStr(p(1))) & """"
                If UBound(p) >= 2 Then
                    If Len(Trim$(CStr(p(2)))) > 0 Then pressJson = pressJson & ",""date"":""" & JsonEscape(CStr(p(2))) & """"
                End If
                pressJson = pressJson & "}"
            Next j
        End If

        If i > 1 Then json = json & ","
        json = json & "{""name"":""" & JsonEscape(deptName) & """,""ok"":true,""notices"":[" & noticesJson & "],""press"":[" & pressJson & "]}"
    Next i

    json = json & "]}"
    BuildMajorPressJson = json
End Function

Private Function BuildPostsJsonForDept(ByVal wsCfg As Worksheet, ByVal wsRes As Worksheet, ByVal targetDept As String) As String
    Dim lastRow As Long, r As Long, i As Long
    Dim cardIndex As Long, startCol As Long, startRow As Long
    Dim deptName As String, title As String, link As String
    Dim postsJson As String
    Dim hl As Hyperlink

    postsJson = ""
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
    cardIndex = 0
    For r = 2 To lastRow
        deptName = Trim$(CStr(wsCfg.Cells(r, 1).Value))
        If Len(deptName) = 0 Then GoTo NextDept
        If deptName = targetDept Then
            startCol = CardStartCol(cardIndex)
            startRow = 3 + (cardIndex \ CARD_COLUMNS) * CARD_HEIGHT
            For i = 1 To MAX_POSTS
                title = Trim$(CStr(wsRes.Cells(startRow + i, startCol + 1).Value))
                link = ""
                On Error Resume Next
                For Each hl In wsRes.Hyperlinks
                    If hl.Range.Row = startRow + i And hl.Range.Column = startCol + 1 Then
                        link = CStr(hl.Address)
                        Exit For
                    End If
                Next hl
                On Error GoTo 0
                If Len(title) > 0 And Len(link) > 0 Then
                    If Len(postsJson) > 0 Then postsJson = postsJson & ","
                    postsJson = postsJson & "{""title"":""" & JsonEscape(title) & """,""url"":""" & JsonEscape(link) & """}"
                End If
            Next i
            BuildPostsJsonForDept = postsJson
            Exit Function
        End If
        cardIndex = cardIndex + 1
NextDept:
    Next r
    BuildPostsJsonForDept = postsJson
End Function

Private Function MajorDeptNameByIndex(ByVal idx As Long) As String
    Select Case idx
        Case 1
            MajorDeptNameByIndex = ChrW(&HC0B0&) & ChrW(&HC5C5&) & ChrW(&HD1B5&) & ChrW(&HC0C1&) & ChrW(&HBD80&)
        Case 2
            MajorDeptNameByIndex = ChrW(&HAE30&) & ChrW(&HD6C4&) & ChrW(&HC5D0&) & ChrW(&HB108&) & ChrW(&HC9C0&) & ChrW(&HD658&) & ChrW(&HACBD&) & ChrW(&HBD80&)
        Case 3
            MajorDeptNameByIndex = ChrW(&HACE0&) & ChrW(&HC6A9&) & ChrW(&HB178&) & ChrW(&HB3D9&) & ChrW(&HBD80&)
        Case Else
            MajorDeptNameByIndex = ChrW(&HC678&) & ChrW(&HAD50&) & ChrW(&HBD80&)
    End Select
End Function

Private Function MajorPressUrlByIndex(ByVal idx As Long) As String
    Select Case idx
        Case 1
            MajorPressUrlByIndex = "https://www.motir.go.kr/kor/article/ATCL3f49a5a8c"
        Case 2
            MajorPressUrlByIndex = "https://mcee.go.kr/home/web/index.do?menuId=10523"
        Case 3
            MajorPressUrlByIndex = "https://www.moel.go.kr/news/enews/report/enewsList.do"
        Case Else
            MajorPressUrlByIndex = "https://www.mofa.go.kr/www/brd/m_4080/list.do"
    End Select
End Function

Private Function MajorPressIncludeByIndex(ByVal idx As Long) As String
    Select Case idx
        Case 1
            MajorPressIncludeByIndex = "/view?"
        Case 2
            MajorPressIncludeByIndex = "newsRead"
        Case 3
            MajorPressIncludeByIndex = "enewsView"
        Case Else
            MajorPressIncludeByIndex = "view.do"
    End Select
End Function


Private Function BuildNoticeJsonFromSheets(ByVal wsCfg As Worksheet, ByVal wsRes As Worksheet) As String
    Dim lastRow As Long, r As Long, i As Long
    Dim cardIndex As Long, startCol As Long, startRow As Long
    Dim deptName As String, title As String, link As String
    Dim json As String, postsJson As String
    Dim okFlag As Boolean
    Dim collected As String
    Dim hl As Hyperlink

    collected = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error Resume Next
    collected = Trim$(CStr(wsRes.Range("A2").Value))
    If InStr(collected, ":") > 0 Then
        ' keep as-is if header already contains time text
    Else
        collected = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    End If
    On Error GoTo 0
    collected = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    json = "{""collectedAt"":""" & JsonEscape(collected) & """,""departments"":["
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
    cardIndex = 0

    For r = 2 To lastRow
        deptName = Trim$(CStr(wsCfg.Cells(r, 1).Value))
        If Len(deptName) = 0 Then GoTo NextDept

        startCol = CardStartCol(cardIndex)
        startRow = 3 + (cardIndex \ CARD_COLUMNS) * CARD_HEIGHT
        postsJson = ""
        okFlag = False

        For i = 1 To MAX_POSTS
            title = Trim$(CStr(wsRes.Cells(startRow + i, startCol + 1).Value))
            link = ""
            On Error Resume Next
            For Each hl In wsRes.Hyperlinks
                If hl.Range.Row = startRow + i And hl.Range.Column = startCol + 1 Then
                    link = CStr(hl.Address)
                    Exit For
                End If
            Next hl
            On Error GoTo 0

            If Len(title) > 0 And Len(link) > 0 Then
                okFlag = True
                If Len(postsJson) > 0 Then postsJson = postsJson & ","
                postsJson = postsJson & "{""title"":""" & JsonEscape(title) & """,""url"":""" & JsonEscape(link) & """}"
            End If
        Next i

        If cardIndex > 0 Then json = json & ","
        If okFlag Then
            json = json & "{""name"":""" & JsonEscape(deptName) & """,""ok"":true,""posts"":[" & postsJson & "]}"
        Else
            json = json & "{""name"":""" & JsonEscape(deptName) & """,""ok"":false,""posts"":[]}"
        End If

        cardIndex = cardIndex + 1
NextDept:
    Next r

    json = json & "]}"
    BuildNoticeJsonFromSheets = json
End Function

Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonEscape = t
End Function

Private Function ReadTextUtf8(ByVal filePath As String) As String
    Dim stm As Object
    On Error GoTo Fail
    If Dir(filePath) = "" Then Exit Function
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextUtf8 = stm.ReadText(-1)
    stm.Close
    Exit Function
Fail:
    ReadTextUtf8 = ""
End Function

Private Sub WriteTextUtf8(ByVal filePath As String, ByVal content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText content
    If Dir(filePath) <> "" Then Kill filePath
    stm.SaveToFile filePath, 2
    stm.Close
End Sub

Private Sub ClearSheetData(ByVal ws As Worksheet, ByVal startRow As Long)
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= startRow Then
        ws.Range(ws.Rows(startRow), ws.Rows(lastRow)).Clear
    End If
    On Error Resume Next
    Do While ws.Hyperlinks.Count > 0
        ws.Hyperlinks(1).Delete
    Loop
    On Error GoTo 0
End Sub

Private Sub WriteLog(ByVal wsLog As Worksheet, ByVal ministry As String, ByVal status As String, _
                     ByVal cnt As Long, ByVal msg As String, ByVal ts As String)
    If wsLog Is Nothing Then Exit Sub
    Dim r As Long
    r = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    wsLog.Cells(r, 1).Value = ts
    wsLog.Cells(r, 2).Value = ministry
    wsLog.Cells(r, 3).Value = status
    wsLog.Cells(r, 4).Value = cnt
    wsLog.Cells(r, 5).Value = msg
End Sub

Public Sub PingMacro()
    MsgBox "Macro OK", vbInformation
End Sub

Public Function DebugFetchHtml(ByVal url As String) As String
    Dim errMsg As String
    DebugFetchHtml = GetHtml(url, errMsg)
    If Len(errMsg) > 0 Then DebugFetchHtml = "ERROR:" & errMsg
End Function

Private Function CardStartCol(ByVal cardIndex As Long) As Long
    Select Case (cardIndex Mod CARD_COLUMNS)
        Case 0: CardStartCol = 1   ' A:B
        Case 1: CardStartCol = 4   ' D:E
        Case 2: CardStartCol = 7   ' G:H
        Case Else: CardStartCol = 10 ' J:K
    End Select
End Function

Private Function DeptHeaderColor(ByVal deptName As String) As Long
    If IsMotirDept(deptName) Then
        DeptHeaderColor = RGB(255, 140, 0)
        Exit Function
    End If
    Select Case (DeptColorIndex(deptName))
        Case 0: DeptHeaderColor = RGB(58, 123, 213)
        Case 1: DeptHeaderColor = RGB(102, 126, 234)
        Case 2: DeptHeaderColor = RGB(0, 150, 136)
        Case 3: DeptHeaderColor = RGB(255, 112, 67)
        Case 4: DeptHeaderColor = RGB(126, 87, 194)
        Case 5: DeptHeaderColor = RGB(46, 125, 50)
        Case 6: DeptHeaderColor = RGB(198, 40, 40)
        Case Else: DeptHeaderColor = RGB(0, 121, 191)
    End Select
End Function

Private Function DeptBorderColor(ByVal deptName As String) As Long
    If IsMotirDept(deptName) Then
        DeptBorderColor = RGB(230, 110, 0)
        Exit Function
    End If
    Select Case (DeptColorIndex(deptName))
        Case 0: DeptBorderColor = RGB(38, 90, 170)
        Case 1: DeptBorderColor = RGB(76, 96, 190)
        Case 2: DeptBorderColor = RGB(0, 121, 107)
        Case 3: DeptBorderColor = RGB(230, 81, 0)
        Case 4: DeptBorderColor = RGB(94, 53, 177)
        Case 5: DeptBorderColor = RGB(27, 94, 32)
        Case 6: DeptBorderColor = RGB(183, 28, 28)
        Case Else: DeptBorderColor = RGB(0, 90, 140)
    End Select
End Function

Private Function DeptBodyColor(ByVal deptName As String) As Long
    If IsMotirDept(deptName) Then
        DeptBodyColor = RGB(255, 243, 224)
        Exit Function
    End If
    Select Case (DeptColorIndex(deptName))
        Case 0: DeptBodyColor = RGB(239, 246, 255)
        Case 1: DeptBodyColor = RGB(243, 240, 255)
        Case 2: DeptBodyColor = RGB(232, 245, 233)
        Case 3: DeptBodyColor = RGB(255, 243, 224)
        Case 4: DeptBodyColor = RGB(245, 240, 255)
        Case 5: DeptBodyColor = RGB(237, 247, 237)
        Case 6: DeptBodyColor = RGB(255, 235, 238)
        Case Else: DeptBodyColor = RGB(236, 248, 255)
    End Select
End Function

Private Function DeptColorIndex(ByVal deptName As String) As Long
    Dim i As Long
    For i = 1 To Len(deptName)
        DeptColorIndex = (DeptColorIndex + AscW(mId$(deptName, i, 1))) Mod 8
    Next i
End Function

Private Function IsMotirDept(ByVal deptName As String) As Boolean
    ' ???????
    IsMotirDept = (deptName = ChrW(&HC0B0&) & ChrW(&HC5C5&) & ChrW(&HD1B5&) & ChrW(&HC0C1&) & ChrW(&HBD80&))
End Function

Private Function KoreanRefreshButtonText() As String
    KoreanRefreshButtonText = ChrW(&HACF5&) & ChrW(&HC9C0&) & " " & ChrW(&HC0C8&) & ChrW(&HB85C&) & ChrW(&HACE0&) & ChrW(&HCE68&)
End Function

Private Function KoreanDashboardButtonText() As String
    ' ??Æä???? ????º¸????¼º
    KoreanDashboardButtonText = ChrW(&HC6F9&) & ChrW(&HD398&) & ChrW(&HC774&) & ChrW(&HC9C0&) & " " & _
        ChrW(&HB300&) & ChrW(&HC2DC&) & ChrW(&HBCF4&) & ChrW(&HB4DC&) & " " & _
        ChrW(&HC0DD&) & ChrW(&HC131&)
End Function

Private Function KoreanFixSettingsText() As String
    KoreanFixSettingsText = ChrW(&HC124&) & ChrW(&HC815&) & " " & ChrW(&HD398&) & ChrW(&HC774&) & ChrW(&HC9C0&) & ChrW(&HB97C&) & " " & ChrW(&HC218&) & ChrW(&HC815&) & ChrW(&HD558&) & ChrW(&HC138&) & ChrW(&HC694&) & "."
End Function

Public Function DebugParseLinks(ByVal url As String, ByVal includeHint As String, ByVal excludeHint As String) As String
    Dim html As String, errMsg As String
    Dim posts As Collection
    Dim i As Long, result As String
    Dim p As Variant
    On Error GoTo DebugErr
    html = GetHtml(url, errMsg)
    If Len(errMsg) > 0 Then
        DebugParseLinks = "FETCH_ERR:" & errMsg
        Exit Function
    End If
    Set posts = ParseAnnouncementLinks(html, url, includeHint, excludeHint, 5)
    If posts Is Nothing Or posts.Count = 0 Then
        DebugParseLinks = "NO_LINKS html_len=" & CStr(Len(html))
        Exit Function
    End If
    result = "COUNT=" & CStr(posts.Count) & vbLf
    For i = 1 To posts.Count
        p = posts(i)
        result = result & CStr(i) & ". " & CStr(p(0)) & " | " & CStr(p(1)) & vbLf
    Next i
    DebugParseLinks = result
    Exit Function
DebugErr:
    DebugParseLinks = "VBA_ERR:" & Err.Number & " " & Err.Description
End Function

Public Function DebugFirstLink(ByVal url As String, ByVal includeHint As String) As String
    Dim html As String, errMsg As String
    Dim lowerHtml As String, pos As Long, aOpen As Long
    Dim hrefPos As Long, endTag As Long, href As String, title As String
    Dim result As String, cnt As Long
    On Error GoTo DFLErr
    html = GetHtml(url, errMsg)
    If Len(errMsg) > 0 Then DebugFirstLink = "FETCH_ERR:" & errMsg: Exit Function
    lowerHtml = LCase$(html)
    pos = 1: cnt = 0
    Do While cnt < 300

        cnt = cnt + 1
        aOpen = InStr(pos, lowerHtml, "<a ")
        If aOpen = 0 Then Exit Do
        hrefPos = InStr(aOpen, lowerHtml, "href=")
        endTag = InStr(aOpen, lowerHtml, ">")
        If hrefPos = 0 Or endTag = 0 Or hrefPos > endTag Then pos = aOpen + 2: GoTo DFLNext
        href = ExtractAttrValue(html, hrefPos + 5)
        If Len(includeHint) > 0 And InStr(1, href, includeHint, vbTextCompare) > 0 Then
            title = ExtractLinkText(html, endTag)
            title = CleanText(title)
            Dim absUrl As String
            absUrl = CleanUrl(ToAbsoluteUrl(url, href))
            result = "MATCH#" & cnt & " href_len=" & Len(href) & " title_len=" & Len(title) & " title=[" & Left$(title, 60) & "] abs_len=" & Len(absUrl) & " abs=[" & Left$(absUrl, 100) & "]"
            DebugFirstLink = result
            Exit Function
        End If
        pos = endTag + 1
DFLNext:
    Loop
    DebugFirstLink = "NO_MATCH after " & cnt & " tags"
    Exit Function
DFLErr:
    DebugFirstLink = "ERR:" & Err.Number & " " & Err.Description
End Function

Private Function EmbeddedHtmlTemplatePc() As String
    Dim b64 As String
    b64 = ""
    b64 = b64 & "PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAi"
    b64 = b64 & "IG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KCV67aA67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wgKFBDKTwvdGl0bGU+DQo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4udGFpbHdpbmRjc3MuY29tP3BsdWdpbnM9Zm9ybXMsY29udGFp"
    b64 = b64 & "bmVyLXF1ZXJpZXMiPjwvc2NyaXB0Pg0KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1QdWJsaWMrU2Fuczp3Z2h0QDQwMDs1MDA7NjAwOzcwMCZhbXA7ZGlzcGxheT1zd2FwIiByZWw9InN0"
    b64 = b64 & "eWxlc2hlZXQiLz4NCjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9TWF0ZXJpYWwrU3ltYm9scytPdXRsaW5lZDp3Z2h0LEZJTExAMTAwLi43MDAsMC4uMSZhbXA7ZGlzcGxheT1zd2FwIiBy"
    b64 = b64 & "ZWw9InN0eWxlc2hlZXQiLz4NCjxzY3JpcHQgaWQ9InRhaWx3aW5kLWNvbmZpZyI+DQogICAgICAgIHRhaWx3aW5kLmNvbmZpZyA9IHsNCiAgICAgICAgICAgIGRhcmtNb2RlOiAiY2xhc3MiLA0KICAgICAgICAgICAgdGhlbWU6IHsNCiAg"
    b64 = b64 & "ICAgICAgICAgICAgICBleHRlbmQ6IHsNCiAgICAgICAgICAgICAgICAgICAgImNvbG9ycyI6IHsNCiAgICAgICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLXZhcmlhbnQiOiAiI2UxZTJlZCIsDQogICAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAic3VyZmFjZS1icmlnaHQiOiAiI2ZhZjhmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1jb250YWluZXIiOiAiI2VlZWZmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1jb250YWluZXIi"
    b64 = b64 & "OiAiIzI1NjNlYiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tZXJyb3IiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS10aW50IjogIiMwMDUzZGIiLA0KICAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgInByaW1hcnkiOiAiIzAwNGFjNiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc2Vjb25kYXJ5LWNvbnRhaW5lciI6ICIjNWMyNDAwIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJ0ZXJ0aWFyeSI6ICIjOTQzNzAwIiwN"
    b64 = b64 & "CiAgICAgICAgICAgICAgICAgICAgICAgICJwcmltYXJ5LWZpeGVkLWRpbSI6ICIjYjRjNWZmIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1zZWNvbmRhcnktZml4ZWQiOiAiIzM0MTEwMCIsDQogICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAib3V0bGluZS12YXJpYW50IjogIiNjM2M2ZDciLA0KICAgICAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWZpeGVkIjogIiNmZmRiY2QiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1jb250YWluZXIi"
    b64 = b64 & "OiAiI2ZkNzYxYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1maXhlZCI6ICIjZGJlMWZmIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXN1cmZhY2UiOiAiIzJlMzAzOSIsDQogICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAib24tdGVydGlhcnkiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1kaW0iOiAiI2Q5ZDllNSIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWluZXItaGln"
    b64 = b64 & "aCI6ICIjZTdlN2YzIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1wcmltYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzAwM2VhOCIsDQogICAgICAgICAgICAgICAgICAgICAgICAiaW52ZXJzZS1vbi1zdXJmYWNlIjogIiNmMGYwZmIi"
    b64 = b64 & "LA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyIjogIiNlZGVkZjkiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeSI6ICIjOWQ0MzAwIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJv"
    b64 = b64 & "bi1wcmltYXJ5IjogIiNmZmZmZmYiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyLWxvdyI6ICIjZjNmM2ZlIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1zZWNvbmRhcnkiOiAiI2ZmZmZmZiIs"
    b64 = b64 & "DQogICAgICAgICAgICAgICAgICAgICAgICAic2Vjb25kYXJ5LWZpeGVkLWRpbSI6ICIjZmZiNjkwIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJ0ZXJ0aWFyeS1maXhlZC1kaW0iOiAiI2ZmYjU5NiIsDQogICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAib24tc3VyZmFjZS12YXJpYW50IjogIiM0MzQ2NTUiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UiOiAiI2ZhZjhmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic2Vjb25kYXJ5LWZpeGVkIjogIiNm"
    b64 = b64 & "ZmRiY2EiLA0KICAgICAgICAgICAgICAgICAgICAgICAgIm91dGxpbmUiOiAiIzczNzY4NiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktY29udGFpbmVyIjogIiNmZmVkZTYiLA0KICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgInN1cmZhY2UtY29udGFpbmVyLWhpZ2hlc3QiOiAiI2UxZTJlZCIsDQogICAgICAgICAgICAgICAgICAgICAgICAidGVydGlhcnktY29udGFpbmVyIjogIiNiYzQ4MDAiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZh"
    b64 = b64 & "Y2UtY29udGFpbmVyLWxvd2VzdCI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJlcnJvci1jb250YWluZXIiOiAiI2ZmZGFkNiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tYmFja2dyb3VuZCI6ICIjMTkx"
    b64 = b64 & "YjIzIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXByaW1hcnkiOiAiI2I0YzVmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc3VyZmFjZSI6ICIjMTkxYjIzIiwNCiAgICAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICJvbi10ZXJ0aWFyeS1maXhlZC12YXJpYW50IjogIiM3ZDJkMDAiLA0KICAgICAgICAgICAgICAgICAgICAgICAgImJhY2tncm91bmQiOiAiI2ZhZjhmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1maXhlZCI6"
    b64 = b64 & "ICIjMDAxNzRiIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1lcnJvci1jb250YWluZXIiOiAiIzkzMDAwYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc2Vjb25kYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzc4MzIwMCIs"
    b64 = b64 & "DQogICAgICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktZml4ZWQiOiAiIzM2MGYwMCIsDQogICAgICAgICAgICAgICAgICAgICAgICAiZXJyb3IiOiAiI2JhMWExYSINCiAgICAgICAgICAgICAgICAgICAgfSwNCiAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgImJvcmRlclJhZGl1cyI6IHsNCiAgICAgICAgICAgICAgICAgICAgICAgICJERUZBVUxUIjogIjAuMjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImxnIjogIjAuNXJlbSIsDQogICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAieGwiOiAiMC43NXJlbSIsDQogICAgICAgICAgICAgICAgICAgICAgICAiZnVsbCI6ICI5OTk5cHgiDQogICAgICAgICAgICAgICAgICAgIH0sDQogICAgICAgICAgICAgICAgICAgICJzcGFjaW5nIjogew0KICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgInN0YWNrLWdhcCI6ICIwLjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImNvbnRhaW5lci1wYWRkaW5nIjogIjJyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImNhcmQtcGFkZGluZyI6ICIx"
    b64 = b64 & "cmVtIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJncmlkLWd1dHRlciI6ICIxLjI1cmVtIg0KICAgICAgICAgICAgICAgICAgICB9LA0KICAgICAgICAgICAgICAgICAgICAiZm9udEZhbWlseSI6IHsNCiAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICJsaXN0LWl0ZW0iOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICAgICAiaGVhZGVyLXRpdGxlIjogWyJQdWJsaWMgU2FucyJdLA0KICAgICAgICAgICAgICAgICAgICAgICAgIm1ldGEtZGF0YSI6IFsi"
    b64 = b64 & "UHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgICAgICJib2FyZC10aXRsZSI6IFsiUHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgICAgICJidXR0b24tdGV4dCI6IFsiUHVibGljIFNhbnMiXQ0KICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICB9LA0KICAgICAgICAgICAgICAgICAgICAiZm9udFNpemUiOiB7DQogICAgICAgICAgICAgICAgICAgICAgICAibGlzdC1pdGVtIjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdodCI6"
    b64 = b64 & "ICI1MDAifV0sDQogICAgICAgICAgICAgICAgICAgICAgICAiaGVhZGVyLXRpdGxlIjogWyIyNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjMycHgiLCAibGV0dGVyU3BhY2luZyI6ICItMC4wMmVtIiwgImZvbnRXZWlnaHQiOiAiNzAwIn1dLA0K"
    b64 = b64 & "ICAgICAgICAgICAgICAgICAgICAgICAgIm1ldGEtZGF0YSI6IFsiMTJweCIsIHsibGluZUhlaWdodCI6ICIxNnB4IiwgImZvbnRXZWlnaHQiOiAiNDAwIn1dLA0KICAgICAgICAgICAgICAgICAgICAgICAgImJvYXJkLXRpdGxlIjogWyIx"
    b64 = b64 & "NnB4IiwgeyJsaW5lSGVpZ2h0IjogIjI0cHgiLCAiZm9udFdlaWdodCI6ICI3MDAifV0sDQogICAgICAgICAgICAgICAgICAgICAgICAiYnV0dG9uLXRleHQiOiBbIjE0cHgiLCB7ImxpbmVIZWlnaHQiOiAiMjBweCIsICJmb250V2VpZ2h0"
    b64 = b64 & "IjogIjYwMCJ9XQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfSwNCiAgICAgICAgICAgIH0sDQogICAgICAgIH0NCiAgICA8L3NjcmlwdD4NCjxzdHlsZT4NCiAgICAgICAgLm1hdGVyaWFsLXN5bWJvbHMtb3V0"
    b64 = b64 & "bGluZWQgew0KICAgICAgICAgICAgZm9udC12YXJpYXRpb24tc2V0dGluZ3M6ICdGSUxMJyAwLCAnd2dodCcgNDAwLCAnR1JBRCcgMCwgJ29wc3onIDI0Ow0KICAgICAgICAgICAgdmVydGljYWwtYWxpZ246IG1pZGRsZTsNCiAgICAgICAg"
    b64 = b64 & "fQ0KICAgICAgICBib2R5IHsgYmFja2dyb3VuZC1jb2xvcjogI2Y5ZmFmYjsgfQ0KICAgICAgICAuYm9hcmQtY2FyZCB7DQogICAgICAgICAgICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMC4ycyBlYXNlLCBib3gtc2hhZG93IDAuMnMgZWFz"
    b64 = b64 & "ZSwgb3BhY2l0eSAwLjEycyBlYXNlOw0KICAgICAgICAgICAgY3Vyc29yOiBncmFiOw0KICAgICAgICAgICAgdXNlci1zZWxlY3Q6IG5vbmU7DQogICAgICAgIH0NCiAgICAgICAgLmJvYXJkLWNhcmQ6YWN0aXZlIHsgY3Vyc29yOiBncmFi"
    b64 = b64 & "YmluZzsgfQ0KICAgICAgICAuYm9hcmQtY2FyZDpob3ZlciB7DQogICAgICAgICAgICB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTRweCk7DQogICAgICAgICAgICBib3gtc2hhZG93OiAwIDEwcHggMjVweCAtNXB4IHJnYmEoMCwgMCwgMCwg"
    b64 = b64 & "MC4xKSwgMCA4cHggMTBweCAtNnB4IHJnYmEoMCwgMCwgMCwgMC4xKTsNCiAgICAgICAgfQ0KICAgICAgICAuYm9hcmQtY2FyZC5kcmFnZ2luZyB7IG9wYWNpdHk6IDAuNDU7IHRyYW5zZm9ybTogc2NhbGUoMC45OCk7IH0NCiAgICAgICAg"
    b64 = b64 & "LmJvYXJkLWNhcmQuZHJhZy1vdmVyIHsgb3V0bGluZTogMnB4IGRhc2hlZCAjMjU2M2ViOyBvdXRsaW5lLW9mZnNldDogMnB4OyB9DQogICAgICAgIC5ib2FyZC1jYXJkIGEgeyBjdXJzb3I6IHBvaW50ZXI7IHVzZXItc2VsZWN0OiB0ZXh0"
    b64 = b64 & "OyB9DQogICAgICAgIC5maWx0ZXItYWN0aXZlIHsNCiAgICAgICAgICAgIGJhY2tncm91bmQtY29sb3I6ICNmZDc2MWEgIWltcG9ydGFudDsNCiAgICAgICAgICAgIGNvbG9yOiAjNWMyNDAwICFpbXBvcnRhbnQ7DQogICAgICAgIH0NCjwv"
    b64 = b64 & "c3R5bGU+DQo8L2hlYWQ+DQo8Ym9keSBjbGFzcz0iZm9udC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlIj4NCjxoZWFkZXIgY2xhc3M9ImJnLXByaW1hcnkgc2hhZG93LW1kIHN0aWNreSB0b3AtMCB6LTUwIHJvdW5kZWQtYi14bCI+DQo8"
    b64 = b64 & "ZGl2IGNsYXNzPSJtYXgtdy1bMTQ0MHB4XSBteC1hdXRvIHB4LWNvbnRhaW5lci1wYWRkaW5nIHB5LTYgZmxleCBmbGV4LWNvbCBtZDpmbGV4LXJvdyBqdXN0aWZ5LWJldHdlZW4gaXRlbXMtY2VudGVyIGdhcC00Ij4NCjxkaXYgY2xhc3M9"
    b64 = b64 & "ImZsZXggaXRlbXMtY2VudGVyIGdhcC0zIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtb24tcHJpbWFyeSB0ZXh0LTN4bCI+YWNjb3VudF9iYWxhbmNlPC9zcGFuPg0KPGRpdj4NCjxoMSBjbGFzcz0i"
    b64 = b64 & "Zm9udC1oZWFkZXItdGl0bGUgdGV4dC1oZWFkZXItdGl0bGUgdGV4dC1vbi1wcmltYXJ5Ij7soJXrtoDrtoDsspgg6rO17KeA7IKs7ZWtICjstZzqt7wgNeqwnCkgwrcgNOuLqCDrs7Trk5w8L2gxPg0KPHAgY2xhc3M9InRleHQtb24tcHJp"
    b64 = b64 & "bWFyeS84MCBmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSBtdC0xIj7subTrk5zrpbwg65Oc656Y6re47ZW0IOychOy5mOulvCDrsJTqv4Ag7IiYIOyeiOyKteuLiOuLpCDCtyDsoJzrqqkg7YG066atIOyLnCDsg4gg7LC9PC9wPg0K"
    b64 = b64 & "PC9kaXY+DQo8L2Rpdj4NCjxidXR0b24gaWQ9ImJ0blJlbG9hZCIgdHlwZT0iYnV0dG9uIiBjbGFzcz0iYmctcHJpbWFyeS1jb250YWluZXIgdGV4dC1vbi1wcmltYXJ5LWNvbnRhaW5lciBmb250LWJ1dHRvbi10ZXh0IHRleHQtYnV0dG9u"
    b64 = b64 & "LXRleHQgcHgtNiBweS0zIHJvdW5kZWQtZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBzaGFkb3ctbGcgaG92ZXI6YmctcHJpbWFyeS1jb250YWluZXIvOTAgYWN0aXZlOnNjYWxlLTk1IHRyYW5zaXRpb24tYWxsIj4NCjxzcGFuIGNs"
    b64 = b64 & "YXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5yZWZyZXNoPC9zcGFuPg0K7ZmU66m0IOyDiOuhnOqzoOy5qA0KPC9idXR0b24+DQo8L2Rpdj4NCjwvaGVhZGVyPg0KDQo8ZGl2IGNsYXNzPSJtYXgtdy1bMTQ0MHB4XSBteC1hdXRv"
    b64 = b64 & "IHB4LWNvbnRhaW5lci1wYWRkaW5nIG10LTQiPg0KPGRpdiBjbGFzcz0iYmctc3VyZmFjZS1jb250YWluZXIgcm91bmRlZC1sZyBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBweC00IHB5LTIgZmxleCBpdGVtcy1jZW50ZXIgZ2Fw"
    b64 = b64 & "LTIiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1wcmltYXJ5IHRleHQtc20iPnNjaGVkdWxlPC9zcGFuPg0KPHAgY2xhc3M9ImZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIHRleHQtb24tc3Vy"
    b64 = b64 & "ZmFjZS12YXJpYW50IiBpZD0ibWV0YUxpbmUiPuuNsOydtO2EsCDspIDruYQg7KSR4oCmPC9wPg0KPC9kaXY+DQo8L2Rpdj4NCg0KPGRpdiBjbGFzcz0ibWF4LXctWzE0NDBweF0gbXgtYXV0byBweC1jb250YWluZXItcGFkZGluZyBweS04"
    b64 = b64 & "IGZsZXggZ2FwLTgiPg0KPGFzaWRlIGNsYXNzPSJoaWRkZW4gbGc6ZmxleCBmbGV4LWNvbCB3LTY0IHNocmluay0wIGdhcC00Ij4NCjxkaXYgY2xhc3M9ImJnLXN1cmZhY2UtY29udGFpbmVyIHJvdW5kZWQteGwgcC00IGJvcmRlciBib3Jk"
    b64 = b64 & "ZXItb3V0bGluZS12YXJpYW50Ij4NCjxoMiBjbGFzcz0iZm9udC1ib2FyZC10aXRsZSB0ZXh0LWJvYXJkLXRpdGxlIHRleHQtcHJpbWFyeSBtYi00IGZsZXggaXRlbXMtY2VudGVyIGdhcC0yIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1z"
    b64 = b64 & "eW1ib2xzLW91dGxpbmVkIj5maWx0ZXJfYWx0PC9zcGFuPg0K67aA7LKYIO2VhO2EsA0KPC9oMj4NCjxuYXYgY2xhc3M9InNwYWNlLXktMSIgaWQ9ImZpbHRlck5hdiI+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImFs"
    b64 = b64 & "bCIgY2xhc3M9ImZpbHRlci1idG4gZmlsdGVyLWFjdGl2ZSB3LWZ1bGwgZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTMgcC0yIHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNz"
    b64 = b64 & "PSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5kYXNoYm9hcmQ8L3NwYW4+7KCE7LK0IOu2gOyymA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImVjb25vbXkiIGNsYXNzPSJmaWx0ZXItYnRuIHct"
    b64 = b64 & "ZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1j"
    b64 = b64 & "b2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5idXNpbmVzc19jZW50ZXI8L3NwYW4+7IKw7JeFL+qyveygnA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1m"
    b64 = b64 & "aWx0ZXI9InNvY2lldHkiIGNsYXNzPSJmaWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0g"
    b64 = b64 & "dGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj50aGVhdGVyX2NvbWVkeTwvc3Bhbj7sgqztmowv66y47ZmUDQo8L2J1"
    b64 = b64 & "dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0iYWRtaW4iIGNsYXNzPSJmaWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQt"
    b64 = b64 & "bGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5hY2NvdW50"
    b64 = b64 & "X2JhbGFuY2U8L3NwYW4+7ZaJ7KCVL+yViOyghA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImRpcGxvbWFjeSIgY2xhc3M9ImZpbHRlci1idG4gdy1mdWxsIGZsZXggaXRlbXMtY2VudGVyIGdhcC0z"
    b64 = b64 & "IHAtMiBob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCB0cmFuc2l0aW9uLWNvbG9ycyB0ZXh0LWxlZnQiPg0KPHNwYW4gY2xh"
    b64 = b64 & "c3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnB1YmxpYzwvc3Bhbj7smbjqtZAv7JWI67O0DQo8L2J1dHRvbj4NCjwvbmF2Pg0KPC9kaXY+DQo8ZGl2IGNsYXNzPSJiZy1zdXJmYWNlLWNvbnRhaW5lci1sb3cgcm91bmRlZC14bCBw"
    b64 = b64 & "LTQgYm9yZGVyIGJvcmRlci1vdXRsaW5lLXZhcmlhbnQgZmxleCBmbGV4LWNvbCBpdGVtcy1jZW50ZXIgdGV4dC1jZW50ZXIiPg0KPGRpdiBjbGFzcz0idy0xNiBoLTE2IGJnLXByaW1hcnkvMTAgcm91bmRlZC1mdWxsIGZsZXggaXRlbXMt"
    b64 = b64 & "Y2VudGVyIGp1c3RpZnktY2VudGVyIG1iLTMiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1wcmltYXJ5IHRleHQtM3hsIj5pbmZvPC9zcGFuPg0KPC9kaXY+DQo8cCBjbGFzcz0iZm9udC1tZXRhLWRh"
    b64 = b64 & "dGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgbWItNCI+7JeR7IWA7JeQ7IScIOyImOynke2VnCDstZzsi6Ag6rO17KeA7IKs7ZWt7J2EIOuztOuTnCDtmJXtg5zroZwg67O07Jes7KSN64uI64ukLjwvcD4NCjxh"
    b64 = b64 & "IGhyZWY9Imdvdl9ub3RpY2VfYm9hcmRfbW9iaWxlLmh0bWwiIGNsYXNzPSJ3LWZ1bGwgcHktMiBiZy1vdXRsaW5lLXZhcmlhbnQgdGV4dC1vbi1zdXJmYWNlIGZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4dCByb3VuZGVkLWxn"
    b64 = b64 & "IGhvdmVyOmJnLW91dGxpbmUgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1jZW50ZXIgbm8tdW5kZXJsaW5lIj7rqqjrsJTsnbwg7Y6Y7J207KeAPC9hPg0KPC9kaXY+DQo8L2FzaWRlPg0KDQo8bWFpbiBjbGFzcz0iZmxleC1ncm93Ij4NCjxk"
    b64 = b64 & "aXYgY2xhc3M9ImdyaWQgZ3JpZC1jb2xzLTEgbWQ6Z3JpZC1jb2xzLTIgeGw6Z3JpZC1jb2xzLTMgMnhsOmdyaWQtY29scy00IGdhcC1ncmlkLWd1dHRlciIgaWQ9ImJvYXJkIj48L2Rpdj4NCjwvbWFpbj4NCjwvZGl2Pg0KDQo8Zm9vdGVy"
    b64 = b64 & "IGNsYXNzPSJiZy1zdXJmYWNlLWRpbSBib3JkZXItdCBib3JkZXItb3V0bGluZS12YXJpYW50IG10LTEyIHB5LTggcHgtY29udGFpbmVyLXBhZGRpbmciPg0KPGRpdiBjbGFzcz0ibWF4LXctWzE0NDBweF0gbXgtYXV0byBmbGV4IGZsZXgt"
    b64 = b64 & "Y29sIG1kOmZsZXgtcm93IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgZ2FwLTYiPg0KPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4"
    b64 = b64 & "dC1wcmltYXJ5Ij5hY2NvdW50X2JhbGFuY2U8L3NwYW4+DQo8c3BhbiBjbGFzcz0iZm9udC1ib2xkIHRleHQtb24tc3VyZmFjZS12YXJpYW50Ij7rjIDtlZzrr7zqta0g7KCV67aAIOu2gOyymCDqs7Xsp4Dsgqztla0g67O065OcPC9zcGFu"
    b64 = b64 & "Pg0KPC9kaXY+DQo8bmF2IGNsYXNzPSJmbGV4IGdhcC02Ij4NCjxhIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBob3Zlcjp0ZXh0LXByaW1hcnkgdHJhbnNpdGlvbi1jb2xv"
    b64 = b64 & "cnMiIGhyZWY9Imdvdl9ub3RpY2VfYm9hcmQuaHRtbCI+UEM8L2E+DQo8YSBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgaG92ZXI6dGV4dC1wcmltYXJ5IHRyYW5zaXRpb24t"
    b64 = b64 & "Y29sb3JzIiBocmVmPSJnb3Zfbm90aWNlX2JvYXJkX21vYmlsZS5odG1sIj7rqqjrsJTsnbw8L2E+DQo8L25hdj4NCjxwIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudC83MCI+"
    b64 = b64 & "wqkg64yA7ZWc66+86rWtIOygleu2gCDrtoDsspgg6rO17KeA7IKs7ZWtIOuztOuTnCDshJzruYTsiqQ8L3A+DQo8L2Rpdj4NCjwvZm9vdGVyPg0KDQo8c2NyaXB0IGlkPSJub3RpY2UtZGF0YSIgdHlwZT0iYXBwbGljYXRpb24vanNvbiI+"
    b64 = b64 & "DQolJU5PVElDRV9KU09OJSUNCjwvc2NyaXB0Pg0KPHNjcmlwdD4NCihmdW5jdGlvbiAoKSB7DQogIGNvbnN0IFNUT1JBR0VfS0VZID0gJ2dvdk5vdGljZUNhcmRPcmRlci52MSc7DQogIGNvbnN0IFBBTEVUVEUgPSBbJyMzYjgyZjYnLCcj"
    b64 = b64 & "NmM3ZmQ4JywnIzAwOTY4OCcsJyNmZDc2MWEnLCcjN2U1N2MyJywnIzJlN2QzMicsJyNmZjZiNmInLCcjMDA5N2E3JywnIzAwYTg4NCcsJyMwMDRhYzYnXTsNCiAgY29uc3QgQ0FURUdPUlkgPSB7DQogICAgJ+yerOygleqyveygnOu2gCc6"
    b64 = b64 & "J2Vjb25vbXknLCfqs7ztlZnquLDsiKDsoJXrs7TthrXsi6DrtoAnOidlY29ub215Jywn7IKw7JeF7Ya17IOB67aAJzonZWNvbm9teScsJ+ykkeyGjOuypOyymOq4sOyXheu2gCc6J2Vjb25vbXknLA0KICAgICfqta3thqDqtZDthrXrtoAn"
    b64 = b64 & "OidlY29ub215Jywn64aN66a87LaV7IKw7Iud7ZKI67aAJzonZWNvbm9teScsJ+2VtOyWkeyImOyCsOu2gCc6J2Vjb25vbXknLCfquLDtm4Tsl5DrhIjsp4DtmZjqsr3rtoAnOidlY29ub215Jywn6riw7ZqN7JiI7IKw7LKYJzonZWNvbm9t"
    b64 = b64 & "eScsDQogICAgJ+q1kOycoeu2gCc6J3NvY2lldHknLCfrs7TqsbTrs7Xsp4DrtoAnOidzb2NpZXR5Jywn6rOg7Jqp64W464+Z67aAJzonc29jaWV0eScsJ+usuO2ZlOyytOycoeq0gOq0keu2gCc6J3NvY2lldHknLCfshLHtj4nrk7HqsIDs"
    b64 = b64 & "obHrtoAnOidzb2NpZXR5Jywn7Iud7ZKI7J2Y7JW97ZKI7JWI7KCE7LKYJzonc29jaWV0eScsDQogICAgJ+2WieygleyViOyghOu2gCc6J2FkbWluJywn67KV66y067aAJzonYWRtaW4nLCfsnbjsgqztmIHsi6DsspgnOidhZG1pbicsJ+uy"
    b64 = b64 & "leygnOyymCc6J2FkbWluJywn6rWt6rCA67O07ZuI67aAJzonYWRtaW4nLA0KICAgICfsmbjqtZDrtoAnOidkaXBsb21hY3knLCfthrXsnbzrtoAnOidkaXBsb21hY3knLCfqta3rsKnrtoAnOidkaXBsb21hY3knDQogIH07DQogIGxldCBk"
    b64 = b64 & "cmFnU3JjID0gbnVsbDsNCiAgbGV0IGN1cnJlbnRGaWx0ZXIgPSAnYWxsJzsNCiAgbGV0IGFsbERlcHRzID0gW107DQoNCiAgZnVuY3Rpb24gY29sb3JGb3IobmFtZSkgew0KICAgIGlmIChuYW1lID09PSAn7IKw7JeF7Ya17IOB67aAJykg"
    b64 = b64 & "cmV0dXJuICcjZmQ3NjFhJzsNCiAgICBsZXQgaWR4ID0gMDsNCiAgICBmb3IgKGxldCBpID0gMDsgaSA8IG5hbWUubGVuZ3RoOyBpKyspIGlkeCA9IChpZHggKyBuYW1lLmNoYXJDb2RlQXQoaSkpICUgUEFMRVRURS5sZW5ndGg7DQogICAg"
    b64 = b64 & "cmV0dXJuIFBBTEVUVEVbaWR4XTsNCiAgfQ0KICBmdW5jdGlvbiBjYXRlZ29yeUZvcihuYW1lKSB7IHJldHVybiBDQVRFR09SWVtuYW1lXSB8fCAnYWRtaW4nOyB9DQogIGZ1bmN0aW9uIGVzYyhzKSB7DQogICAgcmV0dXJuIFN0cmluZyhz"
    b64 = b64 & "ID8/ICcnKS5yZXBsYWNlKC8mL2csJyZhbXA7JykucmVwbGFjZSgvPC9nLCcmbHQ7JykucmVwbGFjZSgvPi9nLCcmZ3Q7JykucmVwbGFjZSgvIi9nLCcmcXVvdDsnKTsNCiAgfQ0KICBmdW5jdGlvbiBzYXZlT3JkZXIoKSB7DQogICAgY29u"
    b64 = b64 & "c3QgbmFtZXMgPSBBcnJheS5mcm9tKGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNib2FyZCAuYm9hcmQtY2FyZCcpKS5tYXAoYyA9PiBjLmdldEF0dHJpYnV0ZSgnZGF0YS1uYW1lJykpOw0KICAgIHRyeSB7IGxvY2FsU3RvcmFnZS5z"
    b64 = b64 & "ZXRJdGVtKFNUT1JBR0VfS0VZLCBKU09OLnN0cmluZ2lmeShuYW1lcykpOyB9IGNhdGNoIChlKSB7fQ0KICB9DQogIGZ1bmN0aW9uIGFwcGx5U2F2ZWRPcmRlcihkZXB0cykgew0KICAgIHRyeSB7DQogICAgICBjb25zdCByYXcgPSBsb2Nh"
    b64 = b64 & "bFN0b3JhZ2UuZ2V0SXRlbShTVE9SQUdFX0tFWSk7DQogICAgICBpZiAoIXJhdykgcmV0dXJuIGRlcHRzOw0KICAgICAgY29uc3Qgb3JkZXIgPSBKU09OLnBhcnNlKHJhdyk7DQogICAgICBpZiAoIUFycmF5LmlzQXJyYXkob3JkZXIpIHx8"
    b64 = b64 & "ICFvcmRlci5sZW5ndGgpIHJldHVybiBkZXB0czsNCiAgICAgIGNvbnN0IG1hcCA9IHt9Ow0KICAgICAgZGVwdHMuZm9yRWFjaChkID0+IHsgbWFwW2QubmFtZV0gPSBkOyB9KTsNCiAgICAgIGNvbnN0IHNvcnRlZCA9IFtdOw0KICAgICAg"
    b64 = b64 & "b3JkZXIuZm9yRWFjaChuYW1lID0+IHsgaWYgKG1hcFtuYW1lXSkgeyBzb3J0ZWQucHVzaChtYXBbbmFtZV0pOyBkZWxldGUgbWFwW25hbWVdOyB9IH0pOw0KICAgICAgT2JqZWN0LmtleXMobWFwKS5mb3JFYWNoKGsgPT4gc29ydGVkLnB1"
    b64 = b64 & "c2gobWFwW2tdKSk7DQogICAgICByZXR1cm4gc29ydGVkOw0KICAgIH0gY2F0Y2ggKGUpIHsgcmV0dXJuIGRlcHRzOyB9DQogIH0NCiAgZnVuY3Rpb24gYmluZERyYWcoY2FyZCkgew0KICAgIGNhcmQuc2V0QXR0cmlidXRlKCdkcmFnZ2Fi"
    b64 = b64 & "bGUnLCAndHJ1ZScpOw0KICAgIGNhcmQuYWRkRXZlbnRMaXN0ZW5lcignZHJhZ3N0YXJ0JywgZnVuY3Rpb24gKGUpIHsNCiAgICAgIGlmIChlLnRhcmdldCAmJiBlLnRhcmdldC5jbG9zZXN0ICYmIGUudGFyZ2V0LmNsb3Nlc3QoJ2EnKSkg"
    b64 = b64 & "eyBlLnByZXZlbnREZWZhdWx0KCk7IHJldHVybjsgfQ0KICAgICAgZHJhZ1NyYyA9IGNhcmQ7DQogICAgICBjYXJkLmNsYXNzTGlzdC5hZGQoJ2RyYWdnaW5nJyk7DQogICAgICBlLmRhdGFUcmFuc2Zlci5lZmZlY3RBbGxvd2VkID0gJ21v"
    b64 = b64 & "dmUnOw0KICAgICAgdHJ5IHsgZS5kYXRhVHJhbnNmZXIuc2V0RGF0YSgndGV4dC9wbGFpbicsIGNhcmQuZ2V0QXR0cmlidXRlKCdkYXRhLW5hbWUnKSB8fCAnJyk7IH0gY2F0Y2ggKGVycikge30NCiAgICB9KTsNCiAgICBjYXJkLmFkZEV2"
    b64 = b64 & "ZW50TGlzdGVuZXIoJ2RyYWdlbmQnLCBmdW5jdGlvbiAoKSB7DQogICAgICBjYXJkLmNsYXNzTGlzdC5yZW1vdmUoJ2RyYWdnaW5nJyk7DQogICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuYm9hcmQtY2FyZC5kcmFnLW92ZXIn"
    b64 = b64 & "KS5mb3JFYWNoKGVsID0+IGVsLmNsYXNzTGlzdC5yZW1vdmUoJ2RyYWctb3ZlcicpKTsNCiAgICAgIGRyYWdTcmMgPSBudWxsOw0KICAgICAgc2F2ZU9yZGVyKCk7DQogICAgfSk7DQogICAgY2FyZC5hZGRFdmVudExpc3RlbmVyKCdkcmFn"
    b64 = b64 & "b3ZlcicsIGZ1bmN0aW9uIChlKSB7DQogICAgICBlLnByZXZlbnREZWZhdWx0KCk7DQogICAgICBlLmRhdGFUcmFuc2Zlci5kcm9wRWZmZWN0ID0gJ21vdmUnOw0KICAgICAgaWYgKGRyYWdTcmMgJiYgZHJhZ1NyYyAhPT0gY2FyZCkgY2Fy"
    b64 = b64 & "ZC5jbGFzc0xpc3QuYWRkKCdkcmFnLW92ZXInKTsNCiAgICB9KTsNCiAgICBjYXJkLmFkZEV2ZW50TGlzdGVuZXIoJ2RyYWdsZWF2ZScsIGZ1bmN0aW9uICgpIHsgY2FyZC5jbGFzc0xpc3QucmVtb3ZlKCdkcmFnLW92ZXInKTsgfSk7DQog"
    b64 = b64 & "ICAgY2FyZC5hZGRFdmVudExpc3RlbmVyKCdkcm9wJywgZnVuY3Rpb24gKGUpIHsNCiAgICAgIGUucHJldmVudERlZmF1bHQoKTsNCiAgICAgIGNhcmQuY2xhc3NMaXN0LnJlbW92ZSgnZHJhZy1vdmVyJyk7DQogICAgICBpZiAoIWRyYWdT"
    b64 = b64 & "cmMgfHwgZHJhZ1NyYyA9PT0gY2FyZCkgcmV0dXJuOw0KICAgICAgY29uc3QgYm9hcmQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYm9hcmQnKTsNCiAgICAgIGNvbnN0IGNhcmRzID0gQXJyYXkuZnJvbShib2FyZC5jaGlsZHJlbik7"
    b64 = b64 & "DQogICAgICBjb25zdCBmcm9tID0gY2FyZHMuaW5kZXhPZihkcmFnU3JjKTsNCiAgICAgIGNvbnN0IHRvID0gY2FyZHMuaW5kZXhPZihjYXJkKTsNCiAgICAgIGlmIChmcm9tIDwgMCB8fCB0byA8IDApIHJldHVybjsNCiAgICAgIGlmIChm"
    b64 = b64 & "cm9tIDwgdG8pIGJvYXJkLmluc2VydEJlZm9yZShkcmFnU3JjLCBjYXJkLm5leHRTaWJsaW5nKTsNCiAgICAgIGVsc2UgYm9hcmQuaW5zZXJ0QmVmb3JlKGRyYWdTcmMsIGNhcmQpOw0KICAgICAgc2F2ZU9yZGVyKCk7DQogICAgfSk7DQog"
    b64 = b64 & "IH0NCiAgZnVuY3Rpb24gcmVuZGVyKCkgew0KICAgIGNvbnN0IGJvYXJkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2JvYXJkJyk7DQogICAgbGV0IGRlcHRzID0gYWxsRGVwdHMuc2xpY2UoKTsNCiAgICBpZiAoY3VycmVudEZpbHRl"
    b64 = b64 & "ciAhPT0gJ2FsbCcpIHsNCiAgICAgIGRlcHRzID0gZGVwdHMuZmlsdGVyKGQgPT4gY2F0ZWdvcnlGb3IoZC5uYW1lKSA9PT0gY3VycmVudEZpbHRlcik7DQogICAgfQ0KICAgIGRlcHRzID0gYXBwbHlTYXZlZE9yZGVyKGRlcHRzKTsNCiAg"
    b64 = b64 & "ICBpZiAoIWRlcHRzLmxlbmd0aCkgew0KICAgICAgYm9hcmQuaW5uZXJIVE1MID0gJzxkaXYgY2xhc3M9ImNvbC1zcGFuLWZ1bGwgdGV4dC1jZW50ZXIgcHktMTYgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgYmctc3VyZmFjZS1jb250YWlu"
    b64 = b64 & "ZXIgcm91bmRlZC14bCBib3JkZXIgYm9yZGVyLWRhc2hlZCBib3JkZXItb3V0bGluZS12YXJpYW50Ij7tkZzsi5ztlaAg6rO17KeAIOuNsOydtO2EsOqwgCDsl4bsirXri4jri6QuPC9kaXY+JzsNCiAgICAgIHJldHVybjsNCiAgICB9DQog"
    b64 = b64 & "ICAgYm9hcmQuaW5uZXJIVE1MID0gZGVwdHMubWFwKGZ1bmN0aW9uIChkKSB7DQogICAgICBjb25zdCBoZWFkID0gY29sb3JGb3IoZC5uYW1lKTsNCiAgICAgIGNvbnN0IGNhdCA9IGNhdGVnb3J5Rm9yKGQubmFtZSk7DQogICAgICBpZiAo"
    b64 = b64 & "ZC5vayA9PT0gZmFsc2UgfHwgIWQucG9zdHMgfHwgIWQucG9zdHMubGVuZ3RoKSB7DQogICAgICAgIHJldHVybiAnPGFydGljbGUgY2xhc3M9ImJvYXJkLWNhcmQgYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0IHJvdW5kZWQteGwgc2hh"
    b64 = b64 & "ZG93LVswcHhfNHB4XzEycHhfcmdiYSgwLDAsMCwwLjA1KV0gYm9yZGVyIGJvcmRlci1vdXRsaW5lLXZhcmlhbnQgZmxleCBmbGV4LWNvbCBoLWZ1bGwgb3ZlcmZsb3ctaGlkZGVuIiBkYXRhLW5hbWU9IicgKyBlc2MoZC5uYW1lKSArICci"
    b64 = b64 & "IGRhdGEtY2F0ZWdvcnk9IicgKyBjYXQgKyAnIj4nICsNCiAgICAgICAgICAnPGhlYWRlciBjbGFzcz0icC00IGZsZXgganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLWNlbnRlciIgc3R5bGU9ImJhY2tncm91bmQ6JyArIGhlYWQgKyAnIj48aDMg"
    b64 = b64 & "Y2xhc3M9ImZvbnQtYm9hcmQtdGl0bGUgdGV4dC1ib2FyZC10aXRsZSB0ZXh0LXdoaXRlIj5bJyArIGVzYyhkLm5hbWUpICsgJ108L2gzPjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtd2hpdGUgdGV4dC1z"
    b64 = b64 & "bSI+bW9yZV92ZXJ0PC9zcGFuPjwvaGVhZGVyPicgKw0KICAgICAgICAgICc8ZGl2IGNsYXNzPSJwLTQgdGV4dC1lcnJvciBmb250LWJvbGQgdGV4dC1zbSBiZy1lcnJvci1jb250YWluZXIiPuyEpOyglSDtjpjsnbTsp4Drpbwg7IiY7KCV"
    b64 = b64 & "7ZWY7IS47JqULjwvZGl2PjwvYXJ0aWNsZT4nOw0KICAgICAgfQ0KICAgICAgY29uc3QgaXRlbXMgPSBkLnBvc3RzLnNsaWNlKDAsIDUpLm1hcChmdW5jdGlvbiAocCwgaSkgew0KICAgICAgICBjb25zdCBib3JkZXIgPSBpIDwgTWF0aC5t"
    b64 = b64 & "aW4oZC5wb3N0cy5sZW5ndGgsIDUpIC0gMSA/ICcgYm9yZGVyLWIgYm9yZGVyLW91dGxpbmUtdmFyaWFudC8zMCcgOiAnJzsNCiAgICAgICAgcmV0dXJuICc8bGkgY2xhc3M9InB5LTMnICsgYm9yZGVyICsgJyBmbGV4IGdhcC0zIGhvdmVy"
    b64 = b64 & "OmJnLXN1cmZhY2UtY29udGFpbmVyLWxvdyB0cmFuc2l0aW9uLWNvbG9ycyBweC0yIHJvdW5kZWQtbWQiPicgKw0KICAgICAgICAgICc8c3BhbiBjbGFzcz0idGV4dC1wcmltYXJ5IGZvbnQtYm9sZCB0ZXh0LXNtIj4nICsgKGkgKyAxKSAr"
    b64 = b64 & "ICcuPC9zcGFuPicgKw0KICAgICAgICAgICc8YSBjbGFzcz0idGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdHJ1bmNhdGUiIGhyZWY9IicgKyBlc2MocC51cmwpICsgJyIgdGFyZ2V0PSJf"
    b64 = b64 & "YmxhbmsiIHJlbD0ibm9vcGVuZXIgbm9yZWZlcnJlciIgZHJhZ2dhYmxlPSJmYWxzZSIgdGl0bGU9IicgKyBlc2MocC50aXRsZSkgKyAnIj4nICsgZXNjKHAudGl0bGUpICsgJzwvYT48L2xpPic7DQogICAgICB9KS5qb2luKCcnKTsNCiAg"
    b64 = b64 & "ICAgIHJldHVybiAnPGFydGljbGUgY2xhc3M9ImJvYXJkLWNhcmQgYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0IHJvdW5kZWQteGwgc2hhZG93LVswcHhfNHB4XzEycHhfcmdiYSgwLDAsMCwwLjA1KV0gYm9yZGVyIGJvcmRlci1vdXRs"
    b64 = b64 & "aW5lLXZhcmlhbnQgZmxleCBmbGV4LWNvbCBoLWZ1bGwgb3ZlcmZsb3ctaGlkZGVuIiBkYXRhLW5hbWU9IicgKyBlc2MoZC5uYW1lKSArICciIGRhdGEtY2F0ZWdvcnk9IicgKyBjYXQgKyAnIj4nICsNCiAgICAgICAgJzxoZWFkZXIgY2xh"
    b64 = b64 & "c3M9InAtNCBmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIiIHN0eWxlPSJiYWNrZ3JvdW5kOicgKyBoZWFkICsgJyI+PGgzIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUgdGV4dC13aGl0ZSI+Wycg"
    b64 = b64 & "KyBlc2MoZC5uYW1lKSArICddPC9oMz48c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB0ZXh0LXdoaXRlIHRleHQtc20iPm1vcmVfdmVydDwvc3Bhbj48L2hlYWRlcj4nICsNCiAgICAgICAgJzxkaXYgY2xhc3M9InAt"
    b64 = b64 & "NCBmbGV4IGZsZXgtY29sIGdhcC0wIj48dWwgY2xhc3M9InNwYWNlLXktMCI+JyArIGl0ZW1zICsgJzwvdWw+PC9kaXY+PC9hcnRpY2xlPic7DQogICAgfSkuam9pbignJyk7DQogICAgQXJyYXkuZnJvbShib2FyZC5xdWVyeVNlbGVjdG9y"
    b64 = b64 & "QWxsKCcuYm9hcmQtY2FyZCcpKS5mb3JFYWNoKGJpbmREcmFnKTsNCiAgfQ0KICBmdW5jdGlvbiBzZXRGaWx0ZXIoZmlsdGVyKSB7DQogICAgY3VycmVudEZpbHRlciA9IGZpbHRlcjsNCiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxs"
    b64 = b64 & "KCcuZmlsdGVyLWJ0bicpLmZvckVhY2goYnRuID0+IHsNCiAgICAgIGNvbnN0IGFjdGl2ZSA9IGJ0bi5nZXRBdHRyaWJ1dGUoJ2RhdGEtZmlsdGVyJykgPT09IGZpbHRlcjsNCiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdmaWx0ZXIt"
    b64 = b64 & "YWN0aXZlJywgYWN0aXZlKTsNCiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCd0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCcsICFhY3RpdmUpOw0KICAgIH0pOw0KICAgIHJlbmRlcigpOw0KICB9DQogIGZ1bmN0aW9uIGxvYWQoKSB7DQog"
    b64 = b64 & "ICAgY29uc3QgcmF3ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25vdGljZS1kYXRhJykudGV4dENvbnRlbnQudHJpbSgpOw0KICAgIGxldCBkYXRhID0geyBjb2xsZWN0ZWRBdDogJycsIGRlcGFydG1lbnRzOiBbXSB9Ow0KICAgIHRy"
    b64 = b64 & "eSB7IGlmIChyYXcgJiYgcmF3LmNoYXJBdCgwKSA9PT0gJ3snKSBkYXRhID0gSlNPTi5wYXJzZShyYXcpOyB9IGNhdGNoIChlKSB7IGNvbnNvbGUuZXJyb3IoZSk7IH0NCiAgICBhbGxEZXB0cyA9IEFycmF5LmlzQXJyYXkoZGF0YS5kZXBh"
    b64 = b64 & "cnRtZW50cykgPyBkYXRhLmRlcGFydG1lbnRzLnNsaWNlKCkgOiBbXTsNCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWV0YUxpbmUnKS5pbm5lckhUTUwgPQ0KICAgICAgJ+yImOynkeyLnOqwgTogPHNwYW4gY2xhc3M9ImZvbnQt"
    b64 = b64 & "Ym9sZCI+JyArIGVzYyhkYXRhLmNvbGxlY3RlZEF0IHx8ICctJykgKyAnPC9zcGFuPiB8IOu2gOyymCA8c3BhbiBjbGFzcz0iZm9udC1ib2xkIj4nICsgYWxsRGVwdHMubGVuZ3RoICsgJ+qwnDwvc3Bhbj4gfCDsubTrk5wg65Oc656Y6re4"
    b64 = b64 & "66GcIOychOy5mCDsnbTrj5kg6rCA64qlJzsNCiAgICByZW5kZXIoKTsNCiAgfQ0KICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuZmlsdGVyLWJ0bicpLmZvckVhY2goYnRuID0+IHsNCiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcign"
    b64 = b64 & "Y2xpY2snLCBmdW5jdGlvbiAoKSB7IHNldEZpbHRlcihidG4uZ2V0QXR0cmlidXRlKCdkYXRhLWZpbHRlcicpKTsgfSk7DQogIH0pOw0KICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnRuUmVsb2FkJykuYWRkRXZlbnRMaXN0ZW5lcign"
    b64 = b64 & "Y2xpY2snLCBmdW5jdGlvbiAoKSB7IGxvY2F0aW9uLnJlbG9hZCgpOyB9KTsNCiAgbG9hZCgpOw0KfSkoKTsNCjwvc2NyaXB0Pg0KPC9ib2R5Pg0KPC9odG1sPg0K"
    EmbeddedHtmlTemplatePc = DecodeBase64Utf8(b64)
End Function

Private Function EmbeddedHtmlTemplateMobile() As String
    Dim b64 As String
    b64 = ""
    b64 = b64 & "PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAi"
    b64 = b64 & "IG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KCV67aA67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wgKOuqqOuwlOydvCk8L3RpdGxlPg0KPHNjcmlwdCBzcmM9Imh0dHBzOi8vY2RuLnRhaWx3aW5kY3NzLmNvbT9wbHVnaW5zPWZvcm1z"
    b64 = b64 & "LGNvbnRhaW5lci1xdWVyaWVzIj48L3NjcmlwdD4NCjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9UHVibGljK1NhbnM6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmFtcDtkaXNwbGF5PXN3"
    b64 = b64 & "YXAiIHJlbD0ic3R5bGVzaGVldCIvPg0KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1NYXRlcmlhbCtTeW1ib2xzK091dGxpbmVkOndnaHQsRklMTEAxMDAuLjcwMCwwLi4xJmFtcDtkaXNw"
    b64 = b64 & "bGF5PXN3YXAiIHJlbD0ic3R5bGVzaGVldCIvPg0KPHNjcmlwdCBpZD0idGFpbHdpbmQtY29uZmlnIj4NCiAgICAgIHRhaWx3aW5kLmNvbmZpZyA9IHsNCiAgICAgICAgZGFya01vZGU6ICJjbGFzcyIsDQogICAgICAgIHRoZW1lOiB7DQog"
    b64 = b64 & "ICAgICAgICAgZXh0ZW5kOiB7DQogICAgICAgICAgICAiY29sb3JzIjogew0KICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktZml4ZWQtdmFyaWFudCI6ICIjN2QyZDAwIiwNCiAgICAgICAgICAgICAgICAgICAgIm91dGxpbmUt"
    b64 = b64 & "dmFyaWFudCI6ICIjYzNjNmQ3IiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXRlcnRpYXJ5IjogIiNmZmZmZmYiLA0KICAgICAgICAgICAgICAgICAgICAiYmFja2dyb3VuZCI6ICIjZmFmOGZmIiwNCiAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "Im9uLXByaW1hcnktY29udGFpbmVyIjogIiNlZWVmZmYiLA0KICAgICAgICAgICAgICAgICAgICAib24tZXJyb3IiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICJzZWNvbmRhcnkiOiAiIzlkNDMwMCIsDQogICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICJ0ZXJ0aWFyeS1maXhlZCI6ICIjZmZkYmNkIiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyIjogIiNlZGVkZjkiLA0KICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeSI6ICIjZmZmZmZm"
    b64 = b64 & "IiwNCiAgICAgICAgICAgICAgICAgICAgImludmVyc2Utc3VyZmFjZSI6ICIjMmUzMDM5IiwNCiAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5IjogIiM5NDM3MDAiLA0KICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWlu"
    b64 = b64 & "ZXItaGlnaCI6ICIjZTdlN2YzIiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtZGltIjogIiNkOWQ5ZTUiLA0KICAgICAgICAgICAgICAgICAgICAic3VyZmFjZSI6ICIjZmFmOGZmIiwNCiAgICAgICAgICAgICAgICAgICAgInBy"
    b64 = b64 & "aW1hcnktZml4ZWQtZGltIjogIiNiNGM1ZmYiLA0KICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1maXhlZCI6ICIjZGJlMWZmIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeS1maXhlZCI6ICIjMzQxMTAwIiwNCiAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWZpeGVkLWRpbSI6ICIjZmZiNTk2IiwNCiAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1maXhlZC1kaW0iOiAiI2ZmYjY5MCIsDQogICAgICAgICAgICAgICAgICAgICJwcmltYXJ5"
    b64 = b64 & "LWNvbnRhaW5lciI6ICIjMjU2M2ViIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeSI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWNvbnRhaW5lciI6ICIjYmM0ODAwIiwNCiAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgIm9uLXNlY29uZGFyeS1jb250YWluZXIiOiAiIzVjMjQwMCIsDQogICAgICAgICAgICAgICAgICAgICJzZWNvbmRhcnktY29udGFpbmVyIjogIiNmZDc2MWEiLA0KICAgICAgICAgICAgICAgICAgICAib24tc2Vjb25k"
    b64 = b64 & "YXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzc4MzIwMCIsDQogICAgICAgICAgICAgICAgICAgICJlcnJvci1jb250YWluZXIiOiAiI2ZmZGFkNiIsDQogICAgICAgICAgICAgICAgICAgICJvbi10ZXJ0aWFyeS1jb250YWluZXIiOiAiI2ZmZWRl"
    b64 = b64 & "NiIsDQogICAgICAgICAgICAgICAgICAgICJvbi1wcmltYXJ5LWZpeGVkIjogIiMwMDE3NGIiLA0KICAgICAgICAgICAgICAgICAgICAib24tc3VyZmFjZS12YXJpYW50IjogIiM0MzQ2NTUiLA0KICAgICAgICAgICAgICAgICAgICAic3Vy"
    b64 = b64 & "ZmFjZS1jb250YWluZXItaGlnaGVzdCI6ICIjZTFlMmVkIiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtdmFyaWFudCI6ICIjZTFlMmVkIiwNCiAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1maXhlZCI6ICIjZmZkYmNh"
    b64 = b64 & "IiwNCiAgICAgICAgICAgICAgICAgICAgIm91dGxpbmUiOiAiIzczNzY4NiIsDQogICAgICAgICAgICAgICAgICAgICJvbi1zdXJmYWNlIjogIiMxOTFiMjMiLA0KICAgICAgICAgICAgICAgICAgICAiaW52ZXJzZS1vbi1zdXJmYWNlIjog"
    b64 = b64 & "IiNmMGYwZmIiLA0KICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1maXhlZC12YXJpYW50IjogIiMwMDNlYTgiLA0KICAgICAgICAgICAgICAgICAgICAiZXJyb3IiOiAiI2JhMWExYSIsDQogICAgICAgICAgICAgICAgICAgICJz"
    b64 = b64 & "dXJmYWNlLWNvbnRhaW5lci1sb3dlc3QiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLWJyaWdodCI6ICIjZmFmOGZmIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXRlcnRpYXJ5LWZpeGVkIjogIiMzNjBm"
    b64 = b64 & "MDAiLA0KICAgICAgICAgICAgICAgICAgICAicHJpbWFyeSI6ICIjMDA0YWM2IiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLWJhY2tncm91bmQiOiAiIzE5MWIyMyIsDQogICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXByaW1hcnki"
    b64 = b64 & "OiAiI2I0YzVmZiIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLWNvbnRhaW5lci1sb3ciOiAiI2YzZjNmZSIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLXRpbnQiOiAiIzAwNTNkYiIsDQogICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICJvbi1lcnJvci1jb250YWluZXIiOiAiIzkzMDAwYSINCiAgICAgICAgICAgIH0sDQogICAgICAgICAgICAiYm9yZGVyUmFkaXVzIjogew0KICAgICAgICAgICAgICAgICAgICAiREVGQVVMVCI6ICIwLjI1cmVtIiwNCiAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgImxnIjogIjAuNXJlbSIsDQogICAgICAgICAgICAgICAgICAgICJ4bCI6ICIwLjc1cmVtIiwNCiAgICAgICAgICAgICAgICAgICAgImZ1bGwiOiAiOTk5OXB4Ig0KICAgICAgICAgICAgfSwNCiAgICAgICAgICAgICJz"
    b64 = b64 & "cGFjaW5nIjogew0KICAgICAgICAgICAgICAgICAgICAiZ3JpZC1ndXR0ZXIiOiAiMS4yNXJlbSIsDQogICAgICAgICAgICAgICAgICAgICJjb250YWluZXItcGFkZGluZyI6ICIycmVtIiwNCiAgICAgICAgICAgICAgICAgICAgInN0YWNr"
    b64 = b64 & "LWdhcCI6ICIwLjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAiY2FyZC1wYWRkaW5nIjogIjFyZW0iDQogICAgICAgICAgICB9LA0KICAgICAgICAgICAgImZvbnRGYW1pbHkiOiB7DQogICAgICAgICAgICAgICAgICAgICJsaXN0LWl0"
    b64 = b64 & "ZW0iOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICJtZXRhLWRhdGEiOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICJoZWFkZXItdGl0bGUiOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICJidXR0b24tdGV4dCI6IFsiUHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgImJvYXJkLXRpdGxlIjogWyJQdWJsaWMgU2FucyJdDQogICAgICAgICAgICB9LA0KICAgICAgICAgICAgImZvbnRTaXpl"
    b64 = b64 & "Ijogew0KICAgICAgICAgICAgICAgICAgICAibGlzdC1pdGVtIjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdodCI6ICI1MDAifV0sDQogICAgICAgICAgICAgICAgICAgICJtZXRhLWRhdGEiOiBbIjEycHgi"
    b64 = b64 & "LCB7ImxpbmVIZWlnaHQiOiAiMTZweCIsICJmb250V2VpZ2h0IjogIjQwMCJ9XSwNCiAgICAgICAgICAgICAgICAgICAgImhlYWRlci10aXRsZSI6IFsiMjRweCIsIHsibGluZUhlaWdodCI6ICIzMnB4IiwgImxldHRlclNwYWNpbmciOiAi"
    b64 = b64 & "LTAuMDJlbSIsICJmb250V2VpZ2h0IjogIjcwMCJ9XSwNCiAgICAgICAgICAgICAgICAgICAgImJ1dHRvbi10ZXh0IjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdodCI6ICI2MDAifV0sDQogICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICJib2FyZC10aXRsZSI6IFsiMTZweCIsIHsibGluZUhlaWdodCI6ICIyNHB4IiwgImZvbnRXZWlnaHQiOiAiNzAwIn1dDQogICAgICAgICAgICB9DQogICAgICAgICAgfSwNCiAgICAgICAgfSwNCiAgICAgIH0NCiAgICA8"
    b64 = b64 & "L3NjcmlwdD4NCjxzdHlsZT4NCiAgICAgICAgYm9keSB7DQogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiAjZmFmOGZmOw0KICAgICAgICAgICAgLXdlYmtpdC10YXAtaGlnaGxpZ2h0LWNvbG9yOiB0cmFuc3BhcmVudDsNCiAgICAg"
    b64 = b64 & "ICAgfQ0KICAgICAgICAubWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB7DQogICAgICAgICAgICBmb250LXZhcmlhdGlvbi1zZXR0aW5nczogJ0ZJTEwnIDAsICd3Z2h0JyA0MDAsICdHUkFEJyAwLCAnb3BzeicgMjQ7DQogICAgICAgIH0N"
    b64 = b64 & "CiAgICAgICAgLmhpZGUtc2Nyb2xsYmFyOjotd2Via2l0LXNjcm9sbGJhciB7IGRpc3BsYXk6IG5vbmU7IH0NCiAgICAgICAgLmhpZGUtc2Nyb2xsYmFyIHsgLW1zLW92ZXJmbG93LXN0eWxlOiBub25lOyBzY3JvbGxiYXItd2lkdGg6IG5v"
    b64 = b64 & "bmU7IH0NCiAgICAgICAgLm5vdGljZS1jYXJkIHsgYm94LXNoYWRvdzogMHB4IDRweCAxMnB4IHJnYmEoMCwwLDAsMC4wNSk7IH0NCiAgICAgICAgLnRhYi1hY3RpdmUgew0KICAgICAgICAgICAgYmFja2dyb3VuZC1jb2xvcjogI2ZkNzYx"
    b64 = b64 & "YSAhaW1wb3J0YW50Ow0KICAgICAgICAgICAgY29sb3I6ICM1YzI0MDAgIWltcG9ydGFudDsNCiAgICAgICAgfQ0KPC9zdHlsZT4NCjwvaGVhZD4NCjxib2R5IGNsYXNzPSJmbGV4IGZsZXgtY29sIG1pbi1oLXNjcmVlbiI+DQo8aGVhZGVy"
    b64 = b64 & "IGNsYXNzPSJiZy1wcmltYXJ5IHRleHQtb24tcHJpbWFyeSBmaXhlZCB0b3AtMCBsZWZ0LTAgcmlnaHQtMCB6LTUwIHJvdW5kZWQtYi14bCBzaGFkb3ctbWQgZmxleCBmbGV4LWNvbCBpdGVtcy1jZW50ZXIgcHgtNCBweS00IHctZnVsbCI+"
    b64 = b64 & "DQo8ZGl2IGNsYXNzPSJmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgdy1mdWxsIj4NCjxoMSBjbGFzcz0iZm9udC1oZWFkZXItdGl0bGUgdGV4dC1oZWFkZXItdGl0bGUgdGV4dC1vbi1wcmltYXJ5Ij7soJXrtoDrtoDsspgg"
    b64 = b64 & "6rO17KeA7IKs7ZWtPC9oMT4NCjxidXR0b24gaWQ9ImJ0blJlbG9hZCIgdHlwZT0iYnV0dG9uIiBjbGFzcz0iZm9udC1idXR0b24tdGV4dCB0ZXh0LWJ1dHRvbi10ZXh0IGZsZXggaXRlbXMtY2VudGVyIGdhcC0xIGhvdmVyOmJnLXByaW1h"
    b64 = b64 & "cnktY29udGFpbmVyLzIwIHRyYW5zaXRpb24tY29sb3JzIHAtMiByb3VuZGVkLWxnIGFjdGl2ZTpzY2FsZS05NSBkdXJhdGlvbi0xNTAiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnJlZnJlc2g8L3NwYW4+"
    b64 = b64 & "DQo8c3BhbiBjbGFzcz0iaGlkZGVuIHNtOmlubGluZSI+7ZmU66m0IOyDiOuhnOqzoOy5qDwvc3Bhbj4NCjwvYnV0dG9uPg0KPC9kaXY+DQo8ZGl2IGNsYXNzPSJ3LWZ1bGwgbXQtNCBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWJldHdl"
    b64 = b64 & "ZW4gZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgb3BhY2l0eS05MCBib3JkZXItdCBib3JkZXItb24tcHJpbWFyeS8xMCBwdC0zIj4NCjxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIj4NCjxzcGFuIGNsYXNzPSJt"
    b64 = b64 & "YXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE0cHhdIj51cGRhdGU8L3NwYW4+DQo8c3BhbiBpZD0ibWV0YVRpbWUiPuuniOyngOuniSDsl4XrjbDsnbTtirg6IC08L3NwYW4+DQo8L2Rpdj4NCjxkaXYgaWQ9Im1ldGFDb3VudCI+"
    b64 = b64 & "67aA7LKYIOy0nTogMOqwnDwvZGl2Pg0KPC9kaXY+DQo8L2hlYWRlcj4NCg0KPG5hdiBjbGFzcz0iZml4ZWQgdG9wLVsxMTBweF0gbGVmdC0wIHJpZ2h0LTAgYmctc3VyZmFjZS1jb250YWluZXIgei00MCBib3JkZXItYiBib3JkZXItb3V0"
    b64 = b64 & "bGluZS12YXJpYW50IHNoYWRvdy1zbSI+DQo8ZGl2IGNsYXNzPSJmbGV4IG92ZXJmbG93LXgtYXV0byBoaWRlLXNjcm9sbGJhciBweC00IHB5LTMgZ2FwLTIiIGlkPSJmaWx0ZXJOYXYiPg0KPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEt"
    b64 = b64 & "ZmlsdGVyPSJhbGwiIGNsYXNzPSJ0YWItYnRuIHRhYi1hY3RpdmUgZmxleC1zaHJpbmstMCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBweC00IHB5LTIgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cmFuc2l0"
    b64 = b64 & "aW9uLWFsbCBkdXJhdGlvbi0yMDAgZWFzZS1pbi1vdXQiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPmRhc2hib2FyZDwvc3Bhbj7soITssrQg67aA7LKYDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0"
    b64 = b64 & "dG9uIiBkYXRhLWZpbHRlcj0iZWNvbm9teSIgY2xhc3M9InRhYi1idG4gZmxleC1zaHJpbmstMCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBweC00IHB5LTIgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgaG92ZXI6Ymctc3VyZmFjZS12YXJp"
    b64 = b64 & "YW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdHJhbnNpdGlvbi1hbGwgZHVyYXRpb24tMjAwIGVhc2UtaW4tb3V0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5idXNpbmVz"
    b64 = b64 & "c19jZW50ZXI8L3NwYW4+7IKw7JeFL+qyveygnA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9InNvY2lldHkiIGNsYXNzPSJ0YWItYnRuIGZsZXgtc2hyaW5rLTAgZmxleCBpdGVtcy1jZW50ZXIgZ2Fw"
    b64 = b64 & "LTIgcHgtNCBweS0yIHRleHQtb24tc3VyZmFjZS12YXJpYW50IGhvdmVyOmJnLXN1cmZhY2UtdmFyaWFudCByb3VuZGVkLWxnIGZvbnQtbGlzdC1pdGVtIHRleHQtbGlzdC1pdGVtIHRyYW5zaXRpb24tYWxsIGR1cmF0aW9uLTIwMCBlYXNl"
    b64 = b64 & "LWluLW91dCI+DQo8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCI+dGhlYXRlcl9jb21lZHk8L3NwYW4+7IKs7ZqML+usuO2ZlA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImFk"
    b64 = b64 & "bWluIiBjbGFzcz0idGFiLWJ0biBmbGV4LXNocmluay0wIGZsZXggaXRlbXMtY2VudGVyIGdhcC0yIHB4LTQgcHktMiB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQgcm91bmRlZC1sZyBmb250LWxp"
    b64 = b64 & "c3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cmFuc2l0aW9uLWFsbCBkdXJhdGlvbi0yMDAgZWFzZS1pbi1vdXQiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPmFjY291bnRfYmFsYW5jZTwvc3Bhbj7tlonsoJUv"
    b64 = b64 & "7JWI7KCEDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0iZGlwbG9tYWN5IiBjbGFzcz0idGFiLWJ0biBmbGV4LXNocmluay0wIGZsZXggaXRlbXMtY2VudGVyIGdhcC0yIHB4LTQgcHktMiB0ZXh0LW9u"
    b64 = b64 & "LXN1cmZhY2UtdmFyaWFudCBob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cmFuc2l0aW9uLWFsbCBkdXJhdGlvbi0yMDAgZWFzZS1pbi1vdXQiPg0KPHNwYW4gY2xh"
    b64 = b64 & "c3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnB1YmxpYzwvc3Bhbj7smbjqtZAv7JWI67O0DQo8L2J1dHRvbj4NCjwvZGl2Pg0KPC9uYXY+DQoNCjxtYWluIGNsYXNzPSJtdC1bMTgwcHhdIG1iLTIwIHB4LTQgZmxleCBmbGV4LWNv"
    b64 = b64 & "bCBnYXAtNSBtYXgtdy1sZyBteC1hdXRvIHctZnVsbCIgaWQ9ImJvYXJkIj48L21haW4+DQoNCjxidXR0b24gaWQ9ImZhYlN5bmMiIHR5cGU9ImJ1dHRvbiIgY2xhc3M9ImZpeGVkIGJvdHRvbS02IHJpZ2h0LTYgdy0xNCBoLTE0IGJnLXBy"
    b64 = b64 & "aW1hcnktY29udGFpbmVyIHRleHQtb24tcHJpbWFyeS1jb250YWluZXIgcm91bmRlZC1mdWxsIHNoYWRvdy1sZyBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciBob3ZlcjpzY2FsZS0xMDUgYWN0aXZlOnNjYWxlLTk1IHRyYW5z"
    b64 = b64 & "aXRpb24tYWxsIHotNTAiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1bMzJweF0iPnN5bmM8L3NwYW4+DQo8L2J1dHRvbj4NCg0KPGZvb3RlciBjbGFzcz0iYmctc3VyZmFjZS1kaW0gdy1mdWxsIHB5"
    b64 = b64 & "LTYgcHgtNCBmbGV4IGZsZXgtY29sIGl0ZW1zLWNlbnRlciBnYXAtNCB0ZXh0LWNlbnRlciBib3JkZXItdCBib3JkZXItb3V0bGluZS12YXJpYW50IG10LWF1dG8iPg0KPGRpdiBjbGFzcz0iZm9udC1ib2xkIHRleHQtb24tc3VyZmFjZS12"
    b64 = b64 & "YXJpYW50IG1iLTEiPuuMgO2VnOuvvOq1rSDsoJXrtoA8L2Rpdj4NCjxkaXYgY2xhc3M9ImZsZXggZmxleC13cmFwIGp1c3RpZnktY2VudGVyIGdhcC00IGZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIHRleHQtb24tc3VyZmFjZS12"
    b64 = b64 & "YXJpYW50IG9wYWNpdHktODAiPg0KPGEgY2xhc3M9ImhvdmVyOnRleHQtcHJpbWFyeSBuby11bmRlcmxpbmUgdHJhbnNpdGlvbi1jb2xvcnMiIGhyZWY9Imdvdl9ub3RpY2VfYm9hcmQuaHRtbCI+UEMg7Y6Y7J207KeAPC9hPg0KPGEgY2xh"
    b64 = b64 & "c3M9ImhvdmVyOnRleHQtcHJpbWFyeSBuby11bmRlcmxpbmUgdHJhbnNpdGlvbi1jb2xvcnMiIGhyZWY9Imdvdl9ub3RpY2VfYm9hcmRfbW9iaWxlLmh0bWwiPuuqqOuwlOydvDwvYT4NCjwvZGl2Pg0KPHAgY2xhc3M9ImZvbnQtbWV0YS1k"
    b64 = b64 & "YXRhIHRleHQtbWV0YS1kYXRhIHRleHQtb24tc3VyZmFjZS12YXJpYW50IG9wYWNpdHktNjAiPsKpIOuMgO2VnOuvvOq1rSDsoJXrtoAg67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wg7ISc67mE7IqkPC9wPg0KPC9mb290ZXI+DQoNCjxz"
    b64 = b64 & "Y3JpcHQgaWQ9Im5vdGljZS1kYXRhIiB0eXBlPSJhcHBsaWNhdGlvbi9qc29uIj4NCiUlTk9USUNFX0pTT04lJQ0KPC9zY3JpcHQ+DQo8c2NyaXB0Pg0KKGZ1bmN0aW9uICgpIHsNCiAgY29uc3QgU1RPUkFHRV9LRVkgPSAnZ292Tm90aWNl"
    b64 = b64 & "Q2FyZE9yZGVyLnYxJzsNCiAgY29uc3QgUEFMRVRURSA9IFsnIzAwNGFjNicsJyM5ZDQzMDAnLCcjOTQzNzAwJywnIzczNzY4NicsJyMyNTYzZWInLCcjMmU3ZDMyJywnIzAwOTY4OCcsJyNmZDc2MWEnLCcjNmM3ZmQ4JywnIzAwOTdhNydd"
    b64 = b64 & "Ow0KICBjb25zdCBJQ09OUyA9IHsNCiAgICAn67O06rG067O17KeA67aAJzonbWVkaWNhbF9zZXJ2aWNlcycsJ+2WieygleyViOyghOu2gCc6J3NlY3VyaXR5Jywn7Jm46rWQ67aAJzoncHVibGljJywn6rWt67Cp67aAJzonc2VjdXJpdHkn"
    b64 = b64 & "LA0KICAgICfqtZDsnKHrtoAnOidzY2hvb2wnLCfqs6Dsmqnrhbjrj5nrtoAnOid3b3JrJywn66y47ZmU7LK07Jyh6rSA6rSR67aAJzondGhlYXRlcl9jb21lZHknLCdkZWZhdWx0Jzonc3RhcnMnDQogIH07DQogIGNvbnN0IENBVEVHT1JZ"
    b64 = b64 & "ID0gew0KICAgICfsnqzsoJXqsr3soJzrtoAnOidlY29ub215Jywn6rO87ZWZ6riw7Iig7KCV67O07Ya17Iug67aAJzonZWNvbm9teScsJ+yCsOyXhe2GteyDgeu2gCc6J2Vjb25vbXknLCfspJHshozrsqTsspjquLDsl4XrtoAnOidlY29u"
    b64 = b64 & "b215JywNCiAgICAn6rWt7Yag6rWQ7Ya167aAJzonZWNvbm9teScsJ+uGjeumvOy2leyCsOyLne2SiOu2gCc6J2Vjb25vbXknLCftlbTslpHsiJjsgrDrtoAnOidlY29ub215Jywn6riw7ZuE7JeQ64SI7KeA7ZmY6rK967aAJzonZWNvbm9t"
    b64 = b64 & "eScsJ+q4sO2ajeyYiOyCsOyymCc6J2Vjb25vbXknLA0KICAgICfqtZDsnKHrtoAnOidzb2NpZXR5Jywn67O06rG067O17KeA67aAJzonc29jaWV0eScsJ+qzoOyaqeuFuOuPmeu2gCc6J3NvY2lldHknLCfrrLjtmZTssrTsnKHqtIDqtJHr"
    b64 = b64 & "toAnOidzb2NpZXR5Jywn7ISx7Y+J65Ox6rCA7KGx67aAJzonc29jaWV0eScsJ+yLne2SiOydmOyVve2SiOyViOyghOyymCc6J3NvY2lldHknLA0KICAgICftlonsoJXslYjsoITrtoAnOidhZG1pbicsJ+uyleustOu2gCc6J2FkbWluJywn"
    b64 = b64 & "7J247IKs7ZiB7Iug7LKYJzonYWRtaW4nLCfrspXsoJzsspgnOidhZG1pbicsJ+q1reqwgOuztO2biOu2gCc6J2FkbWluJywNCiAgICAn7Jm46rWQ67aAJzonZGlwbG9tYWN5Jywn7Ya17J2867aAJzonZGlwbG9tYWN5Jywn6rWt67Cp67aA"
    b64 = b64 & "JzonZGlwbG9tYWN5Jw0KICB9Ow0KICBsZXQgY3VycmVudEZpbHRlciA9ICdhbGwnOw0KICBsZXQgYWxsRGVwdHMgPSBbXTsNCg0KICBmdW5jdGlvbiBjb2xvckZvcihuYW1lKSB7DQogICAgaWYgKG5hbWUgPT09ICfsgrDsl4XthrXsg4Hr"
    b64 = b64 & "toAnKSByZXR1cm4gJyNmZDc2MWEnOw0KICAgIGxldCBpZHggPSAwOw0KICAgIGZvciAobGV0IGkgPSAwOyBpIDwgbmFtZS5sZW5ndGg7IGkrKykgaWR4ID0gKGlkeCArIG5hbWUuY2hhckNvZGVBdChpKSkgJSBQQUxFVFRFLmxlbmd0aDsN"
    b64 = b64 & "CiAgICByZXR1cm4gUEFMRVRURVtpZHhdOw0KICB9DQogIGZ1bmN0aW9uIGNhdGVnb3J5Rm9yKG5hbWUpIHsgcmV0dXJuIENBVEVHT1JZW25hbWVdIHx8ICdhZG1pbic7IH0NCiAgZnVuY3Rpb24gaWNvbkZvcihuYW1lKSB7IHJldHVybiBJ"
    b64 = b64 & "Q09OU1tuYW1lXSB8fCBJQ09OUy5kZWZhdWx0OyB9DQogIGZ1bmN0aW9uIGVzYyhzKSB7DQogICAgcmV0dXJuIFN0cmluZyhzID8/ICcnKS5yZXBsYWNlKC8mL2csJyZhbXA7JykucmVwbGFjZSgvPC9nLCcmbHQ7JykucmVwbGFjZSgvPi9n"
    b64 = b64 & "LCcmZ3Q7JykucmVwbGFjZSgvIi9nLCcmcXVvdDsnKTsNCiAgfQ0KICBmdW5jdGlvbiBhcHBseVNhdmVkT3JkZXIoZGVwdHMpIHsNCiAgICB0cnkgew0KICAgICAgY29uc3QgcmF3ID0gbG9jYWxTdG9yYWdlLmdldEl0ZW0oU1RPUkFHRV9L"
    b64 = b64 & "RVkpOw0KICAgICAgaWYgKCFyYXcpIHJldHVybiBkZXB0czsNCiAgICAgIGNvbnN0IG9yZGVyID0gSlNPTi5wYXJzZShyYXcpOw0KICAgICAgaWYgKCFBcnJheS5pc0FycmF5KG9yZGVyKSB8fCAhb3JkZXIubGVuZ3RoKSByZXR1cm4gZGVw"
    b64 = b64 & "dHM7DQogICAgICBjb25zdCBtYXAgPSB7fTsNCiAgICAgIGRlcHRzLmZvckVhY2goZCA9PiB7IG1hcFtkLm5hbWVdID0gZDsgfSk7DQogICAgICBjb25zdCBzb3J0ZWQgPSBbXTsNCiAgICAgIG9yZGVyLmZvckVhY2gobmFtZSA9PiB7IGlm"
    b64 = b64 & "IChtYXBbbmFtZV0pIHsgc29ydGVkLnB1c2gobWFwW25hbWVdKTsgZGVsZXRlIG1hcFtuYW1lXTsgfSB9KTsNCiAgICAgIE9iamVjdC5rZXlzKG1hcCkuZm9yRWFjaChrID0+IHNvcnRlZC5wdXNoKG1hcFtrXSkpOw0KICAgICAgcmV0dXJu"
    b64 = b64 & "IHNvcnRlZDsNCiAgICB9IGNhdGNoIChlKSB7IHJldHVybiBkZXB0czsgfQ0KICB9DQogIGZ1bmN0aW9uIHBhZE51bShuKSB7IHJldHVybiAobiA8IDEwID8gJzAnIDogJycpICsgbjsgfQ0KICBmdW5jdGlvbiByZW5kZXIoKSB7DQogICAg"
    b64 = b64 & "Y29uc3QgYm9hcmQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYm9hcmQnKTsNCiAgICBsZXQgZGVwdHMgPSBhbGxEZXB0cy5zbGljZSgpOw0KICAgIGlmIChjdXJyZW50RmlsdGVyICE9PSAnYWxsJykgZGVwdHMgPSBkZXB0cy5maWx0"
    b64 = b64 & "ZXIoZCA9PiBjYXRlZ29yeUZvcihkLm5hbWUpID09PSBjdXJyZW50RmlsdGVyKTsNCiAgICBkZXB0cyA9IGFwcGx5U2F2ZWRPcmRlcihkZXB0cyk7DQogICAgaWYgKCFkZXB0cy5sZW5ndGgpIHsNCiAgICAgIGJvYXJkLmlubmVySFRNTCA9"
    b64 = b64 & "ICc8ZGl2IGNsYXNzPSJ0ZXh0LWNlbnRlciBweS0xNiB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBiZy1zdXJmYWNlLWNvbnRhaW5lci1sb3dlc3Qgcm91bmRlZC14bCBib3JkZXIgYm9yZGVyLWRhc2hlZCBib3JkZXItb3V0bGluZS12YXJp"
    b64 = b64 & "YW50Ij7tkZzsi5ztlaAg6rO17KeAIOuNsOydtO2EsOqwgCDsl4bsirXri4jri6QuPC9kaXY+JzsNCiAgICAgIHJldHVybjsNCiAgICB9DQogICAgYm9hcmQuaW5uZXJIVE1MID0gZGVwdHMubWFwKGZ1bmN0aW9uIChkKSB7DQogICAgICBj"
    b64 = b64 & "b25zdCBoZWFkID0gY29sb3JGb3IoZC5uYW1lKTsNCiAgICAgIGNvbnN0IGljb24gPSBpY29uRm9yKGQubmFtZSk7DQogICAgICBpZiAoZC5vayA9PT0gZmFsc2UgfHwgIWQucG9zdHMgfHwgIWQucG9zdHMubGVuZ3RoKSB7DQogICAgICAg"
    b64 = b64 & "IHJldHVybiAnPGFydGljbGUgY2xhc3M9Im5vdGljZS1jYXJkIGJnLXN1cmZhY2UtY29udGFpbmVyLWxvd2VzdCByb3VuZGVkLXhsIG92ZXJmbG93LWhpZGRlbiBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudC8zMCI+JyArDQogICAg"
    b64 = b64 & "ICAgICAgJzxoZWFkZXIgY2xhc3M9InB4LTQgcHktMyBmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIiIHN0eWxlPSJiYWNrZ3JvdW5kOicgKyBoZWFkICsgJyI+JyArDQogICAgICAgICAgJzxoMiBjbGFzcz0iZm9udC1ib2Fy"
    b64 = b64 & "ZC10aXRsZSB0ZXh0LWJvYXJkLXRpdGxlIHRleHQtd2hpdGUiPicgKyBlc2MoZC5uYW1lKSArICc8L2gyPicgKw0KICAgICAgICAgICc8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB0ZXh0LXdoaXRlIj4nICsgaWNv"
    b64 = b64 & "biArICc8L3NwYW4+PC9oZWFkZXI+JyArDQogICAgICAgICAgJzxkaXYgY2xhc3M9InAtY2FyZC1wYWRkaW5nIHRleHQtZXJyb3IgZm9udC1ib2xkIHRleHQtc20iPuyEpOyglSDtjpjsnbTsp4Drpbwg7IiY7KCV7ZWY7IS47JqULjwvZGl2"
    b64 = b64 & "PjwvYXJ0aWNsZT4nOw0KICAgICAgfQ0KICAgICAgY29uc3QgaXRlbXMgPSBkLnBvc3RzLnNsaWNlKDAsIDUpLm1hcChmdW5jdGlvbiAocCwgaSkgew0KICAgICAgICBjb25zdCBsYXN0ID0gaSA9PT0gTWF0aC5taW4oZC5wb3N0cy5sZW5n"
    b64 = b64 & "dGgsIDUpIC0gMTsNCiAgICAgICAgY29uc3QgYm9yZGVyID0gbGFzdCA/ICcnIDogJyBib3JkZXItYiBib3JkZXItb3V0bGluZS12YXJpYW50LzIwIHBiLTInOw0KICAgICAgICByZXR1cm4gJzxsaSBjbGFzcz0iZmxleCBpdGVtcy1zdGFy"
    b64 = b64 & "dCBnYXAtMyI+JyArDQogICAgICAgICAgJzxzcGFuIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSBwdC0xIiBzdHlsZT0iY29sb3I6JyArIGhlYWQgKyAnIj4nICsgcGFkTnVtKGkgKyAxKSArICc8L3NwYW4+JyArDQog"
    b64 = b64 & "ICAgICAgICAgJzxkaXYgY2xhc3M9ImZsZXggZmxleC1jb2wgdy1mdWxsJyArIGJvcmRlciArICciPicgKw0KICAgICAgICAgICc8YSBjbGFzcz0iZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlIGxpbmUt"
    b64 = b64 & "Y2xhbXAtMiBuby11bmRlcmxpbmUiIGhyZWY9IicgKyBlc2MocC51cmwpICsgJyIgdGFyZ2V0PSJfYmxhbmsiIHJlbD0ibm9vcGVuZXIgbm9yZWZlcnJlciI+JyArIGVzYyhwLnRpdGxlKSArICc8L2E+JyArDQogICAgICAgICAgJzwvZGl2"
    b64 = b64 & "PjwvbGk+JzsNCiAgICAgIH0pLmpvaW4oJycpOw0KICAgICAgcmV0dXJuICc8YXJ0aWNsZSBjbGFzcz0ibm90aWNlLWNhcmQgYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0IHJvdW5kZWQteGwgb3ZlcmZsb3ctaGlkZGVuIGJvcmRlciBi"
    b64 = b64 & "b3JkZXItb3V0bGluZS12YXJpYW50LzMwIHRyYW5zaXRpb24tdHJhbnNmb3JtIGFjdGl2ZTpzY2FsZS1bMC45OF0iIGRhdGEtbmFtZT0iJyArIGVzYyhkLm5hbWUpICsgJyI+JyArDQogICAgICAgICc8aGVhZGVyIGNsYXNzPSJweC00IHB5"
    b64 = b64 & "LTMgZmxleCBqdXN0aWZ5LWJldHdlZW4gaXRlbXMtY2VudGVyIiBzdHlsZT0iYmFja2dyb3VuZDonICsgaGVhZCArICciPicgKw0KICAgICAgICAnPGgyIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUgdGV4dC13"
    b64 = b64 & "aGl0ZSI+JyArIGVzYyhkLm5hbWUpICsgJzwvaDI+JyArDQogICAgICAgICc8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB0ZXh0LXdoaXRlIj4nICsgaWNvbiArICc8L3NwYW4+PC9oZWFkZXI+JyArDQogICAgICAg"
    b64 = b64 & "ICc8ZGl2IGNsYXNzPSJwLWNhcmQtcGFkZGluZyI+PHVsIGNsYXNzPSJzcGFjZS15LTQiPicgKyBpdGVtcyArICc8L3VsPjwvZGl2PjwvYXJ0aWNsZT4nOw0KICAgIH0pLmpvaW4oJycpOw0KICB9DQogIGZ1bmN0aW9uIHNldEZpbHRlcihm"
    b64 = b64 & "aWx0ZXIsIGJ0bikgew0KICAgIGN1cnJlbnRGaWx0ZXIgPSBmaWx0ZXI7DQogICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKGIgPT4gew0KICAgICAgYi5jbGFzc0xpc3QucmVtb3ZlKCd0YWItYWN0"
    b64 = b64 & "aXZlJywgJ2JnLXNlY29uZGFyeS1jb250YWluZXInLCAndGV4dC1vbi1zZWNvbmRhcnktY29udGFpbmVyJyk7DQogICAgICBiLmNsYXNzTGlzdC5hZGQoJ3RleHQtb24tc3VyZmFjZS12YXJpYW50JywgJ2hvdmVyOmJnLXN1cmZhY2UtdmFy"
    b64 = b64 & "aWFudCcpOw0KICAgIH0pOw0KICAgIGlmIChidG4pIHsNCiAgICAgIGJ0bi5jbGFzc0xpc3QuYWRkKCd0YWItYWN0aXZlJyk7DQogICAgICBidG4uY2xhc3NMaXN0LnJlbW92ZSgndGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQnLCAnaG92ZXI6"
    b64 = b64 & "Ymctc3VyZmFjZS12YXJpYW50Jyk7DQogICAgICBidG4uc2Nyb2xsSW50b1ZpZXcoeyBiZWhhdmlvcjogJ3Ntb290aCcsIGJsb2NrOiAnbmVhcmVzdCcsIGlubGluZTogJ2NlbnRlcicgfSk7DQogICAgfQ0KICAgIHJlbmRlcigpOw0KICB9"
    b64 = b64 & "DQogIGZ1bmN0aW9uIGxvYWQoKSB7DQogICAgY29uc3QgcmF3ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25vdGljZS1kYXRhJykudGV4dENvbnRlbnQudHJpbSgpOw0KICAgIGxldCBkYXRhID0geyBjb2xsZWN0ZWRBdDogJycsIGRl"
    b64 = b64 & "cGFydG1lbnRzOiBbXSB9Ow0KICAgIHRyeSB7IGlmIChyYXcgJiYgcmF3LmNoYXJBdCgwKSA9PT0gJ3snKSBkYXRhID0gSlNPTi5wYXJzZShyYXcpOyB9IGNhdGNoIChlKSB7IGNvbnNvbGUuZXJyb3IoZSk7IH0NCiAgICBhbGxEZXB0cyA9"
    b64 = b64 & "IEFycmF5LmlzQXJyYXkoZGF0YS5kZXBhcnRtZW50cykgPyBkYXRhLmRlcGFydG1lbnRzLnNsaWNlKCkgOiBbXTsNCiAgICBjb25zdCB0ID0gZGF0YS5jb2xsZWN0ZWRBdCB8fCAnLSc7DQogICAgY29uc3QgdGltZU9ubHkgPSB0LmluZGV4"
    b64 = b64 & "T2YoJyAnKSA+IDAgPyB0LnNwbGl0KCcgJylbMV0gOiB0Ow0KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtZXRhVGltZScpLnRleHRDb250ZW50ID0gJ+uniOyngOuniSDsl4XrjbDsnbTtirg6ICcgKyB0aW1lT25seTsNCiAgICBk"
    b64 = b64 & "b2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWV0YUNvdW50JykudGV4dENvbnRlbnQgPSAn67aA7LKYIOy0nTogJyArIGFsbERlcHRzLmxlbmd0aCArICfqsJwnOw0KICAgIHJlbmRlcigpOw0KICB9DQogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0"
    b64 = b64 & "b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaChidG4gPT4gew0KICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGZ1bmN0aW9uICgpIHsgc2V0RmlsdGVyKGJ0bi5nZXRBdHRyaWJ1dGUoJ2RhdGEtZmlsdGVyJyksIGJ0bik7IH0p"
    b64 = b64 & "Ow0KICB9KTsNCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2J0blJlbG9hZCcpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgZnVuY3Rpb24gKCkgeyBsb2NhdGlvbi5yZWxvYWQoKTsgfSk7DQogIGRvY3VtZW50LmdldEVsZW1lbnRC"
    b64 = b64 & "eUlkKCdmYWJTeW5jJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBmdW5jdGlvbiAoKSB7DQogICAgY29uc3QgaWNvbiA9IHRoaXMucXVlcnlTZWxlY3RvcignLm1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQnKTsNCiAgICBpY29uLmNs"
    b64 = b64 & "YXNzTGlzdC5hZGQoJ2FuaW1hdGUtc3BpbicpOw0KICAgIHNldFRpbWVvdXQoZnVuY3Rpb24gKCkgew0KICAgICAgaWNvbi5jbGFzc0xpc3QucmVtb3ZlKCdhbmltYXRlLXNwaW4nKTsNCiAgICAgIHdpbmRvdy5zY3JvbGxUbyh7IHRvcDog"
    b64 = b64 & "MCwgYmVoYXZpb3I6ICdzbW9vdGgnIH0pOw0KICAgICAgbG9jYXRpb24ucmVsb2FkKCk7DQogICAgfSwgODAwKTsNCiAgfSk7DQogIGxvYWQoKTsNCn0pKCk7DQo8L3NjcmlwdD4NCjwvYm9keT4NCjwvaHRtbD4NCg=="
    EmbeddedHtmlTemplateMobile = DecodeBase64Utf8(b64)
End Function

Private Function EmbeddedHtmlTemplate() As String
    EmbeddedHtmlTemplate = EmbeddedHtmlTemplatePc()
End Function

Private Function DecodeBase64Utf8(ByVal b64 As String) As String
    Dim doc As Object, node As Object, stm As Object
    On Error GoTo Fail
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = doc.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write node.nodeTypedValue
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "UTF-8"
    DecodeBase64Utf8 = stm.ReadText(-1)
    stm.Close
    Exit Function
Fail:
    DecodeBase64Utf8 = ""
End Function
