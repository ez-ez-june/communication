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
    templatePressPc = LoadHtmlTemplateFile("gov_major_press_pc_template.html", EmbeddedMajorPressTemplatePc())
    templatePressMobile = LoadHtmlTemplateFile("gov_major_press_mobile_template.html", EmbeddedMajorPressTemplateMobile())

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

' Write via PowerShell so Fasoo treats it as a native .html create (not Excel-derived).
Private Sub WriteTextUtf8(ByVal filePath As String, ByVal content As String)
    Dim b64 As String
    Dim psPath As String
    Dim cmd As String
    Dim sh As Object
    Dim ex As Object
    Dim errOut As String
    Dim rc As Long

    On Error Resume Next
    If Dir(filePath) <> "" Then Kill filePath
    On Error GoTo 0

    b64 = EncodeBase64Utf8(content)
    If Len(b64) = 0 Then Err.Raise vbObjectError + 910, , "UTF-8 Base64 encode failed"

    psPath = Replace(filePath, "'", "''")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
          """$b=[Console]::In.ReadToEnd().Trim();" & _
          "$b=$b -replace '\s','';" & _
          "[IO.File]::WriteAllBytes('" & psPath & "',[Convert]::FromBase64String($b))"""

    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(cmd)
    ex.StdIn.Write b64
    ex.StdIn.Close

    Do While ex.Status = 0
        DoEvents
    Loop

    errOut = ""
    On Error Resume Next
    errOut = ex.StdErr.ReadAll
    rc = ex.ExitCode
    On Error GoTo 0

    If rc <> 0 Or Dir(filePath) = "" Then
        Err.Raise vbObjectError + 911, , "PowerShell HTML write failed" & _
            IIf(Len(errOut) > 0, ": " & Left$(errOut, 200), " (exit " & CStr(rc) & ")")
    End If
End Sub

Private Function EncodeBase64Utf8(ByVal s As String) As String
    Dim stm As Object
    Dim doc As Object
    Dim node As Object
    On Error GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = doc.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = stm.Read
    stm.Close
    EncodeBase64Utf8 = Replace(Replace(Replace(node.Text, vbCrLf, ""), vbLf, ""), vbCr, "")
    EncodeBase64Utf8 = Replace(EncodeBase64Utf8, " ", "")
    Exit Function
Fail:
    EncodeBase64Utf8 = ""
End Function

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
    ' ???????? ????????????
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
    b64 = b64 & "77u/PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0x"
    b64 = b64 & "LjAiIG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KCV67aA67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wgKFBDKTwvdGl0bGU+DQo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4udGFpbHdpbmRjc3MuY29tP3BsdWdpbnM9Zm9ybXMsY29u"
    b64 = b64 & "dGFpbmVyLXF1ZXJpZXMiPjwvc2NyaXB0Pg0KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1QdWJsaWMrU2Fuczp3Z2h0QDQwMDs1MDA7NjAwOzcwMCZhbXA7ZGlzcGxheT1zd2FwIiByZWw9"
    b64 = b64 & "InN0eWxlc2hlZXQiLz4NCjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9TWF0ZXJpYWwrU3ltYm9scytPdXRsaW5lZDp3Z2h0LEZJTExAMTAwLi43MDAsMC4uMSZhbXA7ZGlzcGxheT1zd2Fw"
    b64 = b64 & "IiByZWw9InN0eWxlc2hlZXQiLz4NCjxzY3JpcHQgaWQ9InRhaWx3aW5kLWNvbmZpZyI+DQogICAgICAgIHRhaWx3aW5kLmNvbmZpZyA9IHsNCiAgICAgICAgICAgIGRhcmtNb2RlOiAiY2xhc3MiLA0KICAgICAgICAgICAgdGhlbWU6IHsN"
    b64 = b64 & "CiAgICAgICAgICAgICAgICBleHRlbmQ6IHsNCiAgICAgICAgICAgICAgICAgICAgImNvbG9ycyI6IHsNCiAgICAgICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLXZhcmlhbnQiOiAiI2U0ZThlZSIsDQogICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAic3VyZmFjZS1icmlnaHQiOiAiI2Y3ZjhmYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1jb250YWluZXIiOiAiIzJmNDA1NSIsDQogICAgICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1jb250YWlu"
    b64 = b64 & "ZXIiOiAiI2I3Y2NlMCIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tZXJyb3IiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS10aW50IjogIiM2YjhmYjgiLA0KICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgInByaW1hcnkiOiAiIzZiOGZiOCIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc2Vjb25kYXJ5LWNvbnRhaW5lciI6ICIjNWE0MzM2IiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJ0ZXJ0aWFyeSI6ICIjYjA4OTc4"
    b64 = b64 & "IiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJwcmltYXJ5LWZpeGVkLWRpbSI6ICIjYzVkNmU4IiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1zZWNvbmRhcnktZml4ZWQiOiAiIzRhMzgyZSIsDQogICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAib3V0bGluZS12YXJpYW50IjogIiNkNWRhZTIiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWZpeGVkIjogIiNmMGUwZDgiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1jb250YWlu"
    b64 = b64 & "ZXIiOiAiI2YwZDVjMCIsDQogICAgICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1maXhlZCI6ICIjZGNlN2YyIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXN1cmZhY2UiOiAiIzNhNDI1MCIsDQogICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAib24tdGVydGlhcnkiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1kaW0iOiAiI2U0ZThlZSIsDQogICAgICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWluZXIt"
    b64 = b64 & "aGlnaCI6ICIjZWNlZmYzIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1wcmltYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzRhNjc4NSIsDQogICAgICAgICAgICAgICAgICAgICAgICAiaW52ZXJzZS1vbi1zdXJmYWNlIjogIiNmNGY2"
    b64 = b64 & "ZjgiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyIjogIiNlZWYxZjUiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeSI6ICIjYzk5NTZlIiwNCiAgICAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICJvbi1wcmltYXJ5IjogIiNmZmZmZmYiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyLWxvdyI6ICIjZjRmNmY4IiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1zZWNvbmRhcnkiOiAiI2ZmZmZm"
    b64 = b64 & "ZiIsDQogICAgICAgICAgICAgICAgICAgICAgICAic2Vjb25kYXJ5LWZpeGVkLWRpbSI6ICIjZThjOWIwIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJ0ZXJ0aWFyeS1maXhlZC1kaW0iOiAiI2UwYzhiYyIsDQogICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAib24tc3VyZmFjZS12YXJpYW50IjogIiM1YTY0NzIiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UiOiAiI2Y3ZjhmYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAic2Vjb25kYXJ5LWZpeGVkIjog"
    b64 = b64 & "IiNmNWU2ZDgiLA0KICAgICAgICAgICAgICAgICAgICAgICAgIm91dGxpbmUiOiAiIzlhYTNiMCIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktY29udGFpbmVyIjogIiNmZmY4ZjUiLA0KICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyLWhpZ2hlc3QiOiAiI2U0ZThlZSIsDQogICAgICAgICAgICAgICAgICAgICAgICAidGVydGlhcnktY29udGFpbmVyIjogIiNkNGI1YTUiLA0KICAgICAgICAgICAgICAgICAgICAgICAgInN1"
    b64 = b64 & "cmZhY2UtY29udGFpbmVyLWxvd2VzdCI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJlcnJvci1jb250YWluZXIiOiAiI2Y1ZDVkMiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tYmFja2dyb3VuZCI6ICIj"
    b64 = b64 & "MmEzMzQwIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXByaW1hcnkiOiAiI2M1ZDZlOCIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc3VyZmFjZSI6ICIjMmEzMzQwIiwNCiAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICJvbi10ZXJ0aWFyeS1maXhlZC12YXJpYW50IjogIiM3YTVhNGMiLA0KICAgICAgICAgICAgICAgICAgICAgICAgImJhY2tncm91bmQiOiAiI2Y3ZjhmYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1maXhl"
    b64 = b64 & "ZCI6ICIjMmY0MDU1IiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJvbi1lcnJvci1jb250YWluZXIiOiAiIzdhMmUyYSIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tc2Vjb25kYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzhhNmE1"
    b64 = b64 & "MiIsDQogICAgICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktZml4ZWQiOiAiIzRhMzQyYyIsDQogICAgICAgICAgICAgICAgICAgICAgICAiZXJyb3IiOiAiI2MwNzA2YyINCiAgICAgICAgICAgICAgICAgICAgfSwNCiAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgImJvcmRlclJhZGl1cyI6IHsNCiAgICAgICAgICAgICAgICAgICAgICAgICJERUZBVUxUIjogIjAuMjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImxnIjogIjAuNXJlbSIsDQogICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAieGwiOiAiMC43NXJlbSIsDQogICAgICAgICAgICAgICAgICAgICAgICAiZnVsbCI6ICI5OTk5cHgiDQogICAgICAgICAgICAgICAgICAgIH0sDQogICAgICAgICAgICAgICAgICAgICJzcGFjaW5nIjogew0KICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgICAgInN0YWNrLWdhcCI6ICIwLjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImNvbnRhaW5lci1wYWRkaW5nIjogIjJyZW0iLA0KICAgICAgICAgICAgICAgICAgICAgICAgImNhcmQtcGFkZGluZyI6"
    b64 = b64 & "ICIxcmVtIiwNCiAgICAgICAgICAgICAgICAgICAgICAgICJncmlkLWd1dHRlciI6ICIxLjI1cmVtIg0KICAgICAgICAgICAgICAgICAgICB9LA0KICAgICAgICAgICAgICAgICAgICAiZm9udEZhbWlseSI6IHsNCiAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICAgICJsaXN0LWl0ZW0iOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICAgICAiaGVhZGVyLXRpdGxlIjogWyJQdWJsaWMgU2FucyJdLA0KICAgICAgICAgICAgICAgICAgICAgICAgIm1ldGEtZGF0YSI6"
    b64 = b64 & "IFsiUHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgICAgICJib2FyZC10aXRsZSI6IFsiUHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgICAgICJidXR0b24tdGV4dCI6IFsiUHVibGljIFNhbnMiXQ0KICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgICB9LA0KICAgICAgICAgICAgICAgICAgICAiZm9udFNpemUiOiB7DQogICAgICAgICAgICAgICAgICAgICAgICAibGlzdC1pdGVtIjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdo"
    b64 = b64 & "dCI6ICI1MDAifV0sDQogICAgICAgICAgICAgICAgICAgICAgICAiaGVhZGVyLXRpdGxlIjogWyIyNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjMycHgiLCAibGV0dGVyU3BhY2luZyI6ICItMC4wMmVtIiwgImZvbnRXZWlnaHQiOiAiNzAwIn1d"
    b64 = b64 & "LA0KICAgICAgICAgICAgICAgICAgICAgICAgIm1ldGEtZGF0YSI6IFsiMTJweCIsIHsibGluZUhlaWdodCI6ICIxNnB4IiwgImZvbnRXZWlnaHQiOiAiNDAwIn1dLA0KICAgICAgICAgICAgICAgICAgICAgICAgImJvYXJkLXRpdGxlIjog"
    b64 = b64 & "WyIxNnB4IiwgeyJsaW5lSGVpZ2h0IjogIjI0cHgiLCAiZm9udFdlaWdodCI6ICI3MDAifV0sDQogICAgICAgICAgICAgICAgICAgICAgICAiYnV0dG9uLXRleHQiOiBbIjE0cHgiLCB7ImxpbmVIZWlnaHQiOiAiMjBweCIsICJmb250V2Vp"
    b64 = b64 & "Z2h0IjogIjYwMCJ9XQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfSwNCiAgICAgICAgICAgIH0sDQogICAgICAgIH0NCiAgICA8L3NjcmlwdD4NCjxzdHlsZT4NCiAgICAgICAgLm1hdGVyaWFsLXN5bWJvbHMt"
    b64 = b64 & "b3V0bGluZWQgew0KICAgICAgICAgICAgZm9udC12YXJpYXRpb24tc2V0dGluZ3M6ICdGSUxMJyAwLCAnd2dodCcgNDAwLCAnR1JBRCcgMCwgJ29wc3onIDI0Ow0KICAgICAgICAgICAgdmVydGljYWwtYWxpZ246IG1pZGRsZTsNCiAgICAg"
    b64 = b64 & "ICAgfQ0KICAgICAgICBib2R5IHsgYmFja2dyb3VuZC1jb2xvcjogI2Y3ZjhmYTsgfQ0KICAgICAgICAuYm9hcmQtY2FyZCB7DQogICAgICAgICAgICB0cmFuc2l0aW9uOiB0cmFuc2Zvcm0gMC4ycyBlYXNlLCBib3gtc2hhZG93IDAuMnMg"
    b64 = b64 & "ZWFzZSwgb3BhY2l0eSAwLjEycyBlYXNlOw0KICAgICAgICAgICAgY3Vyc29yOiBncmFiOw0KICAgICAgICAgICAgdXNlci1zZWxlY3Q6IG5vbmU7DQogICAgICAgIH0NCiAgICAgICAgLmJvYXJkLWNhcmQ6YWN0aXZlIHsgY3Vyc29yOiBn"
    b64 = b64 & "cmFiYmluZzsgfQ0KICAgICAgICAuYm9hcmQtY2FyZDpob3ZlciB7DQogICAgICAgICAgICB0cmFuc2Zvcm06IHRyYW5zbGF0ZVkoLTRweCk7DQogICAgICAgICAgICBib3gtc2hhZG93OiAwIDEwcHggMjVweCAtNXB4IHJnYmEoMCwgMCwg"
    b64 = b64 & "MCwgMC4wOCksIDAgOHB4IDEwcHggLTZweCByZ2JhKDAsIDAsIDAsIDAuMDYpOw0KICAgICAgICB9DQogICAgICAgIC5ib2FyZC1jYXJkLmRyYWdnaW5nIHsgb3BhY2l0eTogMC40NTsgdHJhbnNmb3JtOiBzY2FsZSgwLjk4KTsgfQ0KICAg"
    b64 = b64 & "ICAgICAuYm9hcmQtY2FyZC5kcmFnLW92ZXIgeyBvdXRsaW5lOiAycHggZGFzaGVkICM2YjhmYjg7IG91dGxpbmUtb2Zmc2V0OiAycHg7IH0NCiAgICAgICAgLmJvYXJkLWNhcmQgYSB7IGN1cnNvcjogcG9pbnRlcjsgdXNlci1zZWxlY3Q6"
    b64 = b64 & "IHRleHQ7IH0NCiAgICAgICAgLmZpbHRlci1hY3RpdmUgew0KICAgICAgICAgICAgYmFja2dyb3VuZC1jb2xvcjogI2YwZDVjMCAhaW1wb3J0YW50Ow0KICAgICAgICAgICAgY29sb3I6ICM1YTQzMzYgIWltcG9ydGFudDsNCiAgICAgICAg"
    b64 = b64 & "fQ0KPC9zdHlsZT4NCjwvaGVhZD4NCjxib2R5IGNsYXNzPSJmb250LWxpc3QtaXRlbSB0ZXh0LW9uLXN1cmZhY2UiPg0KPGhlYWRlciBjbGFzcz0iYmctcHJpbWFyeSBzaGFkb3ctbWQgc3RpY2t5IHRvcC0wIHotNTAgcm91bmRlZC1iLXhs"
    b64 = b64 & "Ij4NCjxkaXYgY2xhc3M9Im1heC13LVsxNDQwcHhdIG14LWF1dG8gcHgtY29udGFpbmVyLXBhZGRpbmcgcHktNiBmbGV4IGZsZXgtY29sIG1kOmZsZXgtcm93IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgZ2FwLTQiPg0KPGRpdiBj"
    b64 = b64 & "bGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTMiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1vbi1wcmltYXJ5IHRleHQtM3hsIj5hY2NvdW50X2JhbGFuY2U8L3NwYW4+DQo8ZGl2Pg0KPGgxIGNs"
    b64 = b64 & "YXNzPSJmb250LWhlYWRlci10aXRsZSB0ZXh0LWhlYWRlci10aXRsZSB0ZXh0LW9uLXByaW1hcnkiPuygleu2gOu2gOyymCDqs7Xsp4Dsgqztla0gKOy1nOq3vCA16rCcKTwvaDE+DQo8cCBjbGFzcz0idGV4dC1vbi1wcmltYXJ5LzgwIGZv"
    b64 = b64 & "bnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIG10LTEiPuy5tOuTnOulvCDrk5zrnpjqt7jtlbQg7JyE7LmY66W8IOuwlOq/gCDsiJgg7J6I7Iq164uI64ukIMK3IOygnOuqqSDtgbTrpq0g7IucIOyDiCDssL08L3A+DQo8L2Rpdj4NCjwv"
    b64 = b64 & "ZGl2Pg0KPGJ1dHRvbiBpZD0iYnRuUmVsb2FkIiB0eXBlPSJidXR0b24iIGNsYXNzPSJiZy1wcmltYXJ5LWNvbnRhaW5lciB0ZXh0LW9uLXByaW1hcnktY29udGFpbmVyIGZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4dCBweC02"
    b64 = b64 & "IHB5LTMgcm91bmRlZC1mdWxsIGZsZXggaXRlbXMtY2VudGVyIGdhcC0yIHNoYWRvdy1sZyBob3ZlcjpiZy1wcmltYXJ5LWNvbnRhaW5lci85MCBhY3RpdmU6c2NhbGUtOTUgdHJhbnNpdGlvbi1hbGwiPg0KPHNwYW4gY2xhc3M9Im1hdGVy"
    b64 = b64 & "aWFsLXN5bWJvbHMtb3V0bGluZWQiPnJlZnJlc2g8L3NwYW4+DQrtmZTrqbQg7IOI66Gc6rOg7LmoDQo8L2J1dHRvbj4NCjwvZGl2Pg0KPC9oZWFkZXI+DQoNCjxkaXYgY2xhc3M9Im1heC13LVsxNDQwcHhdIG14LWF1dG8gcHgtY29udGFp"
    b64 = b64 & "bmVyLXBhZGRpbmcgbXQtNCI+DQo8ZGl2IGNsYXNzPSJiZy1zdXJmYWNlLWNvbnRhaW5lciByb3VuZGVkLWxnIGJvcmRlciBib3JkZXItb3V0bGluZS12YXJpYW50IHB4LTQgcHktMiBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiI+DQo8c3Bh"
    b64 = b64 & "biBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB0ZXh0LXByaW1hcnkgdGV4dC1zbSI+c2NoZWR1bGU8L3NwYW4+DQo8cCBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlh"
    b64 = b64 & "bnQiIGlkPSJtZXRhTGluZSI+642w7J207YSwIOykgOu5hCDspJHigKY8L3A+DQo8L2Rpdj4NCjwvZGl2Pg0KDQo8ZGl2IGNsYXNzPSJtYXgtdy1bMTQ0MHB4XSBteC1hdXRvIHB4LWNvbnRhaW5lci1wYWRkaW5nIHB5LTggZmxleCBnYXAt"
    b64 = b64 & "OCI+DQo8YXNpZGUgY2xhc3M9ImhpZGRlbiBsZzpmbGV4IGZsZXgtY29sIHctNjQgc2hyaW5rLTAgZ2FwLTQiPg0KPGRpdiBjbGFzcz0iYmctc3VyZmFjZS1jb250YWluZXIgcm91bmRlZC14bCBwLTQgYm9yZGVyIGJvcmRlci1vdXRsaW5l"
    b64 = b64 & "LXZhcmlhbnQiPg0KPGgyIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUgdGV4dC1wcmltYXJ5IG1iLTQgZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0"
    b64 = b64 & "bGluZWQiPmZpbHRlcl9hbHQ8L3NwYW4+DQrrtoDsspgg7ZWE7YSwDQo8L2gyPg0KPG5hdiBjbGFzcz0ic3BhY2UteS0xIiBpZD0iZmlsdGVyTmF2Ij4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0iYWxsIiBjbGFzcz0i"
    b64 = b64 & "ZmlsdGVyLWJ0biBmaWx0ZXItYWN0aXZlIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0ZXh0LWxlZnQiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFs"
    b64 = b64 & "LXN5bWJvbHMtb3V0bGluZWQiPmRhc2hib2FyZDwvc3Bhbj7soITssrQg67aA7LKYDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0ibWFqb3IiIGNsYXNzPSJmaWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0"
    b64 = b64 & "ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1s"
    b64 = b64 & "ZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5zdGFyPC9zcGFuPuyjvOyalOu2gOyymA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImVjb25vbXkiIGNsYXNzPSJm"
    b64 = b64 & "aWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQg"
    b64 = b64 & "dHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5idXNpbmVzc19jZW50ZXI8L3NwYW4+7IKw7JeFL+qyveygnA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1"
    b64 = b64 & "dHRvbiIgZGF0YS1maWx0ZXI9InNvY2lldHkiIGNsYXNzPSJmaWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4"
    b64 = b64 & "dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj50aGVhdGVyX2NvbWVkeTwvc3Bhbj7sgqztmowv"
    b64 = b64 & "66y47ZmUDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0iYWRtaW4iIGNsYXNzPSJmaWx0ZXItYnRuIHctZnVsbCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMyBwLTIgaG92ZXI6Ymctc3VyZmFjZS12YXJp"
    b64 = b64 & "YW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1sZWZ0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxp"
    b64 = b64 & "bmVkIj5hY2NvdW50X2JhbGFuY2U8L3NwYW4+7ZaJ7KCVL+yViOyghA0KPC9idXR0b24+DQo8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1maWx0ZXI9ImRpcGxvbWFjeSIgY2xhc3M9ImZpbHRlci1idG4gdy1mdWxsIGZsZXggaXRlbXMt"
    b64 = b64 & "Y2VudGVyIGdhcC0zIHAtMiBob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCB0cmFuc2l0aW9uLWNvbG9ycyB0ZXh0LWxlZnQi"
    b64 = b64 & "Pg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnB1YmxpYzwvc3Bhbj7smbjqtZAv7JWI67O0DQo8L2J1dHRvbj4NCjwvbmF2Pg0KPC9kaXY+DQo8ZGl2IGNsYXNzPSJiZy1zdXJmYWNlLWNvbnRhaW5lci1sb3cg"
    b64 = b64 & "cm91bmRlZC14bCBwLTQgYm9yZGVyIGJvcmRlci1vdXRsaW5lLXZhcmlhbnQgZmxleCBmbGV4LWNvbCBpdGVtcy1jZW50ZXIgdGV4dC1jZW50ZXIiPg0KPGRpdiBjbGFzcz0idy0xNiBoLTE2IGJnLXByaW1hcnkvMTAgcm91bmRlZC1mdWxs"
    b64 = b64 & "IGZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktY2VudGVyIG1iLTMiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1wcmltYXJ5IHRleHQtM3hsIj5pbmZvPC9zcGFuPg0KPC9kaXY+DQo8cCBjbGFzcz0i"
    b64 = b64 & "Zm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgbWItNCI+7JeR7IWA7JeQ7IScIOyImOynke2VnCDstZzsi6Ag6rO17KeA7IKs7ZWt7J2EIOuztOuTnCDtmJXtg5zroZwg67O07Jes7KSN64uI"
    b64 = b64 & "64ukLjwvcD4NCjxhIGhyZWY9Imdvdl9ub3RpY2VfYm9hcmRfbW9iaWxlLmh0bWwiIGNsYXNzPSJ3LWZ1bGwgcHktMiBiZy1vdXRsaW5lLXZhcmlhbnQgdGV4dC1vbi1zdXJmYWNlIGZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4"
    b64 = b64 & "dCByb3VuZGVkLWxnIGhvdmVyOmJnLW91dGxpbmUgdHJhbnNpdGlvbi1jb2xvcnMgdGV4dC1jZW50ZXIgbm8tdW5kZXJsaW5lIj7rqqjrsJTsnbwg7Y6Y7J207KeAPC9hPg0KPC9kaXY+DQo8L2FzaWRlPg0KDQo8bWFpbiBjbGFzcz0iZmxl"
    b64 = b64 & "eC1ncm93Ij4NCjxkaXYgaWQ9Im1ham9yQWN0aW9ucyIgY2xhc3M9ImhpZGRlbiBtYi00IGZsZXggZmxleC13cmFwIGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWJldHdlZW4gZ2FwLTMgYmctc3VyZmFjZS1jb250YWluZXIgcm91bmRlZC14bCBi"
    b64 = b64 & "b3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBweC00IHB5LTMiPg0KPHAgY2xhc3M9ImZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIHRleHQtb24tc3VyZmFjZS12YXJpYW50Ij7so7zsmpTrtoDsspggNOqwnCDCtyDsgrDsl4Xt"
    b64 = b64 & "hrXsg4HrtoAg4oaSIOq4sO2bhOyXkOuEiOyngO2ZmOqyveu2gCDihpIg6rOg7Jqp64W464+Z67aAIOKGkiDsmbjqtZDrtoA8L3A+DQo8YSBpZD0iYnRuUHJlc3NUb2dldGhlciIgaHJlZj0iZ292X21ham9yX3ByZXNzX2JvYXJkLmh0bWwi"
    b64 = b64 & "IGNsYXNzPSJpbmxpbmUtZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgYmctc2Vjb25kYXJ5IHRleHQtb24tc2Vjb25kYXJ5IGZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4dCBweC01IHB5LTIuNSByb3VuZGVkLWZ1bGwgc2hhZG93"
    b64 = b64 & "IGhvdmVyOm9wYWNpdHktOTAgbm8tdW5kZXJsaW5lIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE4cHhdIj5uZXdzcGFwZXI8L3NwYW4+67O064+E7J6Q66OMIO2VqOq7mOuztOq4sA0KPC9hPg0K"
    b64 = b64 & "PC9kaXY+DQo8ZGl2IGNsYXNzPSJncmlkIGdyaWQtY29scy0xIG1kOmdyaWQtY29scy0yIHhsOmdyaWQtY29scy0zIDJ4bDpncmlkLWNvbHMtNCBnYXAtZ3JpZC1ndXR0ZXIiIGlkPSJib2FyZCI+PC9kaXY+DQo8L21haW4+DQo8L2Rpdj4N"
    b64 = b64 & "Cg0KPGZvb3RlciBjbGFzcz0iYmctc3VyZmFjZS1kaW0gYm9yZGVyLXQgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBtdC0xMiBweS04IHB4LWNvbnRhaW5lci1wYWRkaW5nIj4NCjxkaXYgY2xhc3M9Im1heC13LVsxNDQwcHhdIG14LWF1dG8g"
    b64 = b64 & "ZmxleCBmbGV4LWNvbCBtZDpmbGV4LXJvdyBqdXN0aWZ5LWJldHdlZW4gaXRlbXMtY2VudGVyIGdhcC02Ij4NCjxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91"
    b64 = b64 & "dGxpbmVkIHRleHQtcHJpbWFyeSI+YWNjb3VudF9iYWxhbmNlPC9zcGFuPg0KPHNwYW4gY2xhc3M9ImZvbnQtYm9sZCB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCI+64yA7ZWc66+86rWtIOygleu2gCDrtoDsspgg6rO17KeA7IKs7ZWtIOuz"
    b64 = b64 & "tOuTnDwvc3Bhbj4NCjwvZGl2Pg0KPG5hdiBjbGFzcz0iZmxleCBnYXAtNiI+DQo8YSBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgaG92ZXI6dGV4dC1wcmltYXJ5IHRyYW5z"
    b64 = b64 & "aXRpb24tY29sb3JzIiBocmVmPSJnb3Zfbm90aWNlX2JvYXJkLmh0bWwiPlBDPC9hPg0KPGEgY2xhc3M9ImZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIHRleHQtb24tc3VyZmFjZS12YXJpYW50IGhvdmVyOnRleHQtcHJpbWFyeSB0"
    b64 = b64 & "cmFuc2l0aW9uLWNvbG9ycyIgaHJlZj0iZ292X25vdGljZV9ib2FyZF9tb2JpbGUuaHRtbCI+66qo67CU7J28PC9hPg0KPC9uYXY+DQo8cCBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZh"
    b64 = b64 & "cmlhbnQvNzAiPsKpIOuMgO2VnOuvvOq1rSDsoJXrtoAg67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wg7ISc67mE7IqkPC9wPg0KPC9kaXY+DQo8L2Zvb3Rlcj4NCg0KPHNjcmlwdCBpZD0ibm90aWNlLWRhdGEiIHR5cGU9ImFwcGxpY2F0"
    b64 = b64 & "aW9uL2pzb24iPg0KJSVOT1RJQ0VfSlNPTiUlDQo8L3NjcmlwdD4NCjxzY3JpcHQ+DQooZnVuY3Rpb24gKCkgew0KICBjb25zdCBTVE9SQUdFX0tFWSA9ICdnb3ZOb3RpY2VDYXJkT3JkZXIudjEnOw0KICBjb25zdCBQQUxFVFRFID0gWycj"
    b64 = b64 & "YjdkMGVhJywnI2M1YzhlOCcsJyNiOGRkZDYnLCcjZjBkMGIwJywnI2Q0YzRlOCcsJyNjNWUwYzgnLCcjZjBjNGM0JywnI2I4ZGNlMCcsJyNjOGUwZDQnLCcjYjBjNGUwJ107DQogIGNvbnN0IENBVEVHT1JZID0gew0KICAgICfsnqzsoJXq"
    b64 = b64 & "sr3soJzrtoAnOidlY29ub215Jywn6rO87ZWZ6riw7Iig7KCV67O07Ya17Iug67aAJzonZWNvbm9teScsJ+yCsOyXhe2GteyDgeu2gCc6J2Vjb25vbXknLCfspJHshozrsqTsspjquLDsl4XrtoAnOidlY29ub215JywNCiAgICAn6rWt7Yag"
    b64 = b64 & "6rWQ7Ya167aAJzonZWNvbm9teScsJ+uGjeumvOy2leyCsOyLne2SiOu2gCc6J2Vjb25vbXknLCftlbTslpHsiJjsgrDrtoAnOidlY29ub215Jywn6riw7ZuE7JeQ64SI7KeA7ZmY6rK967aAJzonZWNvbm9teScsJ+q4sO2ajeyYiOyCsOyy"
    b64 = b64 & "mCc6J2Vjb25vbXknLA0KICAgICfqtZDsnKHrtoAnOidzb2NpZXR5Jywn67O06rG067O17KeA67aAJzonc29jaWV0eScsJ+qzoOyaqeuFuOuPmeu2gCc6J3NvY2lldHknLCfrrLjtmZTssrTsnKHqtIDqtJHrtoAnOidzb2NpZXR5Jywn7ISx"
    b64 = b64 & "7Y+J65Ox6rCA7KGx67aAJzonc29jaWV0eScsJ+yLne2SiOydmOyVve2SiOyViOyghOyymCc6J3NvY2lldHknLA0KICAgICftlonsoJXslYjsoITrtoAnOidhZG1pbicsJ+uyleustOu2gCc6J2FkbWluJywn7J247IKs7ZiB7Iug7LKYJzon"
    b64 = b64 & "YWRtaW4nLCfrspXsoJzsspgnOidhZG1pbicsJ+q1reqwgOuztO2biOu2gCc6J2FkbWluJywNCiAgICAn7Jm46rWQ67aAJzonZGlwbG9tYWN5Jywn7Ya17J2867aAJzonZGlwbG9tYWN5Jywn6rWt67Cp67aAJzonZGlwbG9tYWN5Jw0KICB9"
    b64 = b64 & "Ow0KICBjb25zdCBNQUpPUl9PUkRFUiA9IFsn7IKw7JeF7Ya17IOB67aAJywn6riw7ZuE7JeQ64SI7KeA7ZmY6rK967aAJywn6rOg7Jqp64W464+Z67aAJywn7Jm46rWQ67aAJ107DQogIGxldCBkcmFnU3JjID0gbnVsbDsNCiAgbGV0IGN1"
    b64 = b64 & "cnJlbnRGaWx0ZXIgPSAnYWxsJzsNCiAgbGV0IGFsbERlcHRzID0gW107DQoNCiAgZnVuY3Rpb24gY29sb3JGb3IobmFtZSkgew0KICAgIGlmIChuYW1lID09PSAn7IKw7JeF7Ya17IOB67aAJykgcmV0dXJuICcjZjBkMGIwJzsNCiAgICBs"
    b64 = b64 & "ZXQgaWR4ID0gMDsNCiAgICBmb3IgKGxldCBpID0gMDsgaSA8IG5hbWUubGVuZ3RoOyBpKyspIGlkeCA9IChpZHggKyBuYW1lLmNoYXJDb2RlQXQoaSkpICUgUEFMRVRURS5sZW5ndGg7DQogICAgcmV0dXJuIFBBTEVUVEVbaWR4XTsNCiAg"
    b64 = b64 & "fQ0KICBmdW5jdGlvbiBjYXRlZ29yeUZvcihuYW1lKSB7IHJldHVybiBDQVRFR09SWVtuYW1lXSB8fCAnYWRtaW4nOyB9DQogIGZ1bmN0aW9uIGVzYyhzKSB7DQogICAgcmV0dXJuIFN0cmluZyhzID8/ICcnKS5yZXBsYWNlKC8mL2csJyZh"
    b64 = b64 & "bXA7JykucmVwbGFjZSgvPC9nLCcmbHQ7JykucmVwbGFjZSgvPi9nLCcmZ3Q7JykucmVwbGFjZSgvIi9nLCcmcXVvdDsnKTsNCiAgfQ0KICBmdW5jdGlvbiBzYXZlT3JkZXIoKSB7DQogICAgY29uc3QgbmFtZXMgPSBBcnJheS5mcm9tKGRv"
    b64 = b64 & "Y3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJyNib2FyZCAuYm9hcmQtY2FyZCcpKS5tYXAoYyA9PiBjLmdldEF0dHJpYnV0ZSgnZGF0YS1uYW1lJykpOw0KICAgIHRyeSB7IGxvY2FsU3RvcmFnZS5zZXRJdGVtKFNUT1JBR0VfS0VZLCBKU09O"
    b64 = b64 & "LnN0cmluZ2lmeShuYW1lcykpOyB9IGNhdGNoIChlKSB7fQ0KICB9DQogIGZ1bmN0aW9uIGFwcGx5U2F2ZWRPcmRlcihkZXB0cykgew0KICAgIHRyeSB7DQogICAgICBjb25zdCByYXcgPSBsb2NhbFN0b3JhZ2UuZ2V0SXRlbShTVE9SQUdF"
    b64 = b64 & "X0tFWSk7DQogICAgICBpZiAoIXJhdykgcmV0dXJuIGRlcHRzOw0KICAgICAgY29uc3Qgb3JkZXIgPSBKU09OLnBhcnNlKHJhdyk7DQogICAgICBpZiAoIUFycmF5LmlzQXJyYXkob3JkZXIpIHx8ICFvcmRlci5sZW5ndGgpIHJldHVybiBk"
    b64 = b64 & "ZXB0czsNCiAgICAgIGNvbnN0IG1hcCA9IHt9Ow0KICAgICAgZGVwdHMuZm9yRWFjaChkID0+IHsgbWFwW2QubmFtZV0gPSBkOyB9KTsNCiAgICAgIGNvbnN0IHNvcnRlZCA9IFtdOw0KICAgICAgb3JkZXIuZm9yRWFjaChuYW1lID0+IHsg"
    b64 = b64 & "aWYgKG1hcFtuYW1lXSkgeyBzb3J0ZWQucHVzaChtYXBbbmFtZV0pOyBkZWxldGUgbWFwW25hbWVdOyB9IH0pOw0KICAgICAgT2JqZWN0LmtleXMobWFwKS5mb3JFYWNoKGsgPT4gc29ydGVkLnB1c2gobWFwW2tdKSk7DQogICAgICByZXR1"
    b64 = b64 & "cm4gc29ydGVkOw0KICAgIH0gY2F0Y2ggKGUpIHsgcmV0dXJuIGRlcHRzOyB9DQogIH0NCiAgZnVuY3Rpb24gYmluZERyYWcoY2FyZCkgew0KICAgIGNhcmQuc2V0QXR0cmlidXRlKCdkcmFnZ2FibGUnLCAndHJ1ZScpOw0KICAgIGNhcmQu"
    b64 = b64 & "YWRkRXZlbnRMaXN0ZW5lcignZHJhZ3N0YXJ0JywgZnVuY3Rpb24gKGUpIHsNCiAgICAgIGlmIChlLnRhcmdldCAmJiBlLnRhcmdldC5jbG9zZXN0ICYmIGUudGFyZ2V0LmNsb3Nlc3QoJ2EnKSkgeyBlLnByZXZlbnREZWZhdWx0KCk7IHJl"
    b64 = b64 & "dHVybjsgfQ0KICAgICAgZHJhZ1NyYyA9IGNhcmQ7DQogICAgICBjYXJkLmNsYXNzTGlzdC5hZGQoJ2RyYWdnaW5nJyk7DQogICAgICBlLmRhdGFUcmFuc2Zlci5lZmZlY3RBbGxvd2VkID0gJ21vdmUnOw0KICAgICAgdHJ5IHsgZS5kYXRh"
    b64 = b64 & "VHJhbnNmZXIuc2V0RGF0YSgndGV4dC9wbGFpbicsIGNhcmQuZ2V0QXR0cmlidXRlKCdkYXRhLW5hbWUnKSB8fCAnJyk7IH0gY2F0Y2ggKGVycikge30NCiAgICB9KTsNCiAgICBjYXJkLmFkZEV2ZW50TGlzdGVuZXIoJ2RyYWdlbmQnLCBm"
    b64 = b64 & "dW5jdGlvbiAoKSB7DQogICAgICBjYXJkLmNsYXNzTGlzdC5yZW1vdmUoJ2RyYWdnaW5nJyk7DQogICAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuYm9hcmQtY2FyZC5kcmFnLW92ZXInKS5mb3JFYWNoKGVsID0+IGVsLmNsYXNz"
    b64 = b64 & "TGlzdC5yZW1vdmUoJ2RyYWctb3ZlcicpKTsNCiAgICAgIGRyYWdTcmMgPSBudWxsOw0KICAgICAgc2F2ZU9yZGVyKCk7DQogICAgfSk7DQogICAgY2FyZC5hZGRFdmVudExpc3RlbmVyKCdkcmFnb3ZlcicsIGZ1bmN0aW9uIChlKSB7DQog"
    b64 = b64 & "ICAgICBlLnByZXZlbnREZWZhdWx0KCk7DQogICAgICBlLmRhdGFUcmFuc2Zlci5kcm9wRWZmZWN0ID0gJ21vdmUnOw0KICAgICAgaWYgKGRyYWdTcmMgJiYgZHJhZ1NyYyAhPT0gY2FyZCkgY2FyZC5jbGFzc0xpc3QuYWRkKCdkcmFnLW92"
    b64 = b64 & "ZXInKTsNCiAgICB9KTsNCiAgICBjYXJkLmFkZEV2ZW50TGlzdGVuZXIoJ2RyYWdsZWF2ZScsIGZ1bmN0aW9uICgpIHsgY2FyZC5jbGFzc0xpc3QucmVtb3ZlKCdkcmFnLW92ZXInKTsgfSk7DQogICAgY2FyZC5hZGRFdmVudExpc3RlbmVy"
    b64 = b64 & "KCdkcm9wJywgZnVuY3Rpb24gKGUpIHsNCiAgICAgIGUucHJldmVudERlZmF1bHQoKTsNCiAgICAgIGNhcmQuY2xhc3NMaXN0LnJlbW92ZSgnZHJhZy1vdmVyJyk7DQogICAgICBpZiAoIWRyYWdTcmMgfHwgZHJhZ1NyYyA9PT0gY2FyZCkg"
    b64 = b64 & "cmV0dXJuOw0KICAgICAgY29uc3QgYm9hcmQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYm9hcmQnKTsNCiAgICAgIGNvbnN0IGNhcmRzID0gQXJyYXkuZnJvbShib2FyZC5jaGlsZHJlbik7DQogICAgICBjb25zdCBmcm9tID0gY2Fy"
    b64 = b64 & "ZHMuaW5kZXhPZihkcmFnU3JjKTsNCiAgICAgIGNvbnN0IHRvID0gY2FyZHMuaW5kZXhPZihjYXJkKTsNCiAgICAgIGlmIChmcm9tIDwgMCB8fCB0byA8IDApIHJldHVybjsNCiAgICAgIGlmIChmcm9tIDwgdG8pIGJvYXJkLmluc2VydEJl"
    b64 = b64 & "Zm9yZShkcmFnU3JjLCBjYXJkLm5leHRTaWJsaW5nKTsNCiAgICAgIGVsc2UgYm9hcmQuaW5zZXJ0QmVmb3JlKGRyYWdTcmMsIGNhcmQpOw0KICAgICAgc2F2ZU9yZGVyKCk7DQogICAgfSk7DQogIH0NCiAgZnVuY3Rpb24gcmVuZGVyKCkg"
    b64 = b64 & "ew0KICAgIGNvbnN0IGJvYXJkID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2JvYXJkJyk7DQogICAgY29uc3QgbWFqb3JCb3ggPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWFqb3JBY3Rpb25zJyk7DQogICAgaWYgKG1ham9yQm94"
    b64 = b64 & "KSBtYWpvckJveC5jbGFzc0xpc3QudG9nZ2xlKCdoaWRkZW4nLCBjdXJyZW50RmlsdGVyICE9PSAnbWFqb3InKTsNCiAgICBsZXQgZGVwdHMgPSBhbGxEZXB0cy5zbGljZSgpOw0KICAgIGlmIChjdXJyZW50RmlsdGVyID09PSAnbWFqb3In"
    b64 = b64 & "KSB7DQogICAgICBjb25zdCBtYXAgPSB7fTsNCiAgICAgIGRlcHRzLmZvckVhY2goZCA9PiB7IG1hcFtkLm5hbWVdID0gZDsgfSk7DQogICAgICBkZXB0cyA9IE1BSk9SX09SREVSLm1hcChuID0+IG1hcFtuXSkuZmlsdGVyKEJvb2xlYW4p"
    b64 = b64 & "Ow0KICAgIH0gZWxzZSBpZiAoY3VycmVudEZpbHRlciAhPT0gJ2FsbCcpIHsNCiAgICAgIGRlcHRzID0gZGVwdHMuZmlsdGVyKGQgPT4gY2F0ZWdvcnlGb3IoZC5uYW1lKSA9PT0gY3VycmVudEZpbHRlcik7DQogICAgICBkZXB0cyA9IGFw"
    b64 = b64 & "cGx5U2F2ZWRPcmRlcihkZXB0cyk7DQogICAgfSBlbHNlIHsNCiAgICAgIGRlcHRzID0gYXBwbHlTYXZlZE9yZGVyKGRlcHRzKTsNCiAgICB9DQogICAgaWYgKCFkZXB0cy5sZW5ndGgpIHsNCiAgICAgIGJvYXJkLmlubmVySFRNTCA9ICc8"
    b64 = b64 & "ZGl2IGNsYXNzPSJjb2wtc3Bhbi1mdWxsIHRleHQtY2VudGVyIHB5LTE2IHRleHQtb24tc3VyZmFjZS12YXJpYW50IGJnLXN1cmZhY2UtY29udGFpbmVyIHJvdW5kZWQteGwgYm9yZGVyIGJvcmRlci1kYXNoZWQgYm9yZGVyLW91dGxpbmUt"
    b64 = b64 & "dmFyaWFudCI+7ZGc7Iuc7ZWgIOqzteyngCDrjbDsnbTthLDqsIAg7JeG7Iq164uI64ukLjwvZGl2Pic7DQogICAgICByZXR1cm47DQogICAgfQ0KICAgIGJvYXJkLmlubmVySFRNTCA9IGRlcHRzLm1hcChmdW5jdGlvbiAoZCkgew0KICAg"
    b64 = b64 & "ICAgY29uc3QgaGVhZCA9IGNvbG9yRm9yKGQubmFtZSk7DQogICAgICBjb25zdCBjYXQgPSBjYXRlZ29yeUZvcihkLm5hbWUpOw0KICAgICAgaWYgKGQub2sgPT09IGZhbHNlIHx8ICFkLnBvc3RzIHx8ICFkLnBvc3RzLmxlbmd0aCkgew0K"
    b64 = b64 & "ICAgICAgICByZXR1cm4gJzxhcnRpY2xlIGNsYXNzPSJib2FyZC1jYXJkIGJnLXN1cmZhY2UtY29udGFpbmVyLWxvd2VzdCByb3VuZGVkLXhsIHNoYWRvdy1bMHB4XzRweF8xMnB4X3JnYmEoMCwwLDAsMC4wNSldIGJvcmRlciBib3JkZXIt"
    b64 = b64 & "b3V0bGluZS12YXJpYW50IGZsZXggZmxleC1jb2wgaC1mdWxsIG92ZXJmbG93LWhpZGRlbiIgZGF0YS1uYW1lPSInICsgZXNjKGQubmFtZSkgKyAnIiBkYXRhLWNhdGVnb3J5PSInICsgY2F0ICsgJyI+JyArDQogICAgICAgICAgJzxoZWFk"
    b64 = b64 & "ZXIgY2xhc3M9InAtNCBmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIiIHN0eWxlPSJiYWNrZ3JvdW5kOicgKyBoZWFkICsgJyI+PGgzIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUiIHN0eWxlPSJj"
    b64 = b64 & "b2xvcjojM2Q0YTVjIj5bJyArIGVzYyhkLm5hbWUpICsgJ108L2gzPjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtc20iIHN0eWxlPSJjb2xvcjojNWE2NDcyIj5tb3JlX3ZlcnQ8L3NwYW4+PC9oZWFkZXI+"
    b64 = b64 & "JyArDQogICAgICAgICAgJzxkaXYgY2xhc3M9InAtNCB0ZXh0LWVycm9yIGZvbnQtYm9sZCB0ZXh0LXNtIGJnLWVycm9yLWNvbnRhaW5lciI+7ISk7KCVIO2OmOydtOyngOulvCDsiJjsoJXtlZjshLjsmpQuPC9kaXY+PC9hcnRpY2xlPic7"
    b64 = b64 & "DQogICAgICB9DQogICAgICBjb25zdCBpdGVtcyA9IGQucG9zdHMuc2xpY2UoMCwgNSkubWFwKGZ1bmN0aW9uIChwLCBpKSB7DQogICAgICAgIGNvbnN0IGJvcmRlciA9IGkgPCBNYXRoLm1pbihkLnBvc3RzLmxlbmd0aCwgNSkgLSAxID8g"
    b64 = b64 & "JyBib3JkZXItYiBib3JkZXItb3V0bGluZS12YXJpYW50LzMwJyA6ICcnOw0KICAgICAgICByZXR1cm4gJzxsaSBjbGFzcz0icHktMycgKyBib3JkZXIgKyAnIGZsZXggZ2FwLTMgaG92ZXI6Ymctc3VyZmFjZS1jb250YWluZXItbG93IHRy"
    b64 = b64 & "YW5zaXRpb24tY29sb3JzIHB4LTIgcm91bmRlZC1tZCI+JyArDQogICAgICAgICAgJzxzcGFuIGNsYXNzPSJmb250LWJvbGQgdGV4dC1zbSIgc3R5bGU9ImNvbG9yOiM2YjhmYjgiPicgKyAoaSArIDEpICsgJy48L3NwYW4+JyArDQogICAg"
    b64 = b64 & "ICAgICAgJzxhIGNsYXNzPSJ0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cnVuY2F0ZSIgaHJlZj0iJyArIGVzYyhwLnVybCkgKyAnIiB0YXJnZXQ9Il9ibGFuayIgcmVsPSJub29wZW5l"
    b64 = b64 & "ciBub3JlZmVycmVyIiBkcmFnZ2FibGU9ImZhbHNlIiB0aXRsZT0iJyArIGVzYyhwLnRpdGxlKSArICciPicgKyBlc2MocC50aXRsZSkgKyAnPC9hPjwvbGk+JzsNCiAgICAgIH0pLmpvaW4oJycpOw0KICAgICAgcmV0dXJuICc8YXJ0aWNs"
    b64 = b64 & "ZSBjbGFzcz0iYm9hcmQtY2FyZCBiZy1zdXJmYWNlLWNvbnRhaW5lci1sb3dlc3Qgcm91bmRlZC14bCBzaGFkb3ctWzBweF80cHhfMTJweF9yZ2JhKDAsMCwwLDAuMDUpXSBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBmbGV4IGZs"
    b64 = b64 & "ZXgtY29sIGgtZnVsbCBvdmVyZmxvdy1oaWRkZW4iIGRhdGEtbmFtZT0iJyArIGVzYyhkLm5hbWUpICsgJyIgZGF0YS1jYXRlZ29yeT0iJyArIGNhdCArICciPicgKw0KICAgICAgICAnPGhlYWRlciBjbGFzcz0icC00IGZsZXgganVzdGlm"
    b64 = b64 & "eS1iZXR3ZWVuIGl0ZW1zLWNlbnRlciIgc3R5bGU9ImJhY2tncm91bmQ6JyArIGhlYWQgKyAnIj48aDMgY2xhc3M9ImZvbnQtYm9hcmQtdGl0bGUgdGV4dC1ib2FyZC10aXRsZSIgc3R5bGU9ImNvbG9yOiMzZDRhNWMiPlsnICsgZXNjKGQu"
    b64 = b64 & "bmFtZSkgKyAnXTwvaDM+PHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC1zbSIgc3R5bGU9ImNvbG9yOiM1YTY0NzIiPm1vcmVfdmVydDwvc3Bhbj48L2hlYWRlcj4nICsNCiAgICAgICAgJzxkaXYgY2xhc3M9"
    b64 = b64 & "InAtNCBmbGV4IGZsZXgtY29sIGdhcC0wIj48dWwgY2xhc3M9InNwYWNlLXktMCI+JyArIGl0ZW1zICsgJzwvdWw+PC9kaXY+PC9hcnRpY2xlPic7DQogICAgfSkuam9pbignJyk7DQogICAgQXJyYXkuZnJvbShib2FyZC5xdWVyeVNlbGVj"
    b64 = b64 & "dG9yQWxsKCcuYm9hcmQtY2FyZCcpKS5mb3JFYWNoKGJpbmREcmFnKTsNCiAgfQ0KICBmdW5jdGlvbiBzZXRGaWx0ZXIoZmlsdGVyKSB7DQogICAgY3VycmVudEZpbHRlciA9IGZpbHRlcjsNCiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9y"
    b64 = b64 & "QWxsKCcuZmlsdGVyLWJ0bicpLmZvckVhY2goYnRuID0+IHsNCiAgICAgIGNvbnN0IGFjdGl2ZSA9IGJ0bi5nZXRBdHRyaWJ1dGUoJ2RhdGEtZmlsdGVyJykgPT09IGZpbHRlcjsNCiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCdmaWx0"
    b64 = b64 & "ZXItYWN0aXZlJywgYWN0aXZlKTsNCiAgICAgIGJ0bi5jbGFzc0xpc3QudG9nZ2xlKCd0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCcsICFhY3RpdmUpOw0KICAgIH0pOw0KICAgIHJlbmRlcigpOw0KICB9DQogIGZ1bmN0aW9uIGxvYWQoKSB7"
    b64 = b64 & "DQogICAgY29uc3QgcmF3ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25vdGljZS1kYXRhJykudGV4dENvbnRlbnQudHJpbSgpOw0KICAgIGxldCBkYXRhID0geyBjb2xsZWN0ZWRBdDogJycsIGRlcGFydG1lbnRzOiBbXSB9Ow0KICAg"
    b64 = b64 & "IHRyeSB7IGlmIChyYXcgJiYgcmF3LmNoYXJBdCgwKSA9PT0gJ3snKSBkYXRhID0gSlNPTi5wYXJzZShyYXcpOyB9IGNhdGNoIChlKSB7IGNvbnNvbGUuZXJyb3IoZSk7IH0NCiAgICBhbGxEZXB0cyA9IEFycmF5LmlzQXJyYXkoZGF0YS5k"
    b64 = b64 & "ZXBhcnRtZW50cykgPyBkYXRhLmRlcGFydG1lbnRzLnNsaWNlKCkgOiBbXTsNCiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWV0YUxpbmUnKS5pbm5lckhUTUwgPQ0KICAgICAgJ+yImOynkeyLnOqwgTogPHNwYW4gY2xhc3M9ImZv"
    b64 = b64 & "bnQtYm9sZCI+JyArIGVzYyhkYXRhLmNvbGxlY3RlZEF0IHx8ICctJykgKyAnPC9zcGFuPiB8IOu2gOyymCA8c3BhbiBjbGFzcz0iZm9udC1ib2xkIj4nICsgYWxsRGVwdHMubGVuZ3RoICsgJ+qwnDwvc3Bhbj4gfCDsubTrk5wg65Oc656Y"
    b64 = b64 & "6re466GcIOychOy5mCDsnbTrj5kg6rCA64qlJzsNCiAgICByZW5kZXIoKTsNCiAgfQ0KICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuZmlsdGVyLWJ0bicpLmZvckVhY2goYnRuID0+IHsNCiAgICBidG4uYWRkRXZlbnRMaXN0ZW5l"
    b64 = b64 & "cignY2xpY2snLCBmdW5jdGlvbiAoKSB7IHNldEZpbHRlcihidG4uZ2V0QXR0cmlidXRlKCdkYXRhLWZpbHRlcicpKTsgfSk7DQogIH0pOw0KICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnRuUmVsb2FkJykuYWRkRXZlbnRMaXN0ZW5l"
    b64 = b64 & "cignY2xpY2snLCBmdW5jdGlvbiAoKSB7IGxvY2F0aW9uLnJlbG9hZCgpOyB9KTsNCiAgbG9hZCgpOw0KfSkoKTsNCjwvc2NyaXB0Pg0KPC9ib2R5Pg0KPC9odG1sPg0K"
    EmbeddedHtmlTemplatePc = DecodeBase64Utf8(b64)
End Function

Private Function EmbeddedHtmlTemplateMobile() As String
    Dim b64 As String
    b64 = ""
    b64 = b64 & "77u/PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0x"
    b64 = b64 & "LjAiIG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KCV67aA67aA7LKYIOqzteyngOyCrO2VrSDrs7Trk5wgKOuqqOuwlOydvCk8L3RpdGxlPg0KPHNjcmlwdCBzcmM9Imh0dHBzOi8vY2RuLnRhaWx3aW5kY3NzLmNvbT9wbHVnaW5zPWZv"
    b64 = b64 & "cm1zLGNvbnRhaW5lci1xdWVyaWVzIj48L3NjcmlwdD4NCjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9UHVibGljK1NhbnM6d2dodEA0MDA7NTAwOzYwMDs3MDA7ODAwJmFtcDtkaXNwbGF5"
    b64 = b64 & "PXN3YXAiIHJlbD0ic3R5bGVzaGVldCIvPg0KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1NYXRlcmlhbCtTeW1ib2xzK091dGxpbmVkOndnaHQsRklMTEAxMDAuLjcwMCwwLi4xJmFtcDtk"
    b64 = b64 & "aXNwbGF5PXN3YXAiIHJlbD0ic3R5bGVzaGVldCIvPg0KPHNjcmlwdCBpZD0idGFpbHdpbmQtY29uZmlnIj4NCiAgICAgIHRhaWx3aW5kLmNvbmZpZyA9IHsNCiAgICAgICAgZGFya01vZGU6ICJjbGFzcyIsDQogICAgICAgIHRoZW1lOiB7"
    b64 = b64 & "DQogICAgICAgICAgZXh0ZW5kOiB7DQogICAgICAgICAgICAiY29sb3JzIjogew0KICAgICAgICAgICAgICAgICAgICAib24tdGVydGlhcnktZml4ZWQtdmFyaWFudCI6ICIjN2E1YTRjIiwNCiAgICAgICAgICAgICAgICAgICAgIm91dGxp"
    b64 = b64 & "bmUtdmFyaWFudCI6ICIjZDVkYWUyIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXRlcnRpYXJ5IjogIiNmZmZmZmYiLA0KICAgICAgICAgICAgICAgICAgICAiYmFja2dyb3VuZCI6ICIjZjdmOGZhIiwNCiAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICAgIm9uLXByaW1hcnktY29udGFpbmVyIjogIiMyZjQwNTUiLA0KICAgICAgICAgICAgICAgICAgICAib24tZXJyb3IiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICJzZWNvbmRhcnkiOiAiI2M5OTU2ZSIsDQogICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICJ0ZXJ0aWFyeS1maXhlZCI6ICIjZjBlMGQ4IiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtY29udGFpbmVyIjogIiNlZWYxZjUiLA0KICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeSI6ICIjZmZm"
    b64 = b64 & "ZmZmIiwNCiAgICAgICAgICAgICAgICAgICAgImludmVyc2Utc3VyZmFjZSI6ICIjM2E0MjUwIiwNCiAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5IjogIiNiMDg5NzgiLA0KICAgICAgICAgICAgICAgICAgICAic3VyZmFjZS1jb250"
    b64 = b64 & "YWluZXItaGlnaCI6ICIjZWNlZmYzIiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtZGltIjogIiNlNGU4ZWUiLA0KICAgICAgICAgICAgICAgICAgICAic3VyZmFjZSI6ICIjZjdmOGZhIiwNCiAgICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "InByaW1hcnktZml4ZWQtZGltIjogIiNjNWQ2ZTgiLA0KICAgICAgICAgICAgICAgICAgICAicHJpbWFyeS1maXhlZCI6ICIjZGNlN2YyIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeS1maXhlZCI6ICIjNGEzODJlIiwN"
    b64 = b64 & "CiAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWZpeGVkLWRpbSI6ICIjZTBjOGJjIiwNCiAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1maXhlZC1kaW0iOiAiI2U4YzliMCIsDQogICAgICAgICAgICAgICAgICAgICJwcmlt"
    b64 = b64 & "YXJ5LWNvbnRhaW5lciI6ICIjYjdjY2UwIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeSI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgICAgICAgInRlcnRpYXJ5LWNvbnRhaW5lciI6ICIjZDRiNWE1IiwNCiAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeS1jb250YWluZXIiOiAiIzVhNDMzNiIsDQogICAgICAgICAgICAgICAgICAgICJzZWNvbmRhcnktY29udGFpbmVyIjogIiNmMGQ1YzAiLA0KICAgICAgICAgICAgICAgICAgICAib24tc2Vj"
    b64 = b64 & "b25kYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzhhNmE1MiIsDQogICAgICAgICAgICAgICAgICAgICJlcnJvci1jb250YWluZXIiOiAiI2Y1ZDVkMiIsDQogICAgICAgICAgICAgICAgICAgICJvbi10ZXJ0aWFyeS1jb250YWluZXIiOiAiI2Zm"
    b64 = b64 & "ZjhmNSIsDQogICAgICAgICAgICAgICAgICAgICJvbi1wcmltYXJ5LWZpeGVkIjogIiMyZjQwNTUiLA0KICAgICAgICAgICAgICAgICAgICAib24tc3VyZmFjZS12YXJpYW50IjogIiM1YTY0NzIiLA0KICAgICAgICAgICAgICAgICAgICAi"
    b64 = b64 & "c3VyZmFjZS1jb250YWluZXItaGlnaGVzdCI6ICIjZTRlOGVlIiwNCiAgICAgICAgICAgICAgICAgICAgInN1cmZhY2UtdmFyaWFudCI6ICIjZTRlOGVlIiwNCiAgICAgICAgICAgICAgICAgICAgInNlY29uZGFyeS1maXhlZCI6ICIjZjVl"
    b64 = b64 & "NmQ4IiwNCiAgICAgICAgICAgICAgICAgICAgIm91dGxpbmUiOiAiIzlhYTNiMCIsDQogICAgICAgICAgICAgICAgICAgICJvbi1zdXJmYWNlIjogIiMyYTMzNDAiLA0KICAgICAgICAgICAgICAgICAgICAiaW52ZXJzZS1vbi1zdXJmYWNl"
    b64 = b64 & "IjogIiNmNGY2ZjgiLA0KICAgICAgICAgICAgICAgICAgICAib24tcHJpbWFyeS1maXhlZC12YXJpYW50IjogIiM0YTY3ODUiLA0KICAgICAgICAgICAgICAgICAgICAiZXJyb3IiOiAiI2MwNzA2YyIsDQogICAgICAgICAgICAgICAgICAg"
    b64 = b64 & "ICJzdXJmYWNlLWNvbnRhaW5lci1sb3dlc3QiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLWJyaWdodCI6ICIjZjdmOGZhIiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLXRlcnRpYXJ5LWZpeGVkIjogIiM0"
    b64 = b64 & "YTM0MmMiLA0KICAgICAgICAgICAgICAgICAgICAicHJpbWFyeSI6ICIjNmI4ZmI4IiwNCiAgICAgICAgICAgICAgICAgICAgIm9uLWJhY2tncm91bmQiOiAiIzJhMzM0MCIsDQogICAgICAgICAgICAgICAgICAgICJpbnZlcnNlLXByaW1h"
    b64 = b64 & "cnkiOiAiI2M1ZDZlOCIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLWNvbnRhaW5lci1sb3ciOiAiI2Y0ZjZmOCIsDQogICAgICAgICAgICAgICAgICAgICJzdXJmYWNlLXRpbnQiOiAiIzZiOGZiOCIsDQogICAgICAgICAgICAg"
    b64 = b64 & "ICAgICAgICJvbi1lcnJvci1jb250YWluZXIiOiAiIzdhMmUyYSINCiAgICAgICAgICAgIH0sDQogICAgICAgICAgICAiYm9yZGVyUmFkaXVzIjogew0KICAgICAgICAgICAgICAgICAgICAiREVGQVVMVCI6ICIwLjI1cmVtIiwNCiAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgImxnIjogIjAuNXJlbSIsDQogICAgICAgICAgICAgICAgICAgICJ4bCI6ICIwLjc1cmVtIiwNCiAgICAgICAgICAgICAgICAgICAgImZ1bGwiOiAiOTk5OXB4Ig0KICAgICAgICAgICAgfSwNCiAgICAgICAgICAg"
    b64 = b64 & "ICJzcGFjaW5nIjogew0KICAgICAgICAgICAgICAgICAgICAiZ3JpZC1ndXR0ZXIiOiAiMS4yNXJlbSIsDQogICAgICAgICAgICAgICAgICAgICJjb250YWluZXItcGFkZGluZyI6ICIycmVtIiwNCiAgICAgICAgICAgICAgICAgICAgInN0"
    b64 = b64 & "YWNrLWdhcCI6ICIwLjVyZW0iLA0KICAgICAgICAgICAgICAgICAgICAiY2FyZC1wYWRkaW5nIjogIjFyZW0iDQogICAgICAgICAgICB9LA0KICAgICAgICAgICAgImZvbnRGYW1pbHkiOiB7DQogICAgICAgICAgICAgICAgICAgICJsaXN0"
    b64 = b64 & "LWl0ZW0iOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICJtZXRhLWRhdGEiOiBbIlB1YmxpYyBTYW5zIl0sDQogICAgICAgICAgICAgICAgICAgICJoZWFkZXItdGl0bGUiOiBbIlB1YmxpYyBTYW5zIl0sDQogICAg"
    b64 = b64 & "ICAgICAgICAgICAgICAgICJidXR0b24tdGV4dCI6IFsiUHVibGljIFNhbnMiXSwNCiAgICAgICAgICAgICAgICAgICAgImJvYXJkLXRpdGxlIjogWyJQdWJsaWMgU2FucyJdDQogICAgICAgICAgICB9LA0KICAgICAgICAgICAgImZvbnRT"
    b64 = b64 & "aXplIjogew0KICAgICAgICAgICAgICAgICAgICAibGlzdC1pdGVtIjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdodCI6ICI1MDAifV0sDQogICAgICAgICAgICAgICAgICAgICJtZXRhLWRhdGEiOiBbIjEy"
    b64 = b64 & "cHgiLCB7ImxpbmVIZWlnaHQiOiAiMTZweCIsICJmb250V2VpZ2h0IjogIjQwMCJ9XSwNCiAgICAgICAgICAgICAgICAgICAgImhlYWRlci10aXRsZSI6IFsiMjRweCIsIHsibGluZUhlaWdodCI6ICIzMnB4IiwgImxldHRlclNwYWNpbmci"
    b64 = b64 & "OiAiLTAuMDJlbSIsICJmb250V2VpZ2h0IjogIjcwMCJ9XSwNCiAgICAgICAgICAgICAgICAgICAgImJ1dHRvbi10ZXh0IjogWyIxNHB4IiwgeyJsaW5lSGVpZ2h0IjogIjIwcHgiLCAiZm9udFdlaWdodCI6ICI2MDAifV0sDQogICAgICAg"
    b64 = b64 & "ICAgICAgICAgICAgICJib2FyZC10aXRsZSI6IFsiMTZweCIsIHsibGluZUhlaWdodCI6ICIyNHB4IiwgImZvbnRXZWlnaHQiOiAiNzAwIn1dDQogICAgICAgICAgICB9DQogICAgICAgICAgfSwNCiAgICAgICAgfSwNCiAgICAgIH0NCiAg"
    b64 = b64 & "ICA8L3NjcmlwdD4NCjxzdHlsZT4NCiAgICAgICAgYm9keSB7DQogICAgICAgICAgICBiYWNrZ3JvdW5kLWNvbG9yOiAjZjdmOGZhOw0KICAgICAgICAgICAgLXdlYmtpdC10YXAtaGlnaGxpZ2h0LWNvbG9yOiB0cmFuc3BhcmVudDsNCiAg"
    b64 = b64 & "ICAgICAgfQ0KICAgICAgICAubWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB7DQogICAgICAgICAgICBmb250LXZhcmlhdGlvbi1zZXR0aW5nczogJ0ZJTEwnIDAsICd3Z2h0JyA0MDAsICdHUkFEJyAwLCAnb3BzeicgMjQ7DQogICAgICAg"
    b64 = b64 & "IH0NCiAgICAgICAgLmhpZGUtc2Nyb2xsYmFyOjotd2Via2l0LXNjcm9sbGJhciB7IGRpc3BsYXk6IG5vbmU7IH0NCiAgICAgICAgLmhpZGUtc2Nyb2xsYmFyIHsgLW1zLW92ZXJmbG93LXN0eWxlOiBub25lOyBzY3JvbGxiYXItd2lkdGg6"
    b64 = b64 & "IG5vbmU7IH0NCiAgICAgICAgLm5vdGljZS1jYXJkIHsgYm94LXNoYWRvdzogMHB4IDRweCAxMnB4IHJnYmEoMCwwLDAsMC4wNSk7IH0NCiAgICAgICAgLnRhYi1hY3RpdmUgew0KICAgICAgICAgICAgYmFja2dyb3VuZC1jb2xvcjogI2Yw"
    b64 = b64 & "ZDVjMCAhaW1wb3J0YW50Ow0KICAgICAgICAgICAgY29sb3I6ICM1YTQzMzYgIWltcG9ydGFudDsNCiAgICAgICAgfQ0KPC9zdHlsZT4NCjwvaGVhZD4NCjxib2R5IGNsYXNzPSJmbGV4IGZsZXgtY29sIG1pbi1oLXNjcmVlbiI+DQo8aGVh"
    b64 = b64 & "ZGVyIGNsYXNzPSJiZy1wcmltYXJ5IHRleHQtb24tcHJpbWFyeSBmaXhlZCB0b3AtMCBsZWZ0LTAgcmlnaHQtMCB6LTUwIHJvdW5kZWQtYi14bCBzaGFkb3ctbWQgZmxleCBmbGV4LWNvbCBpdGVtcy1jZW50ZXIgcHgtNCBweS00IHctZnVs"
    b64 = b64 & "bCI+DQo8ZGl2IGNsYXNzPSJmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgdy1mdWxsIj4NCjxoMSBjbGFzcz0iZm9udC1oZWFkZXItdGl0bGUgdGV4dC1oZWFkZXItdGl0bGUgdGV4dC1vbi1wcmltYXJ5Ij7soJXrtoDrtoDs"
    b64 = b64 & "spgg6rO17KeA7IKs7ZWtPC9oMT4NCjxidXR0b24gaWQ9ImJ0blJlbG9hZCIgdHlwZT0iYnV0dG9uIiBjbGFzcz0iZm9udC1idXR0b24tdGV4dCB0ZXh0LWJ1dHRvbi10ZXh0IGZsZXggaXRlbXMtY2VudGVyIGdhcC0xIGhvdmVyOmJnLXBy"
    b64 = b64 & "aW1hcnktY29udGFpbmVyLzIwIHRyYW5zaXRpb24tY29sb3JzIHAtMiByb3VuZGVkLWxnIGFjdGl2ZTpzY2FsZS05NSBkdXJhdGlvbi0xNTAiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnJlZnJlc2g8L3Nw"
    b64 = b64 & "YW4+DQo8c3BhbiBjbGFzcz0iaGlkZGVuIHNtOmlubGluZSI+7ZmU66m0IOyDiOuhnOqzoOy5qDwvc3Bhbj4NCjwvYnV0dG9uPg0KPC9kaXY+DQo8ZGl2IGNsYXNzPSJ3LWZ1bGwgbXQtNCBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWJl"
    b64 = b64 & "dHdlZW4gZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgb3BhY2l0eS05MCBib3JkZXItdCBib3JkZXItb24tcHJpbWFyeS8xMCBwdC0zIj4NCjxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIj4NCjxzcGFuIGNsYXNz"
    b64 = b64 & "PSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE0cHhdIj51cGRhdGU8L3NwYW4+DQo8c3BhbiBpZD0ibWV0YVRpbWUiPuuniOyngOuniSDsl4XrjbDsnbTtirg6IC08L3NwYW4+DQo8L2Rpdj4NCjxkaXYgaWQ9Im1ldGFDb3Vu"
    b64 = b64 & "dCI+67aA7LKYIOy0nTogMOqwnDwvZGl2Pg0KPC9kaXY+DQo8L2hlYWRlcj4NCg0KPG5hdiBjbGFzcz0iZml4ZWQgdG9wLVsxMTBweF0gbGVmdC0wIHJpZ2h0LTAgYmctc3VyZmFjZS1jb250YWluZXIgei00MCBib3JkZXItYiBib3JkZXIt"
    b64 = b64 & "b3V0bGluZS12YXJpYW50IHNoYWRvdy1zbSI+DQo8ZGl2IGNsYXNzPSJmbGV4IG92ZXJmbG93LXgtYXV0byBoaWRlLXNjcm9sbGJhciBweC00IHB5LTMgZ2FwLTIiIGlkPSJmaWx0ZXJOYXYiPg0KPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRh"
    b64 = b64 & "dGEtZmlsdGVyPSJhbGwiIGNsYXNzPSJ0YWItYnRuIHRhYi1hY3RpdmUgZmxleC1zaHJpbmstMCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBweC00IHB5LTIgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cmFu"
    b64 = b64 & "c2l0aW9uLWFsbCBkdXJhdGlvbi0yMDAgZWFzZS1pbi1vdXQiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPmRhc2hib2FyZDwvc3Bhbj7soITssrQg67aA7LKYDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0i"
    b64 = b64 & "YnV0dG9uIiBkYXRhLWZpbHRlcj0ibWFqb3IiIGNsYXNzPSJ0YWItYnRuIGZsZXgtc2hyaW5rLTAgZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgcHgtNCBweS0yIHRleHQtb24tc3VyZmFjZS12YXJpYW50IGhvdmVyOmJnLXN1cmZhY2UtdmFy"
    b64 = b64 & "aWFudCByb3VuZGVkLWxnIGZvbnQtbGlzdC1pdGVtIHRleHQtbGlzdC1pdGVtIHRyYW5zaXRpb24tYWxsIGR1cmF0aW9uLTIwMCBlYXNlLWluLW91dCI+DQo8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCI+c3Rhcjwv"
    b64 = b64 & "c3Bhbj7so7zsmpTrtoDsspgNCjwvYnV0dG9uPg0KPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtZmlsdGVyPSJlY29ub215IiBjbGFzcz0idGFiLWJ0biBmbGV4LXNocmluay0wIGZsZXggaXRlbXMtY2VudGVyIGdhcC0yIHB4LTQgcHkt"
    b64 = b64 & "MiB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQgcm91bmRlZC1sZyBmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0cmFuc2l0aW9uLWFsbCBkdXJhdGlvbi0yMDAgZWFzZS1pbi1vdXQiPg0K"
    b64 = b64 & "PHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPmJ1c2luZXNzX2NlbnRlcjwvc3Bhbj7sgrDsl4Uv6rK97KCcDQo8L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0ic29jaWV0eSIgY2xh"
    b64 = b64 & "c3M9InRhYi1idG4gZmxleC1zaHJpbmstMCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBweC00IHB5LTIgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50IHJvdW5kZWQtbGcgZm9udC1saXN0LWl0ZW0g"
    b64 = b64 & "dGV4dC1saXN0LWl0ZW0gdHJhbnNpdGlvbi1hbGwgZHVyYXRpb24tMjAwIGVhc2UtaW4tb3V0Ij4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj50aGVhdGVyX2NvbWVkeTwvc3Bhbj7sgqztmowv66y47ZmUDQo8"
    b64 = b64 & "L2J1dHRvbj4NCjxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLWZpbHRlcj0iYWRtaW4iIGNsYXNzPSJ0YWItYnRuIGZsZXgtc2hyaW5rLTAgZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgcHgtNCBweS0yIHRleHQtb24tc3VyZmFjZS12YXJp"
    b64 = b64 & "YW50IGhvdmVyOmJnLXN1cmZhY2UtdmFyaWFudCByb3VuZGVkLWxnIGZvbnQtbGlzdC1pdGVtIHRleHQtbGlzdC1pdGVtIHRyYW5zaXRpb24tYWxsIGR1cmF0aW9uLTIwMCBlYXNlLWluLW91dCI+DQo8c3BhbiBjbGFzcz0ibWF0ZXJpYWwt"
    b64 = b64 & "c3ltYm9scy1vdXRsaW5lZCI+YWNjb3VudF9iYWxhbmNlPC9zcGFuPu2WieyglS/slYjsoIQNCjwvYnV0dG9uPg0KPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtZmlsdGVyPSJkaXBsb21hY3kiIGNsYXNzPSJ0YWItYnRuIGZsZXgtc2hy"
    b64 = b64 & "aW5rLTAgZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgcHgtNCBweS0yIHRleHQtb24tc3VyZmFjZS12YXJpYW50IGhvdmVyOmJnLXN1cmZhY2UtdmFyaWFudCByb3VuZGVkLWxnIGZvbnQtbGlzdC1pdGVtIHRleHQtbGlzdC1pdGVtIHRyYW5z"
    b64 = b64 & "aXRpb24tYWxsIGR1cmF0aW9uLTIwMCBlYXNlLWluLW91dCI+DQo8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCI+cHVibGljPC9zcGFuPuyZuOq1kC/slYjrs7QNCjwvYnV0dG9uPg0KPC9kaXY+DQo8L25hdj4NCg0K"
    b64 = b64 & "PGRpdiBjbGFzcz0ibXQtWzE4MHB4XSBweC00IG1heC13LWxnIG14LWF1dG8gdy1mdWxsIj4NCjxkaXYgaWQ9Im1ham9yQWN0aW9ucyIgY2xhc3M9ImhpZGRlbiBtYi0zIj4NCjxhIGlkPSJidG5QcmVzc1RvZ2V0aGVyIiBocmVmPSJnb3Zf"
    b64 = b64 & "bWFqb3JfcHJlc3NfYm9hcmRfbW9iaWxlLmh0bWwiIGNsYXNzPSJ3LWZ1bGwgZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgZ2FwLTIgYmctc2Vjb25kYXJ5IHRleHQtb24tc2Vjb25kYXJ5IGZvbnQtYnV0dG9uLXRleHQgdGV4"
    b64 = b64 & "dC1idXR0b24tdGV4dCBweC00IHB5LTMgcm91bmRlZC14bCBzaGFkb3cgbm8tdW5kZXJsaW5lIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIj5uZXdzcGFwZXI8L3NwYW4+67O064+E7J6Q66OMIO2VqOq7mOuz"
    b64 = b64 & "tOq4sA0KPC9hPg0KPC9kaXY+DQo8L2Rpdj4NCjxtYWluIGNsYXNzPSJtYi0yMCBweC00IGZsZXggZmxleC1jb2wgZ2FwLTUgbWF4LXctbGcgbXgtYXV0byB3LWZ1bGwiIGlkPSJib2FyZCI+PC9tYWluPg0KDQo8YnV0dG9uIGlkPSJmYWJT"
    b64 = b64 & "eW5jIiB0eXBlPSJidXR0b24iIGNsYXNzPSJmaXhlZCBib3R0b20tNiByaWdodC02IHctMTQgaC0xNCBiZy1wcmltYXJ5LWNvbnRhaW5lciB0ZXh0LW9uLXByaW1hcnktY29udGFpbmVyIHJvdW5kZWQtZnVsbCBzaGFkb3ctbGcgZmxleCBp"
    b64 = b64 & "dGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgaG92ZXI6c2NhbGUtMTA1IGFjdGl2ZTpzY2FsZS05NSB0cmFuc2l0aW9uLWFsbCB6LTUwIj4NCjxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzMycHhdIj5z"
    b64 = b64 & "eW5jPC9zcGFuPg0KPC9idXR0b24+DQoNCjxmb290ZXIgY2xhc3M9ImJnLXN1cmZhY2UtZGltIHctZnVsbCBweS02IHB4LTQgZmxleCBmbGV4LWNvbCBpdGVtcy1jZW50ZXIgZ2FwLTQgdGV4dC1jZW50ZXIgYm9yZGVyLXQgYm9yZGVyLW91"
    b64 = b64 & "dGxpbmUtdmFyaWFudCBtdC1hdXRvIj4NCjxkaXYgY2xhc3M9ImZvbnQtYm9sZCB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBtYi0xIj7rjIDtlZzrr7zqta0g7KCV67aAPC9kaXY+DQo8ZGl2IGNsYXNzPSJmbGV4IGZsZXgtd3JhcCBqdXN0"
    b64 = b64 & "aWZ5LWNlbnRlciBnYXAtNCBmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBvcGFjaXR5LTgwIj4NCjxhIGNsYXNzPSJob3Zlcjp0ZXh0LXByaW1hcnkgbm8tdW5kZXJsaW5lIHRyYW5zaXRp"
    b64 = b64 & "b24tY29sb3JzIiBocmVmPSJnb3Zfbm90aWNlX2JvYXJkLmh0bWwiPlBDIO2OmOydtOyngDwvYT4NCjxhIGNsYXNzPSJob3Zlcjp0ZXh0LXByaW1hcnkgbm8tdW5kZXJsaW5lIHRyYW5zaXRpb24tY29sb3JzIiBocmVmPSJnb3Zfbm90aWNl"
    b64 = b64 & "X2JvYXJkX21vYmlsZS5odG1sIj7rqqjrsJTsnbw8L2E+DQo8L2Rpdj4NCjxwIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCBvcGFjaXR5LTYwIj7CqSDrjIDtlZzrr7zqta0g"
    b64 = b64 & "7KCV67aAIOu2gOyymCDqs7Xsp4Dsgqztla0g67O065OcIOyEnOu5hOyKpDwvcD4NCjwvZm9vdGVyPg0KDQo8c2NyaXB0IGlkPSJub3RpY2UtZGF0YSIgdHlwZT0iYXBwbGljYXRpb24vanNvbiI+DQolJU5PVElDRV9KU09OJSUNCjwvc2Ny"
    b64 = b64 & "aXB0Pg0KPHNjcmlwdD4NCihmdW5jdGlvbiAoKSB7DQogIGNvbnN0IFNUT1JBR0VfS0VZID0gJ2dvdk5vdGljZUNhcmRPcmRlci52MSc7DQogIGNvbnN0IFBBTEVUVEUgPSBbJyNiN2QwZWEnLCcjZjBkMGIwJywnI2UwYzhiYycsJyNjNWM4"
    b64 = b64 & "ZTgnLCcjYjhkZGQ2JywnI2M1ZTBjOCcsJyNmMGM0YzQnLCcjYjhkY2UwJywnI2Q0YzRlOCcsJyNiMGM0ZTAnXTsNCiAgY29uc3QgSUNPTlMgPSB7DQogICAgJ+uztOqxtOuzteyngOu2gCc6J21lZGljYWxfc2VydmljZXMnLCftlonsoJXs"
    b64 = b64 & "lYjsoITrtoAnOidzZWN1cml0eScsJ+yZuOq1kOu2gCc6J3B1YmxpYycsJ+q1reuwqeu2gCc6J3NlY3VyaXR5JywNCiAgICAn6rWQ7Jyh67aAJzonc2Nob29sJywn6rOg7Jqp64W464+Z67aAJzond29yaycsJ+usuO2ZlOyytOycoeq0gOq0"
    b64 = b64 & "keu2gCc6J3RoZWF0ZXJfY29tZWR5JywnZGVmYXVsdCc6J3N0YXJzJw0KICB9Ow0KICBjb25zdCBDQVRFR09SWSA9IHsNCiAgICAn7J6s7KCV6rK97KCc67aAJzonZWNvbm9teScsJ+qzvO2Vmeq4sOyIoOygleuztO2GteyLoOu2gCc6J2Vj"
    b64 = b64 & "b25vbXknLCfsgrDsl4XthrXsg4HrtoAnOidlY29ub215Jywn7KSR7IaM67Kk7LKY6riw7JeF67aAJzonZWNvbm9teScsDQogICAgJ+q1re2GoOq1kO2Gteu2gCc6J2Vjb25vbXknLCfrho3rprzstpXsgrDsi53tkojrtoAnOidlY29ub215"
    b64 = b64 & "Jywn7ZW07JaR7IiY7IKw67aAJzonZWNvbm9teScsJ+q4sO2bhOyXkOuEiOyngO2ZmOqyveu2gCc6J2Vjb25vbXknLCfquLDtmo3smIjsgrDsspgnOidlY29ub215JywNCiAgICAn6rWQ7Jyh67aAJzonc29jaWV0eScsJ+uztOqxtOuzteyn"
    b64 = b64 & "gOu2gCc6J3NvY2lldHknLCfqs6Dsmqnrhbjrj5nrtoAnOidzb2NpZXR5Jywn66y47ZmU7LK07Jyh6rSA6rSR67aAJzonc29jaWV0eScsJ+yEse2PieuTseqwgOyhseu2gCc6J3NvY2lldHknLCfsi53tkojsnZjslb3tkojslYjsoITsspgn"
    b64 = b64 & "Oidzb2NpZXR5JywNCiAgICAn7ZaJ7KCV7JWI7KCE67aAJzonYWRtaW4nLCfrspXrrLTrtoAnOidhZG1pbicsJ+yduOyCrO2YgeyLoOyymCc6J2FkbWluJywn67KV7KCc7LKYJzonYWRtaW4nLCfqta3qsIDrs7Ttm4jrtoAnOidhZG1pbics"
    b64 = b64 & "DQogICAgJ+yZuOq1kOu2gCc6J2RpcGxvbWFjeScsJ+2GteydvOu2gCc6J2RpcGxvbWFjeScsJ+q1reuwqeu2gCc6J2RpcGxvbWFjeScNCiAgfTsNCiAgY29uc3QgTUFKT1JfT1JERVIgPSBbJ+yCsOyXhe2GteyDgeu2gCcsJ+q4sO2bhOyX"
    b64 = b64 & "kOuEiOyngO2ZmOqyveu2gCcsJ+qzoOyaqeuFuOuPmeu2gCcsJ+yZuOq1kOu2gCddOw0KICBsZXQgY3VycmVudEZpbHRlciA9ICdhbGwnOw0KICBsZXQgYWxsRGVwdHMgPSBbXTsNCg0KICBmdW5jdGlvbiBjb2xvckZvcihuYW1lKSB7DQog"
    b64 = b64 & "ICAgaWYgKG5hbWUgPT09ICfsgrDsl4XthrXsg4HrtoAnKSByZXR1cm4gJyNmMGQwYjAnOw0KICAgIGxldCBpZHggPSAwOw0KICAgIGZvciAobGV0IGkgPSAwOyBpIDwgbmFtZS5sZW5ndGg7IGkrKykgaWR4ID0gKGlkeCArIG5hbWUuY2hh"
    b64 = b64 & "ckNvZGVBdChpKSkgJSBQQUxFVFRFLmxlbmd0aDsNCiAgICByZXR1cm4gUEFMRVRURVtpZHhdOw0KICB9DQogIGZ1bmN0aW9uIGNhdGVnb3J5Rm9yKG5hbWUpIHsgcmV0dXJuIENBVEVHT1JZW25hbWVdIHx8ICdhZG1pbic7IH0NCiAgZnVu"
    b64 = b64 & "Y3Rpb24gaWNvbkZvcihuYW1lKSB7IHJldHVybiBJQ09OU1tuYW1lXSB8fCBJQ09OUy5kZWZhdWx0OyB9DQogIGZ1bmN0aW9uIGVzYyhzKSB7DQogICAgcmV0dXJuIFN0cmluZyhzID8/ICcnKS5yZXBsYWNlKC8mL2csJyZhbXA7JykucmVw"
    b64 = b64 & "bGFjZSgvPC9nLCcmbHQ7JykucmVwbGFjZSgvPi9nLCcmZ3Q7JykucmVwbGFjZSgvIi9nLCcmcXVvdDsnKTsNCiAgfQ0KICBmdW5jdGlvbiBhcHBseVNhdmVkT3JkZXIoZGVwdHMpIHsNCiAgICB0cnkgew0KICAgICAgY29uc3QgcmF3ID0g"
    b64 = b64 & "bG9jYWxTdG9yYWdlLmdldEl0ZW0oU1RPUkFHRV9LRVkpOw0KICAgICAgaWYgKCFyYXcpIHJldHVybiBkZXB0czsNCiAgICAgIGNvbnN0IG9yZGVyID0gSlNPTi5wYXJzZShyYXcpOw0KICAgICAgaWYgKCFBcnJheS5pc0FycmF5KG9yZGVy"
    b64 = b64 & "KSB8fCAhb3JkZXIubGVuZ3RoKSByZXR1cm4gZGVwdHM7DQogICAgICBjb25zdCBtYXAgPSB7fTsNCiAgICAgIGRlcHRzLmZvckVhY2goZCA9PiB7IG1hcFtkLm5hbWVdID0gZDsgfSk7DQogICAgICBjb25zdCBzb3J0ZWQgPSBbXTsNCiAg"
    b64 = b64 & "ICAgIG9yZGVyLmZvckVhY2gobmFtZSA9PiB7IGlmIChtYXBbbmFtZV0pIHsgc29ydGVkLnB1c2gobWFwW25hbWVdKTsgZGVsZXRlIG1hcFtuYW1lXTsgfSB9KTsNCiAgICAgIE9iamVjdC5rZXlzKG1hcCkuZm9yRWFjaChrID0+IHNvcnRl"
    b64 = b64 & "ZC5wdXNoKG1hcFtrXSkpOw0KICAgICAgcmV0dXJuIHNvcnRlZDsNCiAgICB9IGNhdGNoIChlKSB7IHJldHVybiBkZXB0czsgfQ0KICB9DQogIGZ1bmN0aW9uIHBhZE51bShuKSB7IHJldHVybiAobiA8IDEwID8gJzAnIDogJycpICsgbjsg"
    b64 = b64 & "fQ0KICBmdW5jdGlvbiByZW5kZXIoKSB7DQogICAgY29uc3QgYm9hcmQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYm9hcmQnKTsNCiAgICBjb25zdCBtYWpvckJveCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtYWpvckFjdGlv"
    b64 = b64 & "bnMnKTsNCiAgICBpZiAobWFqb3JCb3gpIG1ham9yQm94LmNsYXNzTGlzdC50b2dnbGUoJ2hpZGRlbicsIGN1cnJlbnRGaWx0ZXIgIT09ICdtYWpvcicpOw0KICAgIGxldCBkZXB0cyA9IGFsbERlcHRzLnNsaWNlKCk7DQogICAgaWYgKGN1"
    b64 = b64 & "cnJlbnRGaWx0ZXIgPT09ICdtYWpvcicpIHsNCiAgICAgIGNvbnN0IG1hcCA9IHt9Ow0KICAgICAgZGVwdHMuZm9yRWFjaChkID0+IHsgbWFwW2QubmFtZV0gPSBkOyB9KTsNCiAgICAgIGRlcHRzID0gTUFKT1JfT1JERVIubWFwKG4gPT4g"
    b64 = b64 & "bWFwW25dKS5maWx0ZXIoQm9vbGVhbik7DQogICAgfSBlbHNlIGlmIChjdXJyZW50RmlsdGVyICE9PSAnYWxsJykgew0KICAgICAgZGVwdHMgPSBkZXB0cy5maWx0ZXIoZCA9PiBjYXRlZ29yeUZvcihkLm5hbWUpID09PSBjdXJyZW50Rmls"
    b64 = b64 & "dGVyKTsNCiAgICAgIGRlcHRzID0gYXBwbHlTYXZlZE9yZGVyKGRlcHRzKTsNCiAgICB9IGVsc2Ugew0KICAgICAgZGVwdHMgPSBhcHBseVNhdmVkT3JkZXIoZGVwdHMpOw0KICAgIH0NCiAgICBpZiAoIWRlcHRzLmxlbmd0aCkgew0KICAg"
    b64 = b64 & "ICAgYm9hcmQuaW5uZXJIVE1MID0gJzxkaXYgY2xhc3M9InRleHQtY2VudGVyIHB5LTE2IHRleHQtb24tc3VyZmFjZS12YXJpYW50IGJnLXN1cmZhY2UtY29udGFpbmVyLWxvd2VzdCByb3VuZGVkLXhsIGJvcmRlciBib3JkZXItZGFzaGVk"
    b64 = b64 & "IGJvcmRlci1vdXRsaW5lLXZhcmlhbnQiPu2RnOyLnO2VoCDqs7Xsp4Ag642w7J207YSw6rCAIOyXhuyKteuLiOuLpC48L2Rpdj4nOw0KICAgICAgcmV0dXJuOw0KICAgIH0NCiAgICBib2FyZC5pbm5lckhUTUwgPSBkZXB0cy5tYXAoZnVu"
    b64 = b64 & "Y3Rpb24gKGQpIHsNCiAgICAgIGNvbnN0IGhlYWQgPSBjb2xvckZvcihkLm5hbWUpOw0KICAgICAgY29uc3QgaWNvbiA9IGljb25Gb3IoZC5uYW1lKTsNCiAgICAgIGlmIChkLm9rID09PSBmYWxzZSB8fCAhZC5wb3N0cyB8fCAhZC5wb3N0"
    b64 = b64 & "cy5sZW5ndGgpIHsNCiAgICAgICAgcmV0dXJuICc8YXJ0aWNsZSBjbGFzcz0ibm90aWNlLWNhcmQgYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0IHJvdW5kZWQteGwgb3ZlcmZsb3ctaGlkZGVuIGJvcmRlciBib3JkZXItb3V0bGluZS12"
    b64 = b64 & "YXJpYW50LzMwIj4nICsNCiAgICAgICAgICAnPGhlYWRlciBjbGFzcz0icHgtNCBweS0zIGZsZXgganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLWNlbnRlciIgc3R5bGU9ImJhY2tncm91bmQ6JyArIGhlYWQgKyAnIj4nICsNCiAgICAgICAgICAn"
    b64 = b64 & "PGgyIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUiIHN0eWxlPSJjb2xvcjojM2Q0YTVjIj4nICsgZXNjKGQubmFtZSkgKyAnPC9oMj4nICsNCiAgICAgICAgICAnPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJv"
    b64 = b64 & "bHMtb3V0bGluZWQiIHN0eWxlPSJjb2xvcjojNWE2NDcyIj4nICsgaWNvbiArICc8L3NwYW4+PC9oZWFkZXI+JyArDQogICAgICAgICAgJzxkaXYgY2xhc3M9InAtY2FyZC1wYWRkaW5nIHRleHQtZXJyb3IgZm9udC1ib2xkIHRleHQtc20i"
    b64 = b64 & "PuyEpOyglSDtjpjsnbTsp4Drpbwg7IiY7KCV7ZWY7IS47JqULjwvZGl2PjwvYXJ0aWNsZT4nOw0KICAgICAgfQ0KICAgICAgY29uc3QgaXRlbXMgPSBkLnBvc3RzLnNsaWNlKDAsIDUpLm1hcChmdW5jdGlvbiAocCwgaSkgew0KICAgICAg"
    b64 = b64 & "ICBjb25zdCBsYXN0ID0gaSA9PT0gTWF0aC5taW4oZC5wb3N0cy5sZW5ndGgsIDUpIC0gMTsNCiAgICAgICAgY29uc3QgYm9yZGVyID0gbGFzdCA/ICcnIDogJyBib3JkZXItYiBib3JkZXItb3V0bGluZS12YXJpYW50LzIwIHBiLTInOw0K"
    b64 = b64 & "ICAgICAgICByZXR1cm4gJzxsaSBjbGFzcz0iZmxleCBpdGVtcy1zdGFydCBnYXAtMyI+JyArDQogICAgICAgICAgJzxzcGFuIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSBwdC0xIiBzdHlsZT0iY29sb3I6IzZiOGZi"
    b64 = b64 & "OCI+JyArIHBhZE51bShpICsgMSkgKyAnPC9zcGFuPicgKw0KICAgICAgICAgICc8ZGl2IGNsYXNzPSJmbGV4IGZsZXgtY29sIHctZnVsbCcgKyBib3JkZXIgKyAnIj4nICsNCiAgICAgICAgICAnPGEgY2xhc3M9ImZvbnQtbGlzdC1pdGVt"
    b64 = b64 & "IHRleHQtbGlzdC1pdGVtIHRleHQtb24tc3VyZmFjZSBsaW5lLWNsYW1wLTIgbm8tdW5kZXJsaW5lIiBocmVmPSInICsgZXNjKHAudXJsKSArICciIHRhcmdldD0iX2JsYW5rIiByZWw9Im5vb3BlbmVyIG5vcmVmZXJyZXIiPicgKyBlc2Mo"
    b64 = b64 & "cC50aXRsZSkgKyAnPC9hPicgKw0KICAgICAgICAgICc8L2Rpdj48L2xpPic7DQogICAgICB9KS5qb2luKCcnKTsNCiAgICAgIHJldHVybiAnPGFydGljbGUgY2xhc3M9Im5vdGljZS1jYXJkIGJnLXN1cmZhY2UtY29udGFpbmVyLWxvd2Vz"
    b64 = b64 & "dCByb3VuZGVkLXhsIG92ZXJmbG93LWhpZGRlbiBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudC8zMCB0cmFuc2l0aW9uLXRyYW5zZm9ybSBhY3RpdmU6c2NhbGUtWzAuOThdIiBkYXRhLW5hbWU9IicgKyBlc2MoZC5uYW1lKSArICci"
    b64 = b64 & "PicgKw0KICAgICAgICAnPGhlYWRlciBjbGFzcz0icHgtNCBweS0zIGZsZXgganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLWNlbnRlciIgc3R5bGU9ImJhY2tncm91bmQ6JyArIGhlYWQgKyAnIj4nICsNCiAgICAgICAgJzxoMiBjbGFzcz0iZm9u"
    b64 = b64 & "dC1ib2FyZC10aXRsZSB0ZXh0LWJvYXJkLXRpdGxlIiBzdHlsZT0iY29sb3I6IzNkNGE1YyI+JyArIGVzYyhkLm5hbWUpICsgJzwvaDI+JyArDQogICAgICAgICc8c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCIgc3R5"
    b64 = b64 & "bGU9ImNvbG9yOiM1YTY0NzIiPicgKyBpY29uICsgJzwvc3Bhbj48L2hlYWRlcj4nICsNCiAgICAgICAgJzxkaXYgY2xhc3M9InAtY2FyZC1wYWRkaW5nIj48dWwgY2xhc3M9InNwYWNlLXktNCI+JyArIGl0ZW1zICsgJzwvdWw+PC9kaXY+"
    b64 = b64 & "PC9hcnRpY2xlPic7DQogICAgfSkuam9pbignJyk7DQogIH0NCiAgZnVuY3Rpb24gc2V0RmlsdGVyKGZpbHRlciwgYnRuKSB7DQogICAgY3VycmVudEZpbHRlciA9IGZpbHRlcjsNCiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcu"
    b64 = b64 & "dGFiLWJ0bicpLmZvckVhY2goYiA9PiB7DQogICAgICBiLmNsYXNzTGlzdC5yZW1vdmUoJ3RhYi1hY3RpdmUnLCAnYmctc2Vjb25kYXJ5LWNvbnRhaW5lcicsICd0ZXh0LW9uLXNlY29uZGFyeS1jb250YWluZXInKTsNCiAgICAgIGIuY2xh"
    b64 = b64 & "c3NMaXN0LmFkZCgndGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQnLCAnaG92ZXI6Ymctc3VyZmFjZS12YXJpYW50Jyk7DQogICAgfSk7DQogICAgaWYgKGJ0bikgew0KICAgICAgYnRuLmNsYXNzTGlzdC5hZGQoJ3RhYi1hY3RpdmUnKTsNCiAg"
    b64 = b64 & "ICAgIGJ0bi5jbGFzc0xpc3QucmVtb3ZlKCd0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCcsICdob3ZlcjpiZy1zdXJmYWNlLXZhcmlhbnQnKTsNCiAgICAgIGJ0bi5zY3JvbGxJbnRvVmlldyh7IGJlaGF2aW9yOiAnc21vb3RoJywgYmxvY2s6"
    b64 = b64 & "ICduZWFyZXN0JywgaW5saW5lOiAnY2VudGVyJyB9KTsNCiAgICB9DQogICAgcmVuZGVyKCk7DQogIH0NCiAgZnVuY3Rpb24gbG9hZCgpIHsNCiAgICBjb25zdCByYXcgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbm90aWNlLWRhdGEn"
    b64 = b64 & "KS50ZXh0Q29udGVudC50cmltKCk7DQogICAgbGV0IGRhdGEgPSB7IGNvbGxlY3RlZEF0OiAnJywgZGVwYXJ0bWVudHM6IFtdIH07DQogICAgdHJ5IHsgaWYgKHJhdyAmJiByYXcuY2hhckF0KDApID09PSAneycpIGRhdGEgPSBKU09OLnBh"
    b64 = b64 & "cnNlKHJhdyk7IH0gY2F0Y2ggKGUpIHsgY29uc29sZS5lcnJvcihlKTsgfQ0KICAgIGFsbERlcHRzID0gQXJyYXkuaXNBcnJheShkYXRhLmRlcGFydG1lbnRzKSA/IGRhdGEuZGVwYXJ0bWVudHMuc2xpY2UoKSA6IFtdOw0KICAgIGNvbnN0"
    b64 = b64 & "IHQgPSBkYXRhLmNvbGxlY3RlZEF0IHx8ICctJzsNCiAgICBjb25zdCB0aW1lT25seSA9IHQuaW5kZXhPZignICcpID4gMCA/IHQuc3BsaXQoJyAnKVsxXSA6IHQ7DQogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21ldGFUaW1lJyku"
    b64 = b64 & "dGV4dENvbnRlbnQgPSAn66eI7KeA66eJIOyXheuNsOydtO2KuDogJyArIHRpbWVPbmx5Ow0KICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtZXRhQ291bnQnKS50ZXh0Q29udGVudCA9ICfrtoDsspgg7LSdOiAnICsgYWxsRGVwdHMu"
    b64 = b64 & "bGVuZ3RoICsgJ+qwnCc7DQogICAgcmVuZGVyKCk7DQogIH0NCiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1idG4nKS5mb3JFYWNoKGJ0biA9PiB7DQogICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgZnVuY3Rp"
    b64 = b64 & "b24gKCkgeyBzZXRGaWx0ZXIoYnRuLmdldEF0dHJpYnV0ZSgnZGF0YS1maWx0ZXInKSwgYnRuKTsgfSk7DQogIH0pOw0KICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYnRuUmVsb2FkJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBm"
    b64 = b64 & "dW5jdGlvbiAoKSB7IGxvY2F0aW9uLnJlbG9hZCgpOyB9KTsNCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2ZhYlN5bmMnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGZ1bmN0aW9uICgpIHsNCiAgICBjb25zdCBpY29uID0gdGhp"
    b64 = b64 & "cy5xdWVyeVNlbGVjdG9yKCcubWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCcpOw0KICAgIGljb24uY2xhc3NMaXN0LmFkZCgnYW5pbWF0ZS1zcGluJyk7DQogICAgc2V0VGltZW91dChmdW5jdGlvbiAoKSB7DQogICAgICBpY29uLmNsYXNz"
    b64 = b64 & "TGlzdC5yZW1vdmUoJ2FuaW1hdGUtc3BpbicpOw0KICAgICAgd2luZG93LnNjcm9sbFRvKHsgdG9wOiAwLCBiZWhhdmlvcjogJ3Ntb290aCcgfSk7DQogICAgICBsb2NhdGlvbi5yZWxvYWQoKTsNCiAgICB9LCA4MDApOw0KICB9KTsNCiAg"
    b64 = b64 & "bG9hZCgpOw0KfSkoKTsNCjwvc2NyaXB0Pg0KPC9ib2R5Pg0KPC9odG1sPg0K"
    EmbeddedHtmlTemplateMobile = DecodeBase64Utf8(b64)
End Function

Private Function EmbeddedMajorPressTemplatePc() As String
    Dim b64 As String
    b64 = ""
    b64 = b64 & "77u/PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0x"
    b64 = b64 & "LjAiIG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KO87JqU67aA7LKYIOqzteyngMK367O064+E7J6Q66OMIO2VqOq7mOuztOq4sCAoUEMpPC90aXRsZT4NCjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi50YWlsd2luZGNzcy5jb20/cGx1"
    b64 = b64 & "Z2lucz1mb3Jtcyxjb250YWluZXItcXVlcmllcyI+PC9zY3JpcHQ+DQo8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PVB1YmxpYytTYW5zOndnaHRANDAwOzUwMDs2MDA7NzAwOzgwMCZhbXA7"
    b64 = b64 & "ZGlzcGxheT1zd2FwIiByZWw9InN0eWxlc2hlZXQiLz4NCjxsaW5rIGhyZWY9Imh0dHBzOi8vZm9udHMuZ29vZ2xlYXBpcy5jb20vY3NzMj9mYW1pbHk9TWF0ZXJpYWwrU3ltYm9scytPdXRsaW5lZDp3Z2h0LEZJTExAMTAwLi43MDAsMC4u"
    b64 = b64 & "MSZhbXA7ZGlzcGxheT1zd2FwIiByZWw9InN0eWxlc2hlZXQiLz4NCjxzY3JpcHQgaWQ9InRhaWx3aW5kLWNvbmZpZyI+DQogICAgICB0YWlsd2luZC5jb25maWcgPSB7DQogICAgICAgIGRhcmtNb2RlOiAiY2xhc3MiLA0KICAgICAgICB0"
    b64 = b64 & "aGVtZTogew0KICAgICAgICAgIGV4dGVuZDogew0KICAgICAgICAgICAgY29sb3JzOiB7DQogICAgICAgICAgICAgICJzdXJmYWNlLWNvbnRhaW5lci1sb3ciOiAiI2Y0ZjZmOCIsInNlY29uZGFyeS1maXhlZCI6ICIjZjVlNmQ4IiwiaW52"
    b64 = b64 & "ZXJzZS1vbi1zdXJmYWNlIjogIiNmNGY2ZjgiLA0KICAgICAgICAgICAgICAic3VyZmFjZS12YXJpYW50IjogIiNlNGU4ZWUiLCJwcmltYXJ5LWZpeGVkLWRpbSI6ICIjYzVkNmU4Iiwic3VyZmFjZS1jb250YWluZXIiOiAiI2VlZjFmNSIs"
    b64 = b64 & "DQogICAgICAgICAgICAgICJzdXJmYWNlLWNvbnRhaW5lci1oaWdoIjogIiNlY2VmZjMiLCJwcmltYXJ5LWZpeGVkIjogIiNkY2U3ZjIiLCJvbi10ZXJ0aWFyeSI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeSI6"
    b64 = b64 & "ICIjZmZmZmZmIiwib24tcHJpbWFyeS1maXhlZC12YXJpYW50IjogIiM0YTY3ODUiLCJzZWNvbmRhcnktY29udGFpbmVyIjogIiNmMGQ1YzAiLA0KICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWluZXItbG93ZXN0IjogIiNmZmZmZmYi"
    b64 = b64 & "LCJwcmltYXJ5LWNvbnRhaW5lciI6ICIjYjdjY2UwIiwiZXJyb3ItY29udGFpbmVyIjogIiNmNWQ1ZDIiLA0KICAgICAgICAgICAgICAib24tZXJyb3ItY29udGFpbmVyIjogIiM3YTJlMmEiLCJvbi10ZXJ0aWFyeS1maXhlZC12YXJpYW50"
    b64 = b64 & "IjogIiM3YTVhNGMiLCJwcmltYXJ5IjogIiM2YjhmYjgiLA0KICAgICAgICAgICAgICAidGVydGlhcnktY29udGFpbmVyIjogIiNkNGI1YTUiLCJpbnZlcnNlLXByaW1hcnkiOiAiI2M1ZDZlOCIsInRlcnRpYXJ5LWZpeGVkIjogIiNmMGUw"
    b64 = b64 & "ZDgiLA0KICAgICAgICAgICAgICAic3VyZmFjZS1kaW0iOiAiI2U0ZThlZSIsIm9uLXRlcnRpYXJ5LWNvbnRhaW5lciI6ICIjZmZmOGY1Iiwic3VyZmFjZS10aW50IjogIiM2YjhmYjgiLA0KICAgICAgICAgICAgICAib3V0bGluZSI6ICIj"
    b64 = b64 & "OWFhM2IwIiwib24tc3VyZmFjZS12YXJpYW50IjogIiM1YTY0NzIiLCJvbi1zdXJmYWNlIjogIiMyYTMzNDAiLCJzZWNvbmRhcnkiOiAiI2M5OTU2ZSIsDQogICAgICAgICAgICAgICJvbi1lcnJvciI6ICIjZmZmZmZmIiwic3VyZmFjZS1j"
    b64 = b64 & "b250YWluZXItaGlnaGVzdCI6ICIjZTRlOGVlIiwiYmFja2dyb3VuZCI6ICIjZjdmOGZhIiwNCiAgICAgICAgICAgICAgIm9uLXNlY29uZGFyeS1maXhlZCI6ICIjNGEzODJlIiwidGVydGlhcnktZml4ZWQtZGltIjogIiNlMGM4YmMiLCJv"
    b64 = b64 & "dXRsaW5lLXZhcmlhbnQiOiAiI2Q1ZGFlMiIsDQogICAgICAgICAgICAgICJvbi1iYWNrZ3JvdW5kIjogIiMyYTMzNDAiLCJvbi10ZXJ0aWFyeS1maXhlZCI6ICIjNGEzNDJjIiwib24tc2Vjb25kYXJ5LWZpeGVkLXZhcmlhbnQiOiAiIzhh"
    b64 = b64 & "NmE1MiIsDQogICAgICAgICAgICAgICJvbi1zZWNvbmRhcnktY29udGFpbmVyIjogIiM1YTQzMzYiLCJzZWNvbmRhcnktZml4ZWQtZGltIjogIiNlOGM5YjAiLCJ0ZXJ0aWFyeSI6ICIjYjA4OTc4IiwNCiAgICAgICAgICAgICAgImVycm9y"
    b64 = b64 & "IjogIiNjMDcwNmMiLCJvbi1wcmltYXJ5IjogIiNmZmZmZmYiLCJvbi1wcmltYXJ5LWNvbnRhaW5lciI6ICIjMmY0MDU1IiwNCiAgICAgICAgICAgICAgImludmVyc2Utc3VyZmFjZSI6ICIjM2E0MjUwIiwic3VyZmFjZSI6ICIjZjdmOGZh"
    b64 = b64 & "Iiwib24tcHJpbWFyeS1maXhlZCI6ICIjMmY0MDU1Iiwic3VyZmFjZS1icmlnaHQiOiAiI2Y3ZjhmYSINCiAgICAgICAgICAgIH0sDQogICAgICAgICAgICBib3JkZXJSYWRpdXM6IHsgREVGQVVMVDoiMC4yNXJlbSIsIGxnOiIwLjVyZW0i"
    b64 = b64 & "LCB4bDoiMC43NXJlbSIsIGZ1bGw6Ijk5OTlweCIgfSwNCiAgICAgICAgICAgIHNwYWNpbmc6IHsgImdyaWQtZ3V0dGVyIjoiMS4yNXJlbSIsICJjb250YWluZXItcGFkZGluZyI6IjJyZW0iLCAic3RhY2stZ2FwIjoiMC41cmVtIiwgImNh"
    b64 = b64 & "cmQtcGFkZGluZyI6IjFyZW0iIH0sDQogICAgICAgICAgICBmb250RmFtaWx5OiB7ICJoZWFkZXItdGl0bGUiOlsiUHVibGljIFNhbnMiXSwgImxpc3QtaXRlbSI6WyJQdWJsaWMgU2FucyJdLCAiYnV0dG9uLXRleHQiOlsiUHVibGljIFNh"
    b64 = b64 & "bnMiXSwgIm1ldGEtZGF0YSI6WyJQdWJsaWMgU2FucyJdLCAiYm9hcmQtdGl0bGUiOlsiUHVibGljIFNhbnMiXSB9LA0KICAgICAgICAgICAgZm9udFNpemU6IHsNCiAgICAgICAgICAgICAgImhlYWRlci10aXRsZSI6WyIyNHB4Iix7Imxp"
    b64 = b64 & "bmVIZWlnaHQiOiIzMnB4IiwibGV0dGVyU3BhY2luZyI6Ii0wLjAyZW0iLCJmb250V2VpZ2h0IjoiNzAwIn1dLA0KICAgICAgICAgICAgICAibGlzdC1pdGVtIjpbIjE0cHgiLHsibGluZUhlaWdodCI6IjIwcHgiLCJmb250V2VpZ2h0Ijoi"
    b64 = b64 & "NTAwIn1dLA0KICAgICAgICAgICAgICAiYnV0dG9uLXRleHQiOlsiMTRweCIseyJsaW5lSGVpZ2h0IjoiMjBweCIsImZvbnRXZWlnaHQiOiI2MDAifV0sDQogICAgICAgICAgICAgICJtZXRhLWRhdGEiOlsiMTJweCIseyJsaW5lSGVpZ2h0"
    b64 = b64 & "IjoiMTZweCIsImZvbnRXZWlnaHQiOiI0MDAifV0sDQogICAgICAgICAgICAgICJib2FyZC10aXRsZSI6WyIxNnB4Iix7ImxpbmVIZWlnaHQiOiIyNHB4IiwiZm9udFdlaWdodCI6IjcwMCJ9XQ0KICAgICAgICAgICAgfQ0KICAgICAgICAg"
    b64 = b64 & "IH0NCiAgICAgICAgfQ0KICAgICAgfQ0KPC9zY3JpcHQ+DQo8c3R5bGU+DQoubWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCB7IGZvbnQtdmFyaWF0aW9uLXNldHRpbmdzOidGSUxMJyAwLCd3Z2h0JyA0MDAsJ0dSQUQnIDAsJ29wc3onIDI0"
    b64 = b64 & "OyB2ZXJ0aWNhbC1hbGlnbjptaWRkbGU7IH0NCmJvZHkgeyBmb250LWZhbWlseTonUHVibGljIFNhbnMnLHNhbnMtc2VyaWY7IGJhY2tncm91bmQtY29sb3I6I2Y3ZjhmYTsgfQ0KLmRlcHQtY2FyZCB7IHRyYW5zaXRpb246IHRyYW5zZm9y"
    b64 = b64 & "bSAwLjJzIGVhc2UsIGJveC1zaGFkb3cgMC4ycyBlYXNlOyB9DQouZGVwdC1jYXJkOmhvdmVyIHsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKC00cHgpOyBib3gtc2hhZG93OiAwIDEycHggMjRweCByZ2JhKDAsMCwwLDAuMDgpOyB9DQo8L3N0"
    b64 = b64 & "eWxlPg0KPC9oZWFkPg0KPGJvZHkgY2xhc3M9ImJnLWJhY2tncm91bmQgdGV4dC1vbi1iYWNrZ3JvdW5kIG1pbi1oLXNjcmVlbiBmbGV4IGZsZXgtY29sIj4NCjxoZWFkZXIgY2xhc3M9ImZpeGVkIHRvcC0wIGxlZnQtMCByaWdodC0wIHot"
    b64 = b64 & "NTAgYmctcHJpbWFyeSB0ZXh0LW9uLXByaW1hcnkgc2hhZG93LW1kIHJvdW5kZWQtYi14bCBmbGV4IGZsZXgtY29sIG1kOmZsZXgtcm93IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgcHgtY29udGFpbmVyLXBhZGRpbmcgcHktNiB3"
    b64 = b64 & "LWZ1bGwiPg0KPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTQiPg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQgdGV4dC0zeGwiPm5ld3NwYXBlcjwvc3Bhbj4NCjxkaXY+DQo8aDEgY2xhc3M9ImZv"
    b64 = b64 & "bnQtaGVhZGVyLXRpdGxlIHRleHQtaGVhZGVyLXRpdGxlIHRyYWNraW5nLXRpZ2h0Ij7so7zsmpTrtoDsspgg67O064+E7J6Q66OMIO2VqOq7mOuztOq4sDwvaDE+DQo8cCBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEg"
    b64 = b64 & "dGV4dC1vbi1wcmltYXJ5LzgwIG10LTEiPuqzteyngOyCrO2VrSArIOuztOuPhOyekOujjCDCtyDrtoDsspjri7kg7LWc6re8IDXqsJw8L3A+DQo8L2Rpdj4NCjwvZGl2Pg0KPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTMg"
    b64 = b64 & "bXQtNCBtZDptdC0wIj4NCjxhIGhyZWY9Imdvdl9ub3RpY2VfYm9hcmQuaHRtbCIgY2xhc3M9InRleHQtb24tcHJpbWFyeS85MCBob3ZlcjpiZy1wcmltYXJ5LWNvbnRhaW5lci8yMCBweC0zIHB5LTIgcm91bmRlZC1sZyBmb250LWJ1dHRv"
    b64 = b64 & "bi10ZXh0IHRleHQtYnV0dG9uLXRleHQgbm8tdW5kZXJsaW5lIj7qs7Xsp4Ag67O065Oc66GcPC9hPg0KPGJ1dHRvbiBpZD0iYnRuUmVsb2FkIiB0eXBlPSJidXR0b24iIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBiZy1vbi1w"
    b64 = b64 & "cmltYXJ5IHRleHQtcHJpbWFyeSBweC00IHB5LTIgcm91bmRlZC1mdWxsIGZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4dCBob3ZlcjpiZy1wcmltYXJ5LWZpeGVkLWRpbSB0cmFuc2l0aW9uLWFsbCBhY3RpdmU6c2NhbGUtOTUi"
    b64 = b64 & "Pg0KPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiPnJlZnJlc2g8L3NwYW4+7ZmU66m0IOyDiOuhnOqzoOy5qA0KPC9idXR0b24+DQo8L2Rpdj4NCjwvaGVhZGVyPg0KDQo8bWFpbiBjbGFzcz0iZmxleC0xIHB4LTQg"
    b64 = b64 & "bWQ6cHgtY29udGFpbmVyLXBhZGRpbmcgcHktOCBwdC0zMiBtYXgtdy1bMTQ0MHB4XSBteC1hdXRvIHctZnVsbCI+DQo8ZGl2IGNsYXNzPSJmbGV4IGZsZXgtY29sIG1kOmZsZXgtcm93IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1zdGFydCBt"
    b64 = b64 & "ZDppdGVtcy1jZW50ZXIgbWItOCBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBiZy1zdXJmYWNlLWNvbnRhaW5lci1sb3cgcm91bmRlZC14bCBweC02IHB5LTMgZ2FwLTQiPg0KPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIg"
    b64 = b64 & "Z2FwLTQgZmxleC13cmFwIj4NCjxzcGFuIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMSBmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFudCIgaWQ9Im1ldGFMaW5lIj4NCjxzcGFuIGNs"
    b64 = b64 & "YXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE2cHhdIj5zY2hlZHVsZTwvc3Bhbj7rjbDsnbTthLAg7KSA67mEIOykkeKApg0KPC9zcGFuPg0KPC9kaXY+DQo8ZGl2IGNsYXNzPSJmbGV4IGdhcC0yIj4NCjxzcGFuIGNs"
    b64 = b64 & "YXNzPSJweC0zIHB5LTEgcm91bmRlZC1mdWxsIGJvcmRlciBib3JkZXItcHJpbWFyeSB0ZXh0LXByaW1hcnkgZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgYmctd2hpdGUiPuqzteyngOyCrO2VrTwvc3Bhbj4NCjxzcGFuIGNsYXNz"
    b64 = b64 & "PSJweC0zIHB5LTEgcm91bmRlZC1mdWxsIGJvcmRlciBib3JkZXItc2Vjb25kYXJ5IHRleHQtc2Vjb25kYXJ5IGZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIGJnLXdoaXRlIj7rs7Trj4TsnpDro4w8L3NwYW4+DQo8L2Rpdj4NCjwv"
    b64 = b64 & "ZGl2Pg0KPHNlY3Rpb24gY2xhc3M9ImZsZXggZmxleC1jb2wgZ2FwLTgiIGlkPSJib2FyZCI+PC9zZWN0aW9uPg0KPC9tYWluPg0KDQo8Zm9vdGVyIGNsYXNzPSJ3LWZ1bGwgcHktNCBweC1jb250YWluZXItcGFkZGluZyBmbGV4IGZsZXgt"
    b64 = b64 & "Y29sIG1kOmZsZXgtcm93IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIgZ2FwLTMgYmctc3VyZmFjZS1kaW0gdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgYm9yZGVyLXQgYm9yZGVyLW91dGxpbmUtdmFyaWFudCBtdC1hdXRvIj4NCjxz"
    b64 = b64 & "cGFuIGNsYXNzPSJmb250LWJvbGQgZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEiPuuMgO2VnOuvvOq1rSDsoJXrtoAgwrcg7KO87JqU67aA7LKYIOuztOuTnDwvc3Bhbj4NCjxkaXYgY2xhc3M9ImZsZXggZ2FwLTYiPg0KPGEgY2xh"
    b64 = b64 & "c3M9ImZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIGhvdmVyOnRleHQtcHJpbWFyeSIgaHJlZj0iZ292X25vdGljZV9ib2FyZC5odG1sIj5QQyDqs7Xsp4A8L2E+DQo8YSBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRh"
    b64 = b64 & "dGEgaG92ZXI6dGV4dC1wcmltYXJ5IiBocmVmPSJnb3ZfbWFqb3JfcHJlc3NfYm9hcmRfbW9iaWxlLmh0bWwiPuuqqOuwlOydvCDtlajqu5jrs7TquLA8L2E+DQo8L2Rpdj4NCjwvZm9vdGVyPg0KDQo8c2NyaXB0IGlkPSJtYWpvci1kYXRh"
    b64 = b64 & "IiB0eXBlPSJhcHBsaWNhdGlvbi9qc29uIj4NCiUlTUFKT1JfSlNPTiUlDQo8L3NjcmlwdD4NCjxzY3JpcHQ+DQooZnVuY3Rpb24gKCkgew0KICBjb25zdCBNRVRBID0gew0KICAgICfsgrDsl4XthrXsg4HrtoAnOiB7IGNvbG9yOicjYjdk"
    b64 = b64 & "MGVhJywgaWNvbjonZmFjdG9yeScsIGRvbWFpbjonTU9USVIuR08uS1InIH0sDQogICAgJ+q4sO2bhOyXkOuEiOyngO2ZmOqyveu2gCc6IHsgY29sb3I6JyNiOGRkZDYnLCBpY29uOidlY28nLCBkb21haW46J01DRUUuR08uS1InIH0sDQog"
    b64 = b64 & "ICAgJ+qzoOyaqeuFuOuPmeu2gCc6IHsgY29sb3I6JyNmMGQwYjAnLCBpY29uOidncm91cHMnLCBkb21haW46J01PRUwuR08uS1InIH0sDQogICAgJ+yZuOq1kOu2gCc6IHsgY29sb3I6JyNjNWM4ZTgnLCBpY29uOidwdWJsaWMnLCBkb21h"
    b64 = b64 & "aW46J01PRkEuR08uS1InIH0NCiAgfTsNCiAgY29uc3QgT1JERVIgPSBbJ+yCsOyXhe2GteyDgeu2gCcsJ+q4sO2bhOyXkOuEiOyngO2ZmOqyveu2gCcsJ+qzoOyaqeuFuOuPmeu2gCcsJ+yZuOq1kOu2gCddOw0KICBmdW5jdGlvbiBlc2Mo"
    b64 = b64 & "cyl7IHJldHVybiBTdHJpbmcocz8/JycpLnJlcGxhY2UoLyYvZywnJmFtcDsnKS5yZXBsYWNlKC88L2csJyZsdDsnKS5yZXBsYWNlKC8+L2csJyZndDsnKS5yZXBsYWNlKC8iL2csJyZxdW90OycpOyB9DQogIGZ1bmN0aW9uIGxpc3RIdG1s"
    b64 = b64 & "KGl0ZW1zLCBhY2NlbnQsIGlzUHJlc3MpIHsNCiAgICBjb25zdCBhcnIgPSAoaXRlbXN8fFtdKS5zbGljZSgwLDUpOw0KICAgIGlmICghYXJyLmxlbmd0aCkgcmV0dXJuICc8bGkgY2xhc3M9InB5LTMgdGV4dC1vbi1zdXJmYWNlLXZhcmlh"
    b64 = b64 & "bnQgZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEiPu2RnOyLnO2VoCDtla3rqqnsnbQg7JeG7Iq164uI64ukLjwvbGk+JzsNCiAgICByZXR1cm4gYXJyLm1hcChmdW5jdGlvbihwLGkpew0KICAgICAgY29uc3QgbGFzdCA9IGkgPT09"
    b64 = b64 & "IGFyci5sZW5ndGgtMTsNCiAgICAgIGNvbnN0IGJvcmRlciA9IGxhc3QgPyAnJyA6IChpc1ByZXNzID8gJyBib3JkZXItYiBib3JkZXItc2Vjb25kYXJ5LzEwJyA6ICcgYm9yZGVyLWIgYm9yZGVyLW91dGxpbmUtdmFyaWFudC8xMCcpOw0K"
    b64 = b64 & "ICAgICAgY29uc3QgZGF0ZSA9IHAuZGF0ZSA/ICc8c3BhbiBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQgd2hpdGVzcGFjZS1ub3dyYXAiPicgKyBlc2MocC5kYXRlKSArICc8"
    b64 = b64 & "L3NwYW4+JyA6ICcnOw0KICAgICAgcmV0dXJuICc8bGkgY2xhc3M9ImZsZXgganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLXN0YXJ0IGdhcC00IHB5LTInICsgYm9yZGVyICsgJyI+JyArDQogICAgICAgICc8YSBjbGFzcz0iZm9udC1saXN0LWl0"
    b64 = b64 & "ZW0gdGV4dC1saXN0LWl0ZW0gdGV4dC1vbi1zdXJmYWNlIGxpbmUtY2xhbXAtMSBob3ZlcjpvcGFjaXR5LTgwIG5vLXVuZGVybGluZSIgc3R5bGU9ImNvbG9yOmluaGVyaXQiIGhyZWY9IicgKyBlc2MocC51cmx8fCcjJykgKyAnIiB0YXJn"
    b64 = b64 & "ZXQ9Il9ibGFuayIgcmVsPSJub29wZW5lciBub3JlZmVycmVyIj4nICsgZXNjKHAudGl0bGUpICsgJzwvYT4nICsgZGF0ZSArICc8L2xpPic7DQogICAgfSkuam9pbignJyk7DQogIH0NCiAgZnVuY3Rpb24gcmVuZGVyKGRhdGEpIHsNCiAg"
    b64 = b64 & "ICBjb25zdCBib2FyZCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdib2FyZCcpOw0KICAgIGNvbnN0IG1hcCA9IHt9Ow0KICAgIChkYXRhLmRlcGFydG1lbnRzfHxbXSkuZm9yRWFjaChmdW5jdGlvbihkKXsgbWFwW2QubmFtZV09ZDsg"
    b64 = b64 & "fSk7DQogICAgY29uc3QgZGVwdHMgPSBPUkRFUi5tYXAoZnVuY3Rpb24obil7IHJldHVybiBtYXBbbl0gfHwgeyBuYW1lOm4sIG9rOmZhbHNlLCBub3RpY2VzOltdLCBwcmVzczpbXSB9OyB9KTsNCiAgICBkb2N1bWVudC5nZXRFbGVtZW50"
    b64 = b64 & "QnlJZCgnbWV0YUxpbmUnKS5pbm5lckhUTUwgPQ0KICAgICAgJzxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE2cHhdIj5zY2hlZHVsZTwvc3Bhbj7stZzqt7wg7JeF642w7J207Yq4OiA8c3Ryb25nIGNs"
    b64 = b64 & "YXNzPSJtbC0xIj4nICsgZXNjKGRhdGEuY29sbGVjdGVkQXR8fCctJykgKyAnPC9zdHJvbmc+JyArDQogICAgICAnPHNwYW4gY2xhc3M9Im14LTIiPnw8L3NwYW4+66qo64uI7YSw66eBIOu2gOyymDogPHN0cm9uZz4nICsgZGVwdHMubGVu"
    b64 = b64 & "Z3RoICsgJ+qwnDwvc3Ryb25nPic7DQogICAgYm9hcmQuaW5uZXJIVE1MID0gZGVwdHMubWFwKGZ1bmN0aW9uKGQpew0KICAgICAgY29uc3QgbSA9IE1FVEFbZC5uYW1lXSB8fCB7IGNvbG9yOicjMDA0YWM2JywgaWNvbjonc3RhcicsIGRv"
    b64 = b64 & "bWFpbjonJyB9Ow0KICAgICAgcmV0dXJuICc8YXJ0aWNsZSBjbGFzcz0iZGVwdC1jYXJkIGJnLXdoaXRlIHJvdW5kZWQteGwgc2hhZG93LVswcHhfNHB4XzEycHhfcmdiYSgwLDAsMCwwLjA1KV0gb3ZlcmZsb3ctaGlkZGVuIGJvcmRlciBi"
    b64 = b64 & "b3JkZXItb3V0bGluZS12YXJpYW50LzMwIGZsZXggZmxleC1jb2wiPicgKw0KICAgICAgICAnPGRpdiBjbGFzcz0icHgtNiBweS00IGZsZXgganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLWNlbnRlciIgc3R5bGU9ImJhY2tncm91bmQ6JyArIG0u"
    b64 = b64 & "Y29sb3IgKyAnIj4nICsNCiAgICAgICAgJzxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0zIj48c3BhbiBjbGFzcz0ibWF0ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCIgc3R5bGU9ImNvbG9yOiMzZDRhNWMiPicgKyBtLmljb24g"
    b64 = b64 & "KyAnPC9zcGFuPicgKw0KICAgICAgICAnPGgzIGNsYXNzPSJmb250LWJvYXJkLXRpdGxlIHRleHQtYm9hcmQtdGl0bGUiIHN0eWxlPSJjb2xvcjojM2Q0YTVjIj4nICsgZXNjKGQubmFtZSkgKyAnPC9oMz48L2Rpdj4nICsNCiAgICAgICAg"
    b64 = b64 & "JzxzcGFuIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSIgc3R5bGU9ImNvbG9yOiM1YTY0NzIiPicgKyBlc2MobS5kb21haW4pICsgJzwvc3Bhbj48L2Rpdj4nICsNCiAgICAgICAgJzxkaXYgY2xhc3M9ImdyaWQgZ3Jp"
    b64 = b64 & "ZC1jb2xzLTEgbWQ6Z3JpZC1jb2xzLTIgZGl2aWRlLXggZGl2aWRlLW91dGxpbmUtdmFyaWFudC8yMCI+JyArDQogICAgICAgICc8ZGl2IGNsYXNzPSJwLTYiPjxoNCBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgbWItNCBmb250"
    b64 = b64 & "LWJ1dHRvbi10ZXh0IHRleHQtYnV0dG9uLXRleHQiIHN0eWxlPSJjb2xvcjojNmI4ZmI4Ij4nICsNCiAgICAgICAgJzxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE4cHhdIj5ub3RpZmljYXRpb25zPC9z"
    b64 = b64 & "cGFuPuqzteyngOyCrO2VrTwvaDQ+JyArDQogICAgICAgICc8dWwgY2xhc3M9InNwYWNlLXktMSI+JyArIGxpc3RIdG1sKGQubm90aWNlcywgbS5jb2xvciwgZmFsc2UpICsgJzwvdWw+PC9kaXY+JyArDQogICAgICAgICc8ZGl2IGNsYXNz"
    b64 = b64 & "PSJwLTYgYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0LzUwIj48aDQgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIG1iLTQgZm9udC1idXR0b24tdGV4dCB0ZXh0LWJ1dHRvbi10ZXh0IHRleHQtc2Vjb25kYXJ5Ij4nICsNCiAg"
    b64 = b64 & "ICAgICAgJzxzcGFuIGNsYXNzPSJtYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHRleHQtWzE4cHhdIj5uZXdzcGFwZXI8L3NwYW4+67O064+E7J6Q66OMPC9oND4nICsNCiAgICAgICAgJzx1bCBjbGFzcz0ic3BhY2UteS0xIj4nICsgbGlz"
    b64 = b64 & "dEh0bWwoZC5wcmVzcywgbS5jb2xvciwgdHJ1ZSkgKyAnPC91bD48L2Rpdj4nICsNCiAgICAgICAgJzwvZGl2PjwvYXJ0aWNsZT4nOw0KICAgIH0pLmpvaW4oJycpOw0KICB9DQogIGZ1bmN0aW9uIGxvYWQoKXsNCiAgICBjb25zdCByYXcg"
    b64 = b64 & "PSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWFqb3ItZGF0YScpLnRleHRDb250ZW50LnRyaW0oKTsNCiAgICBsZXQgZGF0YSA9IHsgY29sbGVjdGVkQXQ6JycsIGRlcGFydG1lbnRzOltdIH07DQogICAgdHJ5IHsgaWYgKHJhdyAmJiBy"
    b64 = b64 & "YXcuY2hhckF0KDApPT09J3snKSBkYXRhID0gSlNPTi5wYXJzZShyYXcpOyB9IGNhdGNoKGUpeyBjb25zb2xlLmVycm9yKGUpOyB9DQogICAgcmVuZGVyKGRhdGEpOw0KICB9DQogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdidG5SZWxv"
    b64 = b64 & "YWQnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGZ1bmN0aW9uKCl7IGxvY2F0aW9uLnJlbG9hZCgpOyB9KTsNCiAgbG9hZCgpOw0KfSkoKTsNCjwvc2NyaXB0Pg0KPC9ib2R5Pg0KPC9odG1sPg0K"
    EmbeddedMajorPressTemplatePc = DecodeBase64Utf8(b64)
End Function

Private Function EmbeddedMajorPressTemplateMobile() As String
    Dim b64 As String
    b64 = ""
    b64 = b64 & "77u/PCFET0NUWVBFIGh0bWw+DQo8aHRtbCBjbGFzcz0ibGlnaHQiIGxhbmc9ImtvIj4NCjxoZWFkPg0KPG1ldGEgY2hhcnNldD0idXRmLTgiLz4NCjxtZXRhIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0x"
    b64 = b64 & "LjAiIG5hbWU9InZpZXdwb3J0Ii8+DQo8dGl0bGU+7KO87JqU67aA7LKYIOqzteyngMK367O064+E7J6Q66OMIO2VqOq7mOuztOq4sCAo66qo67CU7J28KTwvdGl0bGU+DQo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4udGFpbHdpbmRjc3Mu"
    b64 = b64 & "Y29tP3BsdWdpbnM9Zm9ybXMsY29udGFpbmVyLXF1ZXJpZXMiPjwvc2NyaXB0Pg0KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1QdWJsaWMrU2Fuczp3Z2h0QDQwMDs1MDA7NjAwOzcwMDs4"
    b64 = b64 & "MDAmYW1wO2Rpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0Ii8+DQo8bGluayBocmVmPSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tL2NzczI/ZmFtaWx5PU1hdGVyaWFsK1N5bWJvbHMrT3V0bGluZWQ6d2dodCxGSUxMQDEwMC4u"
    b64 = b64 & "NzAwLDAuLjEmYW1wO2Rpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0Ii8+DQo8c2NyaXB0IGlkPSJ0YWlsd2luZC1jb25maWciPg0KICAgICAgdGFpbHdpbmQuY29uZmlnID0gew0KICAgICAgICBkYXJrTW9kZTogImNsYXNzIiwNCiAg"
    b64 = b64 & "ICAgICAgdGhlbWU6IHsNCiAgICAgICAgICBleHRlbmQ6IHsNCiAgICAgICAgICAgIGNvbG9yczogew0KICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWluZXItbG93IjogIiNmNGY2ZjgiLCJzZWNvbmRhcnktZml4ZWQiOiAiI2Y1ZTZk"
    b64 = b64 & "OCIsImludmVyc2Utb24tc3VyZmFjZSI6ICIjZjRmNmY4IiwNCiAgICAgICAgICAgICAgInN1cmZhY2UtdmFyaWFudCI6ICIjZTRlOGVlIiwicHJpbWFyeS1maXhlZC1kaW0iOiAiI2M1ZDZlOCIsInN1cmZhY2UtY29udGFpbmVyIjogIiNl"
    b64 = b64 & "ZWYxZjUiLA0KICAgICAgICAgICAgICAic3VyZmFjZS1jb250YWluZXItaGlnaCI6ICIjZWNlZmYzIiwicHJpbWFyeS1maXhlZCI6ICIjZGNlN2YyIiwib24tdGVydGlhcnkiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICJvbi1zZWNv"
    b64 = b64 & "bmRhcnkiOiAiI2ZmZmZmZiIsInNlY29uZGFyeS1jb250YWluZXIiOiAiI2YwZDVjMCIsInN1cmZhY2UtY29udGFpbmVyLWxvd2VzdCI6ICIjZmZmZmZmIiwNCiAgICAgICAgICAgICAgInByaW1hcnktY29udGFpbmVyIjogIiNiN2NjZTAi"
    b64 = b64 & "LCJlcnJvci1jb250YWluZXIiOiAiI2Y1ZDVkMiIsInByaW1hcnkiOiAiIzZiOGZiOCIsDQogICAgICAgICAgICAgICJ0ZXJ0aWFyeS1jb250YWluZXIiOiAiI2Q0YjVhNSIsInN1cmZhY2UtZGltIjogIiNlNGU4ZWUiLCJvdXRsaW5lIjog"
    b64 = b64 & "IiM5YWEzYjAiLA0KICAgICAgICAgICAgICAib24tc3VyZmFjZS12YXJpYW50IjogIiM1YTY0NzIiLCJvbi1zdXJmYWNlIjogIiMyYTMzNDAiLCJzZWNvbmRhcnkiOiAiI2M5OTU2ZSIsDQogICAgICAgICAgICAgICJzdXJmYWNlLWNvbnRh"
    b64 = b64 & "aW5lci1oaWdoZXN0IjogIiNlNGU4ZWUiLCJiYWNrZ3JvdW5kIjogIiNmN2Y4ZmEiLCJvdXRsaW5lLXZhcmlhbnQiOiAiI2Q1ZGFlMiIsDQogICAgICAgICAgICAgICJvbi1iYWNrZ3JvdW5kIjogIiMyYTMzNDAiLCJ0ZXJ0aWFyeSI6ICIj"
    b64 = b64 & "YjA4OTc4IiwiZXJyb3IiOiAiI2MwNzA2YyIsIm9uLXByaW1hcnkiOiAiI2ZmZmZmZiIsDQogICAgICAgICAgICAgICJvbi1wcmltYXJ5LWNvbnRhaW5lciI6ICIjMmY0MDU1Iiwib24tdGVydGlhcnktY29udGFpbmVyIjogIiNmZmY4ZjUi"
    b64 = b64 & "LCJzdXJmYWNlIjogIiNmN2Y4ZmEiDQogICAgICAgICAgICB9LA0KICAgICAgICAgICAgYm9yZGVyUmFkaXVzOiB7IERFRkFVTFQ6IjAuMjVyZW0iLCBsZzoiMC41cmVtIiwgeGw6IjAuNzVyZW0iLCBmdWxsOiI5OTk5cHgiIH0sDQogICAg"
    b64 = b64 & "ICAgICAgICBzcGFjaW5nOiB7ICJjb250YWluZXItcGFkZGluZyI6IjJyZW0iLCAiY2FyZC1wYWRkaW5nIjoiMXJlbSIgfSwNCiAgICAgICAgICAgIGZvbnRGYW1pbHk6IHsgImhlYWRlci10aXRsZSI6WyJQdWJsaWMgU2FucyJdLCAibGlz"
    b64 = b64 & "dC1pdGVtIjpbIlB1YmxpYyBTYW5zIl0sICJidXR0b24tdGV4dCI6WyJQdWJsaWMgU2FucyJdLCAibWV0YS1kYXRhIjpbIlB1YmxpYyBTYW5zIl0sICJib2FyZC10aXRsZSI6WyJQdWJsaWMgU2FucyJdIH0sDQogICAgICAgICAgICBmb250"
    b64 = b64 & "U2l6ZTogew0KICAgICAgICAgICAgICAiaGVhZGVyLXRpdGxlIjpbIjI0cHgiLHsibGluZUhlaWdodCI6IjMycHgiLCJsZXR0ZXJTcGFjaW5nIjoiLTAuMDJlbSIsImZvbnRXZWlnaHQiOiI3MDAifV0sDQogICAgICAgICAgICAgICJsaXN0"
    b64 = b64 & "LWl0ZW0iOlsiMTRweCIseyJsaW5lSGVpZ2h0IjoiMjBweCIsImZvbnRXZWlnaHQiOiI1MDAifV0sDQogICAgICAgICAgICAgICJidXR0b24tdGV4dCI6WyIxNHB4Iix7ImxpbmVIZWlnaHQiOiIyMHB4IiwiZm9udFdlaWdodCI6IjYwMCJ9"
    b64 = b64 & "XSwNCiAgICAgICAgICAgICAgIm1ldGEtZGF0YSI6WyIxMnB4Iix7ImxpbmVIZWlnaHQiOiIxNnB4IiwiZm9udFdlaWdodCI6IjQwMCJ9XSwNCiAgICAgICAgICAgICAgImJvYXJkLXRpdGxlIjpbIjE2cHgiLHsibGluZUhlaWdodCI6IjI0"
    b64 = b64 & "cHgiLCJmb250V2VpZ2h0IjoiNzAwIn1dDQogICAgICAgICAgICB9DQogICAgICAgICAgfQ0KICAgICAgICB9DQogICAgICB9DQo8L3NjcmlwdD4NCjxzdHlsZT4NCi5tYXRlcmlhbC1zeW1ib2xzLW91dGxpbmVkIHsgZm9udC12YXJpYXRp"
    b64 = b64 & "b24tc2V0dGluZ3M6J0ZJTEwnIDAsJ3dnaHQnIDQwMCwnR1JBRCcgMCwnb3BzeicgMjQ7IH0NCmJvZHkgeyBmb250LWZhbWlseTonUHVibGljIFNhbnMnLHNhbnMtc2VyaWY7IH0NCjwvc3R5bGU+DQo8L2hlYWQ+DQo8Ym9keSBjbGFzcz0i"
    b64 = b64 & "YmctYmFja2dyb3VuZCB0ZXh0LW9uLWJhY2tncm91bmQgbWluLWgtc2NyZWVuIGZsZXggZmxleC1jb2wiPg0KPGhlYWRlciBjbGFzcz0iYmctcHJpbWFyeSB0ZXh0LW9uLXByaW1hcnkgZml4ZWQgdG9wLTAgbGVmdC0wIHJpZ2h0LTAgei01"
    b64 = b64 & "MCByb3VuZGVkLWIteGwgc2hhZG93LW1kIj4NCjxkaXYgY2xhc3M9ImZsZXggZmxleC1jb2wganVzdGlmeS1iZXR3ZWVuIGl0ZW1zLXN0cmV0Y2ggcHgtNCBweS01IHctZnVsbCBnYXAtMyI+DQo8ZGl2IGNsYXNzPSJmbGV4IGp1c3RpZnkt"
    b64 = b64 & "YmV0d2VlbiBpdGVtcy1jZW50ZXIiPg0KPGgxIGNsYXNzPSJmb250LWhlYWRlci10aXRsZSB0ZXh0LWhlYWRlci10aXRsZSI+7KO87JqU67aA7LKYIO2VqOq7mOuztOq4sDwvaDE+DQo8YnV0dG9uIGlkPSJidG5SZWxvYWQiIHR5cGU9ImJ1"
    b64 = b64 & "dHRvbiIgY2xhc3M9ImZvbnQtYnV0dG9uLXRleHQgdGV4dC1idXR0b24tdGV4dCBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMSBwLTIgcm91bmRlZC1sZyBob3ZlcjpiZy1wcmltYXJ5LWNvbnRhaW5lci8yMCI+DQo8c3BhbiBjbGFzcz0ibWF0"
    b64 = b64 & "ZXJpYWwtc3ltYm9scy1vdXRsaW5lZCI+cmVmcmVzaDwvc3Bhbj4NCjwvYnV0dG9uPg0KPC9kaXY+DQo8YSBocmVmPSJnb3Zfbm90aWNlX2JvYXJkX21vYmlsZS5odG1sIiBjbGFzcz0idGV4dC1vbi1wcmltYXJ5LzkwIGZvbnQtbWV0YS1k"
    b64 = b64 & "YXRhIHRleHQtbWV0YS1kYXRhIG5vLXVuZGVybGluZSI+4oaQIOuqqOuwlOydvCDqs7Xsp4Ag67O065OcPC9hPg0KPC9kaXY+DQo8L2hlYWRlcj4NCg0KPGRpdiBjbGFzcz0ibXQtMjggcHgtNCBtYi00Ij4NCjxkaXYgY2xhc3M9ImJnLXN1"
    b64 = b64 & "cmZhY2UtY29udGFpbmVyLWxvdyBib3JkZXIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCByb3VuZGVkLXhsIHAtMyBmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIiPg0KPGRpdiBjbGFzcz0iZmxleCBmbGV4LWNvbCI+DQo8c3Bh"
    b64 = b64 & "biBjbGFzcz0iZm9udC1tZXRhLWRhdGEgdGV4dC1tZXRhLWRhdGEgdGV4dC1vbi1zdXJmYWNlLXZhcmlhbnQiPuyXheuNsOydtO2KuCDsi5zqsIQ8L3NwYW4+DQo8c3BhbiBjbGFzcz0iZm9udC1saXN0LWl0ZW0gdGV4dC1saXN0LWl0ZW0g"
    b64 = b64 & "dGV4dC1vbi1zdXJmYWNlIiBpZD0ibWV0YVRpbWUiPi08L3NwYW4+DQo8L2Rpdj4NCjxkaXYgY2xhc3M9ImZsZXggZmxleC1jb2wgaXRlbXMtZW5kIj4NCjxzcGFuIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSB0ZXh0"
    b64 = b64 & "LW9uLXN1cmZhY2UtdmFyaWFudCI+67aA7LKYIOyImDwvc3Bhbj4NCjxzcGFuIGNsYXNzPSJmb250LWxpc3QtaXRlbSB0ZXh0LWxpc3QtaXRlbSB0ZXh0LW9uLXN1cmZhY2UgZm9udC1ib2xkIj7so7zsmpTrtoDsspggNOqwnDwvc3Bhbj4N"
    b64 = b64 & "CjwvZGl2Pg0KPC9kaXY+DQo8L2Rpdj4NCg0KPG1haW4gY2xhc3M9InB4LTQgcGItMjQgZ3JpZCBncmlkLWNvbHMtMSBnYXAtNiIgaWQ9ImJvYXJkIj48L21haW4+DQoNCjxmb290ZXIgY2xhc3M9ImJnLXN1cmZhY2UtZGltIHRleHQtb24t"
    b64 = b64 & "c3VyZmFjZS12YXJpYW50IGJvcmRlci10IGJvcmRlci1vdXRsaW5lLXZhcmlhbnQgdy1mdWxsIHB5LTQgcHgtNCBmbGV4IGZsZXgtY29sIGl0ZW1zLWNlbnRlciBnYXAtMyBtdC1hdXRvIj4NCjxwIGNsYXNzPSJmb250LW1ldGEtZGF0YSB0"
    b64 = b64 & "ZXh0LW1ldGEtZGF0YSI+wqkg64yA7ZWc66+86rWtIOygleu2gCDso7zsmpTrtoDsspgg67O065OcPC9wPg0KPGEgY2xhc3M9ImZvbnQtbWV0YS1kYXRhIHRleHQtbWV0YS1kYXRhIG5vLXVuZGVybGluZSBob3Zlcjp0ZXh0LXByaW1hcnki"
    b64 = b64 & "IGhyZWY9Imdvdl9tYWpvcl9wcmVzc19ib2FyZC5odG1sIj5QQyDtlajqu5jrs7TquLA8L2E+DQo8L2Zvb3Rlcj4NCg0KPHNjcmlwdCBpZD0ibWFqb3ItZGF0YSIgdHlwZT0iYXBwbGljYXRpb24vanNvbiI+DQolJU1BSk9SX0pTT04lJQ0K"
    b64 = b64 & "PC9zY3JpcHQ+DQo8c2NyaXB0Pg0KKGZ1bmN0aW9uICgpIHsNCiAgY29uc3QgTUVUQSA9IHsNCiAgICAn7IKw7JeF7Ya17IOB67aAJzogeyBjb2xvcjonI2I3ZDBlYScsIGljb246J2ZhY3RvcnknIH0sDQogICAgJ+q4sO2bhOyXkOuEiOyn"
    b64 = b64 & "gO2ZmOqyveu2gCc6IHsgY29sb3I6JyNiOGRkZDYnLCBpY29uOidlY28nIH0sDQogICAgJ+qzoOyaqeuFuOuPmeu2gCc6IHsgY29sb3I6JyNmMGQwYjAnLCBpY29uOidncm91cHMnIH0sDQogICAgJ+yZuOq1kOu2gCc6IHsgY29sb3I6JyNj"
    b64 = b64 & "NWM4ZTgnLCBpY29uOidwdWJsaWMnIH0NCiAgfTsNCiAgY29uc3QgT1JERVIgPSBbJ+yCsOyXhe2GteyDgeu2gCcsJ+q4sO2bhOyXkOuEiOyngO2ZmOqyveu2gCcsJ+qzoOyaqeuFuOuPmeu2gCcsJ+yZuOq1kOu2gCddOw0KICBmdW5jdGlv"
    b64 = b64 & "biBlc2Mocyl7IHJldHVybiBTdHJpbmcocz8/JycpLnJlcGxhY2UoLyYvZywnJmFtcDsnKS5yZXBsYWNlKC88L2csJyZsdDsnKS5yZXBsYWNlKC8+L2csJyZndDsnKS5yZXBsYWNlKC8iL2csJyZxdW90OycpOyB9DQogIGZ1bmN0aW9uIGl0"
    b64 = b64 & "ZW1zSHRtbChhcnIsIGNvbG9yQ2xhc3MpIHsNCiAgICBjb25zdCBsaXN0ID0gKGFycnx8W10pLnNsaWNlKDAsNSk7DQogICAgaWYgKCFsaXN0Lmxlbmd0aCkgcmV0dXJuICc8bGkgY2xhc3M9InAtMyB0ZXh0LW9uLXN1cmZhY2UtdmFyaWFu"
    b64 = b64 & "dCBmb250LW1ldGEtZGF0YSB0ZXh0LW1ldGEtZGF0YSI+7ZWt66qpIOyXhuydjDwvbGk+JzsNCiAgICByZXR1cm4gbGlzdC5tYXAoZnVuY3Rpb24ocCxpKXsNCiAgICAgIHJldHVybiAnPGxpIGNsYXNzPSJwLTMgZm9udC1saXN0LWl0ZW0g"
    b64 = b64 & "dGV4dC1saXN0LWl0ZW0gZmxleCBnYXAtMyBpdGVtcy1zdGFydCByb3VuZGVkLWxnIj4nICsNCiAgICAgICAgJzxzcGFuIGNsYXNzPSJmb250LWJvbGQgJyArIGNvbG9yQ2xhc3MgKyAnIj4nICsgKGkrMSkgKyAnPC9zcGFuPicgKw0KICAg"
    b64 = b64 & "ICAgICAnPGEgY2xhc3M9ImxpbmUtY2xhbXAtMiB0ZXh0LW9uLXN1cmZhY2Ugbm8tdW5kZXJsaW5lIiBocmVmPSInICsgZXNjKHAudXJsfHwnIycpICsgJyIgdGFyZ2V0PSJfYmxhbmsiIHJlbD0ibm9vcGVuZXIgbm9yZWZlcnJlciI+JyAr"
    b64 = b64 & "IGVzYyhwLnRpdGxlKSArICc8L2E+PC9saT4nOw0KICAgIH0pLmpvaW4oJycpOw0KICB9DQogIGZ1bmN0aW9uIHRvZ2dsZVRhYihidG4sIHNob3dJZCwgaGlkZUlkKSB7DQogICAgY29uc3QgY29udGFpbmVyID0gYnRuLnBhcmVudEVsZW1l"
    b64 = b64 & "bnQ7DQogICAgY29udGFpbmVyLnF1ZXJ5U2VsZWN0b3JBbGwoJ2J1dHRvbicpLmZvckVhY2goZnVuY3Rpb24oYil7DQogICAgICBiLmNsYXNzTGlzdC5yZW1vdmUoJ2JvcmRlci1wcmltYXJ5JywndGV4dC1wcmltYXJ5Jyk7DQogICAgICBi"
    b64 = b64 & "LmNsYXNzTGlzdC5hZGQoJ2JvcmRlci10cmFuc3BhcmVudCcsJ3RleHQtb24tc3VyZmFjZS12YXJpYW50Jyk7DQogICAgfSk7DQogICAgYnRuLmNsYXNzTGlzdC5hZGQoJ2JvcmRlci1wcmltYXJ5JywndGV4dC1wcmltYXJ5Jyk7DQogICAg"
    b64 = b64 & "YnRuLmNsYXNzTGlzdC5yZW1vdmUoJ2JvcmRlci10cmFuc3BhcmVudCcsJ3RleHQtb24tc3VyZmFjZS12YXJpYW50Jyk7DQogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoc2hvd0lkKS5jbGFzc0xpc3QucmVtb3ZlKCdoaWRkZW4nKTsN"
    b64 = b64 & "CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChzaG93SWQpLmNsYXNzTGlzdC5hZGQoJ2Jsb2NrJyk7DQogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaGlkZUlkKS5jbGFzc0xpc3QuYWRkKCdoaWRkZW4nKTsNCiAgICBkb2N1bWVu"
    b64 = b64 & "dC5nZXRFbGVtZW50QnlJZChoaWRlSWQpLmNsYXNzTGlzdC5yZW1vdmUoJ2Jsb2NrJyk7DQogIH0NCiAgd2luZG93LnRvZ2dsZVRhYiA9IHRvZ2dsZVRhYjsNCiAgZnVuY3Rpb24gcmVuZGVyKGRhdGEpIHsNCiAgICBkb2N1bWVudC5nZXRF"
    b64 = b64 & "bGVtZW50QnlJZCgnbWV0YVRpbWUnKS50ZXh0Q29udGVudCA9IGRhdGEuY29sbGVjdGVkQXQgfHwgJy0nOw0KICAgIGNvbnN0IG1hcCA9IHt9Ow0KICAgIChkYXRhLmRlcGFydG1lbnRzfHxbXSkuZm9yRWFjaChmdW5jdGlvbihkKXsgbWFw"
    b64 = b64 & "W2QubmFtZV09ZDsgfSk7DQogICAgY29uc3QgZGVwdHMgPSBPUkRFUi5tYXAoZnVuY3Rpb24obil7IHJldHVybiBtYXBbbl0gfHwgeyBuYW1lOm4sIG5vdGljZXM6W10sIHByZXNzOltdIH07IH0pOw0KICAgIGRvY3VtZW50LmdldEVsZW1l"
    b64 = b64 & "bnRCeUlkKCdib2FyZCcpLmlubmVySFRNTCA9IGRlcHRzLm1hcChmdW5jdGlvbihkLCBpZHgpew0KICAgICAgY29uc3QgbSA9IE1FVEFbZC5uYW1lXSB8fCB7IGNvbG9yOicjMDA0YWM2JywgaWNvbjonc3RhcicgfTsNCiAgICAgIGNvbnN0"
    b64 = b64 & "IG5pZCA9ICdub3RpY2UtJyArIGlkeDsNCiAgICAgIGNvbnN0IHBpZCA9ICdwcmVzcy0nICsgaWR4Ow0KICAgICAgcmV0dXJuICc8YXJ0aWNsZSBjbGFzcz0iYmctc3VyZmFjZS1jb250YWluZXItbG93ZXN0IHJvdW5kZWQteGwgc2hhZG93"
    b64 = b64 & "LVswcHhfNHB4XzEycHhfcmdiYSgwLDAsMCwwLjA1KV0gb3ZlcmZsb3ctaGlkZGVuIj4nICsNCiAgICAgICAgJzxoZWFkZXIgY2xhc3M9InAtNCBmbGV4IGp1c3RpZnktYmV0d2VlbiBpdGVtcy1jZW50ZXIiIHN0eWxlPSJiYWNrZ3JvdW5k"
    b64 = b64 & "OicgKyBtLmNvbG9yICsgJyI+JyArDQogICAgICAgICc8aDIgY2xhc3M9ImZvbnQtYm9hcmQtdGl0bGUgdGV4dC1ib2FyZC10aXRsZSIgc3R5bGU9ImNvbG9yOiMzZDRhNWMiPicgKyBlc2MoZC5uYW1lKSArICc8L2gyPicgKw0KICAgICAg"
    b64 = b64 & "ICAnPHNwYW4gY2xhc3M9Im1hdGVyaWFsLXN5bWJvbHMtb3V0bGluZWQiIHN0eWxlPSJmb250LXZhcmlhdGlvbi1zZXR0aW5nczpcJ0ZJTExcJyAxO2NvbG9yOiM1YTY0NzIiPnN0YXI8L3NwYW4+PC9oZWFkZXI+JyArDQogICAgICAgICc8"
    b64 = b64 & "ZGl2IGNsYXNzPSJwLTAiPjxkaXYgY2xhc3M9ImZsZXggYm9yZGVyLWIgYm9yZGVyLW91dGxpbmUtdmFyaWFudCI+JyArDQogICAgICAgICc8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgY2xhc3M9ImZsZXgtMSBweS0zIGZvbnQtYnV0dG9uLXRl"
    b64 = b64 & "eHQgdGV4dC1idXR0b24tdGV4dCBib3JkZXItYi0yIGJvcmRlci1wcmltYXJ5IHRleHQtcHJpbWFyeSIgb25jbGljaz0idG9nZ2xlVGFiKHRoaXMsXCcnICsgbmlkICsgJ1wnLFwnJyArIHBpZCArICdcJykiPuqzteyngOyCrO2VrTwvYnV0"
    b64 = b64 & "dG9uPicgKw0KICAgICAgICAnPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJmbGV4LTEgcHktMyBmb250LWJ1dHRvbi10ZXh0IHRleHQtYnV0dG9uLXRleHQgYm9yZGVyLWItMiBib3JkZXItdHJhbnNwYXJlbnQgdGV4dC1vbi1zdXJm"
    b64 = b64 & "YWNlLXZhcmlhbnQiIG9uY2xpY2s9InRvZ2dsZVRhYih0aGlzLFwnJyArIHBpZCArICdcJyxcJycgKyBuaWQgKyAnXCcpIj7rs7Trj4TsnpDro4w8L2J1dHRvbj4nICsNCiAgICAgICAgJzwvZGl2PicgKw0KICAgICAgICAnPHVsIGNsYXNz"
    b64 = b64 & "PSJkaXZpZGUteSBkaXZpZGUtcHJpbWFyeS8xMCBwLTIgYmxvY2siIGlkPSInICsgbmlkICsgJyI+JyArIGl0ZW1zSHRtbChkLm5vdGljZXMsICd0ZXh0LXByaW1hcnknKSArICc8L3VsPicgKw0KICAgICAgICAnPHVsIGNsYXNzPSJkaXZp"
    b64 = b64 & "ZGUteSBkaXZpZGUtcHJpbWFyeS8xMCBwLTIgaGlkZGVuIiBpZD0iJyArIHBpZCArICciPicgKyBpdGVtc0h0bWwoZC5wcmVzcywgJ3RleHQtc2Vjb25kYXJ5JykgKyAnPC91bD4nICsNCiAgICAgICAgJzwvZGl2PjwvYXJ0aWNsZT4nOw0K"
    b64 = b64 & "ICAgIH0pLmpvaW4oJycpOw0KICB9DQogIGZ1bmN0aW9uIGxvYWQoKXsNCiAgICBjb25zdCByYXcgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWFqb3ItZGF0YScpLnRleHRDb250ZW50LnRyaW0oKTsNCiAgICBsZXQgZGF0YSA9IHsg"
    b64 = b64 & "Y29sbGVjdGVkQXQ6JycsIGRlcGFydG1lbnRzOltdIH07DQogICAgdHJ5IHsgaWYgKHJhdyAmJiByYXcuY2hhckF0KDApPT09J3snKSBkYXRhID0gSlNPTi5wYXJzZShyYXcpOyB9IGNhdGNoKGUpeyBjb25zb2xlLmVycm9yKGUpOyB9DQog"
    b64 = b64 & "ICAgcmVuZGVyKGRhdGEpOw0KICB9DQogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdidG5SZWxvYWQnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGZ1bmN0aW9uKCl7IGxvY2F0aW9uLnJlbG9hZCgpOyB9KTsNCiAgbG9hZCgpOw0K"
    b64 = b64 & "fSkoKTsNCjwvc2NyaXB0Pg0KPC9ib2R5Pg0KPC9odG1sPg0K"
    EmbeddedMajorPressTemplateMobile = DecodeBase64Utf8(b64)
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
