Attribute VB_Name = "Plugin_WebSearch"
Option Explicit

' Plugin_WebSearch: 要約なしの軽いWeb検索 (DuckDuckGo HTMLエンドポイント)。
' 内蔵 web_fetch は特定URLを読むのに対し、これはキーワードで検索結果一覧を返す。
' ネットワーク取得は MSXML2.ServerXMLHTTP.6.0 で同期GETする(自己テストのみ許可)。

Public Function Plugin_Name() As String
    Plugin_Name = "websearch"
End Function

Public Function Plugin_ToolNames() As String
    Plugin_ToolNames = "web_search"
End Function

Public Function Plugin_WriteTools() As String
    Plugin_WriteTools = ""
End Function

Public Function Plugin_ToolsOpenAI() As String
    Plugin_ToolsOpenAI = ToolDef("web_search", "Search the public web by keyword and return a plain-text list of results (title, URL, snippet) without any summary. Use this for current facts, news, or anything you do not already know.", _
        "{""query"":{""type"":""string"",""description"":""Search keywords""}," & _
        """top"":{""type"":""integer"",""description"":""Number of results to return, default 5, max 10""}}")
End Function

Public Function Plugin_ToolsAnthropic() As String
    Plugin_ToolsAnthropic = AnthropicToolDef("web_search", "Search the public web by keyword and return a plain-text list of results (title, URL, snippet) without any summary. Use this for current facts, news, or anything you do not already know.", _
        "{""query"":{""type"":""string"",""description"":""Search keywords""}," & _
        """top"":{""type"":""integer"",""description"":""Number of results to return, default 5, max 10""}}")
End Function

Public Function Plugin_Execute(ByVal toolName As String, ByVal argsJson As String) As String
    If LCase$(toolName) <> "web_search" Then
        Plugin_Execute = "Plugin_WebSearch: 未知のツール " & toolName
        Exit Function
    End If
    Plugin_Execute = WebSearch_Run(JsonGetArg(argsJson, "query"), JsonGetArg(argsJson, "top"))
End Function

' SelfTest: 実際に DuckDuckGo から q=excel を取得できたら OK。
' (自己テストでのネットワークアクセスはここでだけ許可)
Public Function Plugin_SelfTest() As String
    Dim res As String
    res = HttpGet("https://html.duckduckgo.com/html/?q=excel")
    If InStr(1, res, "result") > 0 Or InStr(1, res, "duckduckgo") > 0 Then
        Plugin_SelfTest = "OK"
    ElseIf Left$(res, 1) = "E" Then
        Plugin_SelfTest = res
    Else
        Plugin_SelfTest = "応答に検索結果が含まれない"
    End If
End Function

Private Function WebSearch_Run(ByVal query As String, ByVal top As String) As String
    Dim q As String
    Dim t As Long
    Dim html As String

    q = Trim$(query)
    If Len(q) = 0 Then
        WebSearch_Run = "検索キーワード(query)を指定してください。"
        Exit Function
    End If

    On Error Resume Next
    t = CLng(top)
    If Err.Number <> 0 Then t = 5
    Err.Clear
    On Error GoTo 0
    If t < 1 Then t = 1
    If t > 10 Then t = 10

    html = HttpGet("https://html.duckduckgo.com/html/?q=" & UrlEncode(q))
    If Left$(html, 1) = "E" Then
        WebSearch_Run = html
        Exit Function
    End If

    WebSearch_Run = ParseResults(html, t)
End Function

Private Function HttpGet(ByVal url As String) As String
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 10000, 30000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.send ""
    If http.Status <> 200 Then
        HttpGet = "Web検索に失敗しました: HTTP " & http.Status
        Exit Function
    End If
    HttpGet = http.responseText
    Exit Function
Fail:
    HttpGet = "Web検索に失敗しました: " & Err.Description
End Function

' HTMLからタイトル/URL/スニペットを取り出し、1結果=1行に整形して返す。
' 8000文字で切る。HTMLタグはRegExpで除去。
Private Function ParseResults(ByVal html As String, ByVal top As Long) As String
    Dim re As Object
    Dim m As Object
    Dim i As Long
    Dim title As String
    Dim href As String
    Dim snip As String
    Dim out As String
    Dim line As String

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' DuckDuckGo HTMLは 結果見出しが <a class="result__a" href="...">タイトル</a>
    re.Pattern = "class=""result__a"" href=""([^""]+)""[^>]*>(.*?)</a>"
    If re.Test(html) Then
        Set m = re.Execute(html)
        For i = 0 To m.Count - 1
            If i >= top Then Exit For
            href = CleanTags(m.Item(i).SubMatches(0))
            title = CleanTags(m.Item(i).SubMatches(1))
            snip = ExtractSnippet(html, m.Item(i).FirstIndex)
            line = title & " | " & href
            If Len(snip) > 0 Then line = line & " | " & snip
            out = out & IIf(Len(out) = 0, "", vbLf) & line
        Next i
    End If

    If Len(out) = 0 Then
        out = "検索結果がありません(キーワードを変えるかサイト側の応答を確認)。"
    End If
    If Len(out) > 8000 Then out = Left$(out, 8000)
    ParseResults = out
End Function

' 与えられた見出し位置の直後にある result__snippet を取得する
Private Function ExtractSnippet(ByVal html As String, ByVal startAt As Long) As String
    Dim re As Object
    Dim m As Object
    Dim chunk As String

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "class=""result__snippet""[^>]*>(.*?)</a>"
    chunk = Mid$(html, startAt + 1, 6000)
    If re.Test(chunk) Then
        Set m = re.Execute(chunk)
        If m.Count > 0 Then ExtractSnippet = CleanTags(m.Item(0).SubMatches(0))
    End If
End Function

' タグ除去・文字参照を緩和し、余分な空白を潰す
Private Function CleanTags(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "<[^>]+>"
    s = re.Replace(s, "")
    re.Pattern = "&nbsp;"
    s = re.Replace(s, " ")
    re.Pattern = "&amp;"
    s = re.Replace(s, "&")
    re.Pattern = "&lt;"
    s = re.Replace(s, "<")
    re.Pattern = "&gt;"
    s = re.Replace(s, ">")
    re.Pattern = "&quot;"
    s = re.Replace(s, """")
    re.Pattern = "\s+"
    s = re.Replace(s, " ")
    CleanTags = Trim$(s)
End Function

Private Function UrlEncode(ByVal s As String) As String
    Dim stream As Object
    Dim bytes() As Byte
    Dim i As Long
    Dim b As Long
    Dim c As String
    Dim out As String
    Dim ch As String

    ' ASCIIセーフ文字はそのまま、それ以外はUTF-8バイトを%XXで符号化する
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-zA-Z0-9._~-]" Then
            out = out & ch
        Else
            out = out & EncodeUtf8Byte(ch)
        End If
    Next i
    UrlEncode = out
End Function

Private Function EncodeUtf8Byte(ByVal c As String) As String
    Dim stream As Object
    Dim bytes() As Byte
    Dim i As Long
    Dim out As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText c
    stream.Position = 0
    stream.Type = 1
    bytes = stream.Read
    stream.Close
    For i = 0 To UBound(bytes)
        out = out & "%" & Right$("0" & Hex$(bytes(i)), 2)
    Next i
    EncodeUtf8Byte = out
End Function