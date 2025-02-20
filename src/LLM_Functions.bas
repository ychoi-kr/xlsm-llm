Option Explicit

Const SERVER_URL As String = "http://localhost:1234/v1/chat/completions"
Const DEFAULT_MODEL As String = "exaone-3.5-7.8b-instruct"

Private Function BuildJsonPayload(ByVal modelName As String, ByVal fullPrompt As String, _
                                    Optional ByVal temperature As Variant, Optional ByVal max_tokens As Variant) As String
    Dim jsonPayload As String
    jsonPayload = "{" & _
                  """model"": """ & modelName & """," & _
                  """messages"": [{" & _
                  """role"": ""user""," & _
                  """content"": """ & Replace(fullPrompt, """", "\""") & """" & _
                  "}]"

    ' 온도 (temperature) 추가
    If Not IsMissing(temperature) And Not IsEmpty(temperature) Then
        If IsNumeric(temperature) Then
            jsonPayload = jsonPayload & ", ""temperature"": " & temperature
        End If
    End If

    ' 최대 토큰 (max_tokens) 추가
    If Not IsMissing(max_tokens) And Not IsEmpty(max_tokens) Then
        If IsNumeric(max_tokens) Then
            jsonPayload = jsonPayload & ", ""max_tokens"": " & max_tokens
        End If
    End If

    jsonPayload = jsonPayload & "}"
    BuildJsonPayload = jsonPayload
End Function

Private Function SendLLMRequest(ByVal url As String, ByVal jsonPayload As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo ErrorHandler
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    ' 응답 처리
    If http.Status = 200 Then
        SendLLMRequest = http.responseText
    Else
        Dim serverMsg As String
        serverMsg = http.responseText
        If serverMsg <> "" Then
            SendLLMRequest = "Error: " & http.Status & " " & http.statusText & " - " & serverMsg
        Else
            SendLLMRequest = "Error: " & http.Status & " " & http.statusText
        End If
    End If
    
    On Error GoTo 0
    Set http = Nothing
    Exit Function

ErrorHandler:
    Dim errMsg As String
    errMsg = "Error: " & Err.Number & " " & Err.Description
    If Err.Number = 12029 Then
        errMsg = errMsg & " - Please ensure the correct URL is specified and the server is accessible on the network."
    End If
    SendLLMRequest = errMsg
    On Error GoTo 0
    Set http = Nothing
End Function


Private Function ExtractContent(ByVal response As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Pattern = """content"":\s*""([\s\S]*?)""\s*(?:,|\})"
    regEx.IgnoreCase = True
    regEx.Global = False
    
    Dim matches As Object
    Set matches = regEx.Execute(response)
    
    If matches.Count > 0 Then
        ExtractContent = Replace(matches(0).SubMatches(0), "\""", """")
    Else
        ExtractContent = "Error: Failed to parse response"
    End If
End Function

Private Function EscapeText(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    EscapeText = result
End Function

Private Function UnescapeText(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "\n", vbLf)
    result = Replace(result, "\\", "\")
    UnescapeText = result
End Function


Function LLM_Base(prompt As String, Optional value As String = "", Optional temperature As Variant, _
                  Optional max_tokens As Variant, Optional model As Variant, Optional base_url As Variant) As String
    Dim url As String
    If Not IsMissing(base_url) And Not IsEmpty(base_url) Then
        url = CStr(base_url) & "/v1/chat/completions"
    Else
        url = SERVER_URL
    End If
    
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_MODEL
    Else
        modelName = CStr(model)
    End If
    
    Dim fullPrompt As String
    If value = "" Then
        fullPrompt = prompt
    Else
        fullPrompt = prompt & " " & value
    End If
    
    fullPrompt = EscapeText(fullPrompt)
    
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload(modelName, fullPrompt, temperature, max_tokens)
    
    Dim response As String
    response = SendLLMRequest(url, jsonPayload)
    
    If Left(response, 6) = "Error:" Then
        LLM_Base = response
        Exit Function
    End If
    
    LLM_Base = UnescapeText(ExtractContent(response))
End Function

' Private 함수: 문자열 앞부분에 있는 모든 줄 바꿈(CR, LF) 제거
Private Function RemoveLeadingLineBreaks(ByVal text As String) As String
    Do While Len(text) > 0
        If Left(text, 1) = vbCr Or Left(text, 1) = vbLf Then
            text = Mid(text, 2)
        Else
            Exit Do
        End If
    Loop
    RemoveLeadingLineBreaks = text
End Function

' Private 함수: 응답을 처리하여 showThink 옵션에 따라 결과를 분리 반환
Private Function ProcessLLMResponse(response As String, Optional showThink As Boolean = False) As Variant
    Dim thinkStart As Long, thinkEnd As Long
    Dim thinkContent As String, remainingContent As String
    
    thinkStart = InStr(1, response, "<think>")
    thinkEnd = InStr(1, response, "</think>")
    
    If thinkStart > 0 And thinkEnd > 0 Then
        ' <think> 태그 내부의 내용 추출
        thinkContent = Mid(response, thinkStart + Len("<think>"), thinkEnd - thinkStart - Len("<think>"))
        ' <think> 태그 전체를 제거하여 나머지 내용 확보 후 Trim으로 양쪽 공백 제거
        remainingContent = Trim(Replace(response, Mid(response, thinkStart, thinkEnd - thinkStart + Len("</think>")), ""))
    Else
        thinkContent = ""
        remainingContent = response
    End If
    
    ' 두 결과 모두 선행 줄 바꿈 제거
    thinkContent = RemoveLeadingLineBreaks(thinkContent)
    remainingContent = RemoveLeadingLineBreaks(remainingContent)
    
    If showThink Then
        ProcessLLMResponse = Array(thinkContent, remainingContent)
    Else
        ProcessLLMResponse = remainingContent
    End If
End Function

Function LLM(prompt As String, Optional value As String = "", Optional temperature As Variant, _
             Optional max_tokens As Variant, Optional model As Variant, Optional base_url As Variant, _
             Optional showThink As Boolean = False) As Variant
    Dim response As String
    response = LLM_Base(prompt, value, temperature, max_tokens, model, base_url)
    LLM = ProcessLLMResponse(response, showThink)
End Function

Function LLM_SUMMARIZE(text As String, Optional prompt As String, _
                     Optional temperature As Variant, Optional max_tokens As Variant, _
                     Optional model As Variant, Optional base_url As Variant, _
                     Optional showThink As Boolean = False) As Variant
    If prompt = "" Then
        prompt = "Summarize in one line:"
    End If
    Dim fullPrompt As String
    fullPrompt = prompt & " " & text
    Dim response As String
    response = LLM_Base(fullPrompt, "", temperature, max_tokens, model, base_url)
    LLM_SUMMARIZE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_CODE(programDetails As String, programmingLanguage As String, _
                  Optional model As Variant, Optional base_url As Variant, _
                  Optional showThink As Boolean = False) As Variant
    Dim prompt As String
    prompt = "Generate a " & programmingLanguage & " program that fulfills the following requirements:" & vbLf & programDetails
    Dim response As String
    response = LLM_Base(prompt, "", 0.2, , model, base_url)
    LLM_CODE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_LIST(prompt As String, Optional model As Variant, Optional base_url As Variant, Optional showThink As Boolean = False) As Variant
    Dim listPrompt As String
    ' 영어 프롬프트에 출력 형식 예시를 포함하여, 모델이 <list>와 <item> 태그를 사용해 출력하도록 명시합니다.
    listPrompt = prompt & vbCrLf & _
                 "Example:" & vbCrLf & _
                 "<list><item>Apple</item><item>Banana</item><item>Cherry</item></list>" & vbCrLf & _
                 "Please output only the list items enclosed within <list> and <item> tags, exactly in the above format, with no additional commentary."
    
    Dim response As String
    response = LLM_Base(listPrompt, "", , , model, base_url)
    
    Dim processedResponse As Variant
    processedResponse = ProcessLLMResponse(response, showThink)
    
    Dim contentText As String, thinkText As String
    If showThink Then
        thinkText = processedResponse(0)
        contentText = processedResponse(1)
    Else
        contentText = processedResponse
    End If
    
    ' <item> 태그 사이의 리스트 항목들을 추출
    Dim items() As String
    Dim itemCount As Long
    itemCount = 0
    Dim searchPos As Long, startPos As Long, endPos As Long, currentItem As String
    searchPos = 1
    Do
        startPos = InStr(searchPos, contentText, "<item>")
        If startPos = 0 Then Exit Do
        endPos = InStr(startPos, contentText, "</item>")
        If endPos = 0 Then Exit Do
        currentItem = Mid(contentText, startPos + Len("<item>"), endPos - startPos - Len("<item>"))
        currentItem = Trim(currentItem)
        ReDim Preserve items(itemCount)
        items(itemCount) = currentItem
        itemCount = itemCount + 1
        searchPos = endPos + Len("</item>")
    Loop
    
    ' showThink 옵션에 따라 결과 배열 반환:
    ' - showThink가 False이면, 순수 리스트 배열만 반환
    ' - showThink가 True이면, 첫 번째 요소는 think 내용, 두 번째 요소는 리스트 배열 반환
    If showThink Then
        Dim resultArray(1) As Variant
        resultArray(0) = thinkText
        resultArray(1) = items
        LLM_LIST = resultArray
    Else
        LLM_LIST = items
    End If
End Function

Function LLM_EDIT(text As String, Optional prompt As String, Optional temperature As Variant, _
                  Optional max_tokens As Variant, Optional model As Variant, Optional base_url As Variant, _
                  Optional showThink As Boolean = False) As Variant
    ' 기본 프롬프트 설정: 사용자가 prompt를 입력하지 않은 경우
    If prompt = "" Then
        prompt = "Please correct the following sentence for clarity, grammar, and punctuation:"
    End If
    
    ' 입력 문장과 프롬프트를 결합하여 전체 요청 문장을 구성합니다.
    Dim fullPrompt As String
    fullPrompt = prompt & " " & text
    
    ' LLM에게 요청 전송
    Dim response As String
    response = LLM_Base(fullPrompt, "", temperature, max_tokens, model, base_url)
    
    ' 응답을 파싱하여 최종 결과를 반환합니다.
    LLM_EDIT = ProcessLLMResponse(response, showThink)
End Function
