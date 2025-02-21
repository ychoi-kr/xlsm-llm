Option Explicit

Const BASE_URL_DEFAULT As String = "https://api.openai.com/v1/"
Const DEFAULT_MODEL As String = "gpt-4o-mini"

' 공통 인증 로직: API 키가 전달되면 HTTP 요청 헤더에 추가
Private Sub SetAuthorizationHeader(ByRef http As Object, Optional apiKey As Variant)
    If Not IsMissing(apiKey) And Not IsEmpty(apiKey) Then
        If CStr(apiKey) <> "" Then
            http.setRequestHeader "Authorization", "Bearer " & CStr(apiKey)
        End If
    End If
End Sub

Private Function BuildJsonPayload(ByVal modelName As String, ByVal fullPrompt As String, _
                                    Optional ByVal temperature As Variant, Optional ByVal maxTokens As Variant) As String
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
    
    ' 최대 토큰 (maxTokens) 추가
    If Not IsMissing(maxTokens) And Not IsEmpty(maxTokens) Then
        If IsNumeric(maxTokens) Then
            jsonPayload = jsonPayload & ", ""max_tokens"": " & maxTokens
        End If
    End If
    
    jsonPayload = jsonPayload & "}"
    BuildJsonPayload = jsonPayload
End Function

Private Function SendLLMRequest(ByVal url As String, ByVal jsonPayload As String, Optional apiKey As Variant) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo ErrorHandler
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    ' API 키가 제공되면 인증 헤더 추가
    SetAuthorizationHeader http, apiKey
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
    
    If matches.count > 0 Then
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


Function LLM_Base(prompt As String, Optional temperature As Variant, Optional maxTokens As Variant, _
                  Optional model As Variant, Optional baseUrl As Variant, Optional apiKey As Variant) As String
    Dim url As String
    If Not IsMissing(baseUrl) And Not IsEmpty(baseUrl) Then
        Dim baseStr As String
        baseStr = CStr(baseUrl)
        If Right(baseStr, 1) <> "/" Then baseStr = baseStr & "/"
        url = baseStr & "chat/completions"
    Else
        url = BASE_URL_DEFAULT
        If Right(url, 1) <> "/" Then url = url & "/"
        url = url & "chat/completions"
    End If
    
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_MODEL
    Else
        modelName = CStr(model)
    End If
    
    ' API 키 처리: 각 서비스별로 필수 키 확인
    Dim finalApiKey As String
    Dim lowerUrl As String
    lowerUrl = LCase(url)
    
    If IsMissing(apiKey) Or IsEmpty(apiKey) Then
        If InStr(lowerUrl, "gemini") > 0 Or InStr(LCase(modelName), "gemini") > 0 Then
            finalApiKey = Environ("GEMINI_API_KEY")
            If finalApiKey = "" Then
                LLM_Base = "Error: Gemini API requires an API key. Provide it as the last argument or set the GEMINI_API_KEY environment variable."
                Exit Function
            End If
        ElseIf InStr(lowerUrl, "openai.com") > 0 Then
            finalApiKey = Environ("OPENAI_API_KEY")
            If finalApiKey = "" Then
                LLM_Base = "Error: OpenAI API requires an API key. Provide it as the last argument or set the OPENAI_API_KEY environment variable."
                Exit Function
            End If
        ElseIf InStr(lowerUrl, "upstage.ai") > 0 Or InStr(LCase(modelName), "upstage") > 0 Then
            finalApiKey = Environ("UPSTAGE_API_KEY")
            If finalApiKey = "" Then
                LLM_Base = "Error: Upstage API requires an API key. Provide it as the last argument or set the UPSTAGE_API_KEY environment variable."
                Exit Function
            End If
        Else
            ' 기본적으로 로컬 서버로 가정, 키 없이 진행
            finalApiKey = ""
        End If
    Else
        finalApiKey = CStr(apiKey)
    End If
    
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload(modelName, EscapeText(prompt), temperature, maxTokens)
    
    Dim response As String
    response = SendLLMRequest(url, jsonPayload, finalApiKey)
    
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
             Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
             Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    Dim fullPrompt As String
    fullPrompt = prompt
    If value <> "" Then fullPrompt = fullPrompt & " " & value

    Dim response As String
    response = LLM_Base(fullPrompt, temperature, maxTokens, model, baseUrl, apiKey)
    LLM = ProcessLLMResponse(response, showThink)
End Function

Function LLM_SUMMARIZE(text As String, Optional prompt As String, Optional temperature As Variant, _
                     Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                     Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    If prompt = "" Then
        prompt = "Summarize in one line:"
    End If
    Dim fullPrompt As String
    fullPrompt = prompt & " " & text
    Dim response As String
    response = LLM_Base(fullPrompt, temperature, maxTokens, model, baseUrl, apiKey)
    LLM_SUMMARIZE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_CODE(programDetails As String, programmingLanguage As String, _
                  Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    Dim prompt As String
    prompt = "Generate a " & programmingLanguage & " program that fulfills the following requirements:" & vbCrLf & programDetails
    Dim response As String
    response = LLM_Base(prompt, 0.2, , model, baseUrl, apiKey)
    LLM_CODE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_LIST(prompt As String, Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    Dim listPrompt As String
    ' 영어 프롬프트에 출력 형식 예시를 포함하여, 모델이 <list>와 <item> 태그를 사용해 출력하도록 명시합니다.
    listPrompt = prompt & vbCrLf & _
                 "Example:" & vbCrLf & _
                 "<list><item>Apple</item><item>Banana</item><item>Cherry</item></list>" & vbCrLf & _
                 "Please output only the list items enclosed within <list> and <item> tags, exactly in the above format, with no additional commentary."
    
    Dim response As String
    response = LLM_Base(listPrompt, , , model, baseUrl, apiKey)
    
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
                  Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 기본 프롬프트 설정: 사용자가 prompt를 입력하지 않은 경우
    If prompt = "" Then
        prompt = "Please correct the following sentence for clarity, grammar, and punctuation:"
    End If
    
    ' 입력 문장과 프롬프트를 결합하여 전체 요청 문장을 구성합니다.
    Dim fullPrompt As String
    fullPrompt = prompt & " " & text
    
    ' LLM에게 요청 전송
    Dim response As String
    response = LLM_Base(fullPrompt, temperature, maxTokens, model, baseUrl, apiKey)
    
    ' 응답을 파싱하여 최종 결과를 반환합니다.
    LLM_EDIT = ProcessLLMResponse(response, showThink)
End Function

Function LLM_TRANSLATE(text As String, Optional targetLang As String = "", Optional sourceLang As String = "", _
                       Optional customPrompt As String = "", Optional temperature As Variant, _
                       Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                       Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    Dim finalPrompt As String
    Dim effectiveBaseUrl As String
    Dim effectiveModel As String
    
    ' baseUrl과 model은 생략 시 LLM_Base에서 기본값 처리
    If Not IsMissing(baseUrl) And Not IsEmpty(baseUrl) Then effectiveBaseUrl = CStr(baseUrl)
    If Not IsMissing(model) And Not IsEmpty(model) Then effectiveModel = CStr(model)
    
    ' 실제 사용될 baseUrl과 model 결정
    Dim resolvedBaseUrl As String
    Dim resolvedModel As String
    resolvedBaseUrl = IIf(effectiveBaseUrl = "", BASE_URL_DEFAULT, effectiveBaseUrl)
    resolvedModel = IIf(effectiveModel = "", DEFAULT_MODEL, effectiveModel)
    
    ' Upstage 번역 모델 체크
    If LCase(resolvedBaseUrl) = "https://api.upstage.ai/v1/solar" And _
       (LCase(resolvedModel) = "translation-enko" Or LCase(resolvedModel) = "translation-koen") Then
        finalPrompt = text ' 프롬프트 없이 text만 사용
    Else
        ' 다른 모델의 경우 targetLang 또는 customPrompt 중 하나가 필요
        If targetLang = "" And customPrompt = "" Then
            LLM_TRANSLATE = "Error: Either targetLang or customPrompt must be provided for non-Upstage translation models"
            Exit Function
        End If
        
        ' 프롬프트 생성: customPrompt가 있으면 text와 조합, 없으면 기본 프롬프트
        If customPrompt <> "" Then
            finalPrompt = customPrompt & " " & text
        Else
            finalPrompt = IIf(sourceLang <> "", "Translate the following text from " & sourceLang & " to " & targetLang & ": ", _
                              "Translate the following text to " & targetLang & ": ") & text
        End If
    End If

    Dim response As String
    response = LLM_Base(finalPrompt, temperature, maxTokens, model, baseUrl, apiKey)
    LLM_TRANSLATE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_REVIEW_TRANSLATION(originalText As String, translatedText As String, _
                                Optional focus As String = "", _
                                Optional temperature As Variant, Optional maxTokens As Variant, _
                                Optional model As Variant, Optional baseUrl As Variant, _
                                Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    Dim fullPrompt As String
    
    ' 기본 프롬프트: 주안점이 없는 경우 균형 잡힌 감수 요청
    If focus = "" Then
        fullPrompt = "Review the following translation for accuracy, grammar, fluency, and overall quality. " & _
                     "Provide feedback and suggest improvements if necessary." & vbCrLf & _
                     "Original text: " & originalText & vbCrLf & _
                     "Translated text: " & translatedText
    Else
        ' 주안점이 있는 경우, 사용자가 지정한 초점에 맞춘 프롬프트
        fullPrompt = "Review the following translation with a focus on " & focus & ". " & _
                     "Provide feedback and suggest improvements if necessary." & vbCrLf & _
                     "Original text: " & originalText & vbCrLf & _
                     "Translated text: " & translatedText
    End If
    
    ' LLM 요청 전송
    Dim response As String
    response = LLM_Base(fullPrompt, temperature, maxTokens, model, baseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_REVIEW_TRANSLATION = ProcessLLMResponse(response, showThink)
End Function
