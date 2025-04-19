Option Explicit

' 상수 정의
Const BASE_URL_DEFAULT As String = "https://api.openai.com/v1/"
Const ANTHROPIC_URL As String = "https://api.anthropic.com/v1/messages"
Const UPSTAGE_URL As String = "https://api.upstage.ai/v1/solar/" ' Added Upstage URL constant
Const DEFAULT_MODEL As String = "gpt-4o-mini"
Const DEFAULT_ANTHROPIC_MODEL As String = "claude-3-5-sonnet-20240620"

'===============================================================
' 유틸리티 함수 섹션
'===============================================================

' API 요청 헤더에 인증 정보 추가
Private Sub SetAuthorizationHeader(ByRef http As Object, Optional apiKey As Variant, Optional isAnthropicAPI As Boolean = False)
    If Not IsMissing(apiKey) And Not IsEmpty(apiKey) Then
        If CStr(apiKey) <> "" Then
            If isAnthropicAPI Then
                ' Anthropic API는 x-api-key 헤더 사용
                http.setRequestHeader "x-api-key", CStr(apiKey)
                http.setRequestHeader "anthropic-version", "2023-06-01"
            Else
                ' 다른 API는 Bearer 토큰 사용
                http.setRequestHeader "Authorization", "Bearer " & CStr(apiKey)
            End If
        End If
    End If
End Sub

' URL이 특정 서비스인지 확인하는 함수들
Private Function isAnthropicUrl(ByVal url As String) As Boolean
    isAnthropicUrl = (InStr(LCase(url), "anthropic") > 0)
End Function

Function LLM_EDIT(text As String, Optional prompt As String = "", Optional temperature As Variant, _
                  Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant, Optional includeReview As Boolean = False) As Variant
    
    ' Ensure text is not empty
    If Trim(text) = "" Then
        LLM_EDIT = "Error: Input text cannot be empty"
        Exit Function
    End If
    
    ' 모델 이름 및 URL 설정
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    Dim systemPrompt As String
    Dim userPromptText As String
    
    ' Set appropriate prompts based on model type and user input
    If prompt = "" Then
        ' Default prompts when no prompt is provided
        If InStr(1, LCase(modelName), "solar-") > 0 Then
            ' For Upstage Solar models
            systemPrompt = "Generate proofreading results for the input document."
            userPromptText = text
        Else
            ' For OpenAI/other models - use the original prompt structure from your code
            systemPrompt = "You are a helpful assistant that corrects text for grammar, spelling, punctuation, and clarity."
            
            ' Use the original detailed prompt with all five important points
            userPromptText = "Please correct the following text:" & vbCrLf & vbCrLf & _
                text & vbCrLf & vbCrLf & _
                "Use these exact delimiters in your response:" & vbCrLf & _
                "===REVIEW START===" & vbCrLf & _
                "[Your explanation of corrections made]" & vbCrLf & _
                "===REVIEW END===" & vbCrLf & _
                "===RESULT START===" & vbCrLf & _
                "[Corrected text only]" & vbCrLf & _
                "===RESULT END===" & vbCrLf & vbCrLf & _
                "IMPORTANT:" & vbCrLf & _
                "1. Do NOT translate the text to another language." & vbCrLf & _
                "2. Only correct grammar, spelling, punctuation, and clarity in the original language." & vbCrLf & _
                "3. Use EXACTLY the format above with the exact delimiters shown." & vbCrLf & _
                "4. Do not include any code blocks, backticks, or markdown formatting." & vbCrLf & _
                "5. Provide your review comments in the SAME LANGUAGE as the original text."
        End If
    Else
        ' User provided a custom prompt - use that as system prompt
        systemPrompt = prompt
        
        ' If no custom prompt was provided, use a standard format
        ' Otherwise, for custom prompts, don't add any special formatting
        userPromptText = text
    End If
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    If IsAnthropicModelOrUrl(modelName, effectiveBaseUrl) Then
        ' For Anthropic models
        Dim anthropicPrompt As String
        If systemPrompt <> "" Then
            ' Include system instruction in the user prompt for Anthropic
            anthropicPrompt = "System instruction: " & systemPrompt & vbCrLf & vbCrLf & "User request: " & userPromptText
        Else
            anthropicPrompt = userPromptText
        End If
        response = LLM_Base_Anthropic("", anthropicPrompt, temperature, maxTokens, modelName, apiKey)
    Else
        ' For OpenAI and other models
        response = LLM_Base_OpenAI(systemPrompt, userPromptText, temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    End If
    
    ' Clean any markdown code blocks from the response
    response = CleanTextFromMarkdown(response)
    
    ' Process the response
    If includeReview Then
        ' Extract review and result sections
        Dim reviewSection As String, resultSection As String
        reviewSection = ExtractContentBetweenDelimiters(response, "===REVIEW START===", "===REVIEW END===", "")
        resultSection = ExtractContentBetweenDelimiters(response, "===RESULT START===", "===RESULT END===", "")
        
        ' If sections weren't found, try to intelligently split the response
        If resultSection = "" Then
            ' Default to the whole response as the result
            resultSection = response
            reviewSection = ""
        End If
        
        ' Return both sections as an array
        If showThink Then
            ' If showThink is true, process for think tags
            Dim thinkResult As Variant
            thinkResult = ProcessLLMResponse(resultSection, True)
            
            Dim finalResult(2) As Variant
            finalResult(0) = thinkResult(0) ' think content
            finalResult(1) = thinkResult(1) ' result content
            finalResult(2) = reviewSection  ' review content
            LLM_EDIT = finalResult
        Else
            ' Standard case without think tags
            Dim result(1) As Variant
            result(0) = resultSection
            result(1) = reviewSection
            LLM_EDIT = result
        End If
    Else
        ' No review requested, just process the standard response
        If prompt = "" And Not (InStr(1, LCase(modelName), "solar-") > 0) Then
            ' For default prompts, still extract the result section
            Dim resultOnly As String
            resultOnly = ExtractContentBetweenDelimiters(response, "===RESULT START===", "===RESULT END===", "")
            
            ' If extraction failed, use the whole response
            If resultOnly = "" Then
                resultOnly = response
            End If
            
            LLM_EDIT = ProcessLLMResponse(resultOnly, showThink)
        Else
            ' For custom prompts or Solar models, use the full response
            LLM_EDIT = ProcessLLMResponse(response, showThink)
        End If
    End If
End Function

Function LLM_TRANSLATE(text As String, Optional targetLang As String = "", Optional sourceLang As String = "", _
                       Optional customPrompt As String = "", Optional temperature As Variant, _
                       Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                       Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 모델 이름 및 URL 설정
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    ' 프롬프트 구성
    Dim finalPrompt As String
    
    ' Upstage 번역 모델 체크
    If (Left(LCase(modelName), 6) = "solar-" Or IsUpstageUrl(effectiveBaseUrl)) And _
       (LCase(modelName) = "translation-enko" Or LCase(modelName) = "translation-koen") Then
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
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    response = LLM_Dispatcher(finalPrompt, "", temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_TRANSLATE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_REVIEW_TRANSLATION(originalText As String, translatedText As String, _
                                Optional focus As String = "", _
                                Optional temperature As Variant, Optional maxTokens As Variant, _
                                Optional model As Variant, Optional baseUrl As Variant, _
                                Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 모델 이름 및 URL 설정
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    ' 프롬프트 구성
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
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    response = LLM_Dispatcher(fullPrompt, "", temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_REVIEW_TRANSLATION = ProcessLLMResponse(response, showThink)
End Function

Function LLM_Advanced(Optional systemPrompt As String = "", Optional userPrompt As String = "", _
                     Optional temperature As Variant, Optional maxTokens As Variant, _
                     Optional model As Variant, Optional baseUrl As Variant, _
                     Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 입력 검증: userPrompt는 필수
    If userPrompt = "" Then
        LLM_Advanced = "Error: userPrompt is required for LLM_Advanced"
        Exit Function
    End If
    
    ' 모델 이름 및 URL 설정
    Dim modelName As String, effectiveBaseUrl As String
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    ' LLM 요청 및 처리
    Dim response As String
    response = LLM_Dispatcher(userPrompt, systemPrompt, temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_Advanced = ProcessLLMResponse(response, showThink)
End Function

Private Function IsOpenAiUrl(ByVal url As String) As Boolean
    IsOpenAiUrl = (InStr(LCase(url), "openai.com") > 0)
End Function

Private Function IsUpstageUrl(ByVal url As String) As Boolean
    IsUpstageUrl = (InStr(LCase(url), "upstage.ai") > 0)
End Function

Private Function IsGeminiUrl(ByVal url As String) As Boolean
    IsGeminiUrl = (InStr(LCase(url), "gemini") > 0)
End Function

Private Function IsLocalLLM(ByVal url As String) As Boolean
    Dim lowerUrl As String
    lowerUrl = LCase(url)
    IsLocalLLM = (InStr(lowerUrl, "localhost") > 0 Or InStr(lowerUrl, "127.0.0.1") > 0)
End Function

' 모델명으로 baseUrl 자동 결정 함수
Private Function GetBaseUrlFromModel(ByVal modelName As String, Optional ByVal defaultUrl As String = "") As String
    Dim lowerModel As String
    lowerModel = LCase(modelName)
    
    If Left(lowerModel, 6) = "solar-" Then
        ' For Solar models, return Upstage base URL without "solar/" appended
        ' This will be handled in ResolveModelAndBaseUrl
        GetBaseUrlFromModel = "https://api.upstage.ai/v1/"
    ElseIf InStr(lowerModel, "claude") > 0 Then
        ' Claude model returns Anthropic URL
        GetBaseUrlFromModel = ANTHROPIC_URL
    ElseIf defaultUrl <> "" Then
        ' Default URL provided
        GetBaseUrlFromModel = defaultUrl
    Else
        ' Otherwise, return default OpenAI URL
        GetBaseUrlFromModel = BASE_URL_DEFAULT
    End If
End Function

' API 키 환경 변수 조회
Private Function GetApiKeyFromEnv(ByVal url As String, ByVal modelName As String) As String
    Dim lowerUrl As String
    lowerUrl = LCase(url)
    Dim lowerModel As String
    lowerModel = LCase(modelName)
    
    If isAnthropicUrl(url) Or InStr(lowerModel, "claude") > 0 Then
        GetApiKeyFromEnv = Environ("ANTHROPIC_API_KEY")
    ElseIf IsGeminiUrl(url) Or InStr(lowerModel, "gemini") > 0 Then
        GetApiKeyFromEnv = Environ("GEMINI_API_KEY")
    ElseIf IsOpenAiUrl(url) Or InStr(lowerModel, "gpt") > 0 Then
        GetApiKeyFromEnv = Environ("OPENAI_API_KEY")
    ElseIf IsUpstageUrl(url) Or Left(lowerModel, 6) = "solar-" Then
        GetApiKeyFromEnv = Environ("UPSTAGE_API_KEY")
    Else
        GetApiKeyFromEnv = ""  ' 로컬 LLM 등 다른 경우
    End If
End Function

' Modify ResolveModelAndBaseUrl function to correctly handle Upstage Solar models
Private Function ResolveModelAndBaseUrl(ByRef modelName As String, ByRef baseUrl As String, _
                                       Optional model As Variant, Optional baseUrlParam As Variant) As Boolean
    ' Set model name from parameter or default
    If Not IsMissing(model) And Not IsEmpty(model) Then
        modelName = CStr(model)
    Else
        modelName = DEFAULT_MODEL
    End If
    
    ' Set baseUrl from parameter or determine from model
    If IsMissing(baseUrlParam) Or IsEmpty(baseUrlParam) Then
        ' Auto-determine baseUrl from model name
        baseUrl = GetBaseUrlFromModel(modelName)
    Else
        baseUrl = CStr(baseUrlParam)
    End If
    
    ' Ensure URL ends with "/"
    If Right(baseUrl, 1) <> "/" Then
        baseUrl = baseUrl & "/"
    End If
    
    ' For Upstage Solar models, construct correct endpoint format
    If Left(LCase(modelName), 6) = "solar-" Then
        ' If base URL already contains "solar/", don't append it again
        If InStr(LCase(baseUrl), "solar/") = 0 Then
            baseUrl = baseUrl & "solar/"
        End If
    End If
    
    ResolveModelAndBaseUrl = True
End Function

' Anthropic 모델 또는 URL인지 확인하는 함수
Private Function IsAnthropicModelOrUrl(ByVal modelName As String, ByVal baseUrl As String) As Boolean
    IsAnthropicModelOrUrl = (InStr(LCase(modelName), "claude") > 0) Or (InStr(LCase(baseUrl), "anthropic") > 0)
End Function

' API 키 확인 및 오류 처리를 위한 공통 함수
Private Function ResolveApiKey(ByVal url As String, ByVal modelName As String, Optional apiKey As Variant) As String
    Dim finalApiKey As String
    Dim errorMessage As String
    
    ' API 키 설정
    If IsMissing(apiKey) Or IsEmpty(apiKey) Then
        finalApiKey = GetApiKeyFromEnv(url, modelName)
        
        ' 필요한 API 키가 없는 경우 에러 메시지 생성
        If finalApiKey = "" And Not IsLocalLLM(url) Then
            ' 로컬 LLM이 아닌 경우만 API 키 필요
            Dim apiKeyEnvName As String
            If IsOpenAiUrl(url) Then
                apiKeyEnvName = "OPENAI_API_KEY"
            ElseIf IsGeminiUrl(url) Then
                apiKeyEnvName = "GEMINI_API_KEY"
            ElseIf IsUpstageUrl(url) Or Left(LCase(modelName), 6) = "solar-" Then
                apiKeyEnvName = "UPSTAGE_API_KEY"
            Else
                apiKeyEnvName = "appropriate"
            End If
            
            errorMessage = "Error: API requires an api key. Provide it as the last argument or set the " & apiKeyEnvName & " environment variable."
            Err.Raise vbObjectError + 1000, "ResolveApiKey", errorMessage
        End If
    Else
        finalApiKey = CStr(apiKey)
    End If
    
    ResolveApiKey = finalApiKey
End Function

' Modify LLM_Dispatcher to use appropriate payload builder for Upstage models
Private Function LLM_Dispatcher(ByVal userPrompt As String, Optional systemPrompt As String = "", _
                               Optional temperature As Variant, Optional maxTokens As Variant, _
                               Optional model As Variant, Optional baseUrl As Variant, _
                               Optional apiKey As Variant) As String
    ' Variables initialization
    Dim modelName As String, effectiveBaseUrl As String
    
    ' Normalize model and URL
    Call ResolveModelAndBaseUrl(modelName, effectiveBaseUrl, model, baseUrl)
    
    ' Request processing
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    If IsAnthropicModelOrUrl(modelName, effectiveBaseUrl) Then
        ' Anthropic API
        response = LLM_Base_Anthropic(systemPrompt, userPrompt, temperature, maxTokens, modelName, apiKey)
    ElseIf Left(LCase(modelName), 6) = "solar-" Then
        ' Upstage Solar model - use OpenAI chat completions format
        Dim solarEndpoint As String
        solarEndpoint = effectiveBaseUrl ' Already prepared in ResolveModelAndBaseUrl
        response = LLM_Base_OpenAI(systemPrompt, userPrompt, temperature, maxTokens, modelName, solarEndpoint, apiKey)
    ElseIf systemPrompt <> "" Then
        ' OpenAI compatible API (with system prompt)
        response = LLM_Base_OpenAI(systemPrompt, userPrompt, temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    Else
        ' Single prompt method
        response = LLM_Base_Simple(userPrompt, temperature, maxTokens, modelName, effectiveBaseUrl, apiKey)
    End If
    
    LLM_Dispatcher = response
    Exit Function
    
ErrorHandler:
    LLM_Dispatcher = "Error: " & Err.Description
End Function

' 텍스트 이스케이프
Private Function EscapeText(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "\", "\\")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, """", "\""")
    EscapeText = result
End Function

' 텍스트 이스케이프 해제
Private Function UnescapeText(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "\n", vbLf)
    result = Replace(result, "\\", "\")
    result = Replace(result, "\""", """")
    UnescapeText = result
End Function

' Helper function to properly escape text for JSON
Private Function EscapeForJSON(ByVal text As String) As String
    ' First replace backslashes
    Dim result As String
    result = Replace(text, "\", "\\")
    
    ' Replace quotes
    result = Replace(result, """", "\""")
    
    ' Replace newlines
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbCr, "\n")
    
    ' Replace tabs
    result = Replace(result, vbTab, "\t")
    
    ' Replace backspace
    result = Replace(result, Chr(8), "\b")
    
    ' Replace form feed
    result = Replace(result, Chr(12), "\f")
    
    EscapeForJSON = result
End Function

'===============================================================
' JSON 페이로드 구성 함수들
'===============================================================

' JSON 페이로드 구성 - 단일 프롬프트 방식
Private Function BuildJsonPayload_Simple(ByVal modelName As String, ByVal prompt As String, _
                                  Optional ByVal temperature As Variant, Optional ByVal maxTokens As Variant) As String
    Dim jsonPayload As String
    
    ' 기존 OpenAI 호환 페이로드 구성
    jsonPayload = "{" & _
                  """model"": """ & modelName & """," & _
                  """messages"": [{" & _
                  """role"": ""user""," & _
                  """content"": """ & Replace(prompt, """", "\""") & """" & _
                  "}]"
    
    ' 온도 (temperature) 추가
    If Not IsMissing(temperature) And Not IsEmpty(temperature) Then
        If IsNumeric(temperature) Then
            jsonPayload = jsonPayload & ", ""temperature"": " & temperature
        End If
    End If
    
    ' max_tokens 추가
    If Not IsMissing(maxTokens) And Not IsEmpty(maxTokens) Then
        If IsNumeric(maxTokens) Then
            jsonPayload = jsonPayload & ", ""max_tokens"": " & maxTokens
        End If
    End If
    
    jsonPayload = jsonPayload & "}"
    Debug.Print "jsonPayload in BuildJsonPayload_Simple:" & jsonPayload
    BuildJsonPayload_Simple = jsonPayload
End Function

' JSON 페이로드 구성 - OpenAI role별 메시지 방식
Private Function BuildJsonPayload_OpenAI_responses(ByVal modelName As String, ByVal developerPrompt As String, ByVal userPrompt As String, _
                                  Optional ByVal temperature As Variant, Optional ByVal maxTokens As Variant) As String
    Dim jsonPayload As String
    
    jsonPayload = "{" & _
                  """model"": """ & modelName & """," & _
                  """input"": ["
    
    ' System 메시지 추가 (선택적)
    If developerPrompt <> "" Then
        jsonPayload = jsonPayload & "{" & _
                      """role"": ""developer""," & _
                      """content"": """ & Replace(developerPrompt, """", "\""") & """" & _
                      "}"
        
        ' User 메시지가 있으면 쉼표 추가
        If userPrompt <> "" Then
            jsonPayload = jsonPayload & ", "
        End If
    End If
    
    ' User 메시지 추가
    If userPrompt <> "" Then
        jsonPayload = jsonPayload & "{" & _
                      """role"": ""user""," & _
                      """content"": """ & Replace(userPrompt, """", "\""") & """" & _
                      "}"
    End If
    
    jsonPayload = jsonPayload & "]"
    
    ' 온도 (temperature) 추가
    If Not IsMissing(temperature) And Not IsEmpty(temperature) Then
        If IsNumeric(temperature) Then
            jsonPayload = jsonPayload & ", ""temperature"": " & temperature
        End If
    End If
    
    ' max_tokens 추가
    If Not IsMissing(maxTokens) And Not IsEmpty(maxTokens) Then
        If IsNumeric(maxTokens) Then
            jsonPayload = jsonPayload & ", ""max_tokens"": " & maxTokens
        End If
    End If
    
    jsonPayload = jsonPayload & "}"
    Debug.Print jsonPayload
    BuildJsonPayload_OpenAI_responses = jsonPayload
End Function

' JSON 페이로드 구성 - OpenAI chat completions 방식
Private Function BuildJsonPayload_OpenAI_chatcompletions(ByVal modelName As String, ByVal systemPrompt As String, ByVal userPrompt As String, _
                                  Optional ByVal temperature As Variant, Optional ByVal maxTokens As Variant) As String
    ' Properly escape text for JSON
    systemPrompt = EscapeForJSON(systemPrompt)
    userPrompt = EscapeForJSON(userPrompt)
    
    Dim jsonPayload As String
    
    jsonPayload = "{" & _
                  """model"": """ & modelName & """," & _
                  """messages"": ["
    
    ' Add System message if provided
    If Trim(systemPrompt) <> "" Then
        jsonPayload = jsonPayload & "{" & _
                      """role"": ""system""," & _
                      """content"": """ & systemPrompt & """" & _
                      "}"
        
        ' Add comma if we're also adding a user message
        If Trim(userPrompt) <> "" Then
            jsonPayload = jsonPayload & ", "
        End If
    End If
    
    ' Add User message if provided
    If Trim(userPrompt) <> "" Then
        jsonPayload = jsonPayload & "{" & _
                      """role"": ""user""," & _
                      """content"": """ & userPrompt & """" & _
                      "}"
    End If
    
    jsonPayload = jsonPayload & "]"
    
    ' Add temperature if provided
    If Not IsMissing(temperature) And Not IsEmpty(temperature) Then
        If IsNumeric(temperature) Then
            jsonPayload = jsonPayload & ", ""temperature"": " & temperature
        End If
    End If
    
    ' Add max_tokens if provided
    If Not IsMissing(maxTokens) And Not IsEmpty(maxTokens) Then
        If IsNumeric(maxTokens) Then
            jsonPayload = jsonPayload & ", ""max_tokens"": " & maxTokens
        End If
    End If
    
    jsonPayload = jsonPayload & "}"
    Debug.Print "jsonPayload in BuildJsonPayload_OpenAI_chatcompletions:" & jsonPayload
    BuildJsonPayload_OpenAI_chatcompletions = jsonPayload
End Function

' JSON 페이로드 구성 - Anthropic 메시지 방식
Private Function BuildJsonPayload_Anthropic(ByVal modelName As String, ByVal systemPrompt As String, ByVal userPrompt As String, _
                                     Optional ByVal temperature As Variant, Optional ByVal maxTokens As Variant) As String
    ' Properly escape text for JSON
    systemPrompt = EscapeForJSON(systemPrompt)
    userPrompt = EscapeForJSON(userPrompt)
    
    Dim jsonPayload As String
    
    ' Anthropic API payload structure
    jsonPayload = "{" & _
                  """model"": """ & modelName & """"
    
    ' Add System message if provided
    If systemPrompt <> "" Then
        jsonPayload = jsonPayload & ", ""system"": """ & systemPrompt & """"
    End If
    
    ' Add User message
    jsonPayload = jsonPayload & ", ""messages"": [{" & _
                  """role"": ""user""," & _
                  """content"": """ & userPrompt & """" & _
                  "}]"
    
    ' Anthropic requires max_tokens
    If IsMissing(maxTokens) Or IsEmpty(maxTokens) Then
        jsonPayload = jsonPayload & ", ""max_tokens"": 4096"
    ElseIf IsNumeric(maxTokens) Then
        jsonPayload = jsonPayload & ", ""max_tokens"": " & maxTokens
    Else
        jsonPayload = jsonPayload & ", ""max_tokens"": 4096"
    End If
    
    ' Add temperature if provided
    If Not IsMissing(temperature) And Not IsEmpty(temperature) Then
        If IsNumeric(temperature) Then
            jsonPayload = jsonPayload & ", ""temperature"": " & temperature
        End If
    End If
    
    jsonPayload = jsonPayload & "}"
    BuildJsonPayload_Anthropic = jsonPayload
End Function

'===============================================================
' HTTP 통신 및 응답 처리 함수들
'===============================================================

' HTTP 요청 전송
Private Function SendLLMRequest(ByVal url As String, ByVal jsonPayload As String, _
                               Optional apiKey As Variant, Optional isAnthropicAPI As Boolean = False) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo ErrorHandler
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    
    ' API 키가 제공되면 인증 헤더 추가
    SetAuthorizationHeader http, apiKey, isAnthropicAPI
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

' 응답에서 콘텐츠 추출
Private Function ExtractContent(ByVal response As String, Optional isAnthropicAPI As Boolean = False) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    If isAnthropicAPI Then
        ' Anthropic 응답 포맷에 맞는 패턴 설정
        regEx.Pattern = """content"":\s*\[\s*{\s*""type"":\s*""text"",\s*""text"":\s*""([\s\S]*?)""}\s*\]"
    Else
        ' 기존 OpenAI/기타 API 패턴
        regEx.Pattern = """content"":\s*""([\s\S]*?)""\s*(?:,|\})"
    End If
    
    regEx.IgnoreCase = True
    regEx.Global = False
    
    Dim matches As Object
    Set matches = regEx.Execute(response)
    
    If matches.Count > 0 Then
        ExtractContent = Replace(matches(0).SubMatches(0), "\""", """")
    Else
        ' 패턴 매칭 실패 시 대체 방법
        ' Anthropic API 형식 변경 가능성에 대비
        If isAnthropicAPI Then
            ' 다른 Anthropic 응답 형식 시도
            regEx.Pattern = """text"":\s*""([\s\S]*?)"""
            Set matches = regEx.Execute(response)
            If matches.Count > 0 Then
                ExtractContent = Replace(matches(0).SubMatches(0), "\""", """")
                Exit Function
            End If
        End If
        ExtractContent = "Error: Failed to parse response"
    End If
End Function

' 구분자 사이의 내용을 추출하고 불필요한 공백과 빈 줄을 제거하는 함수
Private Function ExtractContentBetweenDelimiters(ByVal text As String, ByVal startDelimiter As String, ByVal endDelimiter As String, Optional ByVal defaultValue As String = "") As String
    Dim startPos As Long, endPos As Long
    
    ' 시작 구분자 찾기
    startPos = InStr(1, text, startDelimiter)
    If startPos = 0 Then
        ExtractContentBetweenDelimiters = defaultValue
        Exit Function
    End If
    
    ' 시작 구분자 다음 위치 계산
    startPos = startPos + Len(startDelimiter)
    
    ' 끝 구분자 찾기
    endPos = InStr(startPos, text, endDelimiter)
    If endPos = 0 Then
        ExtractContentBetweenDelimiters = defaultValue
        Exit Function
    End If
    
    ' 구분자 사이의 내용 추출
    Dim extractedContent As String
    extractedContent = Mid(text, startPos, endPos - startPos)
    
    ' 앞뒤 공백과 빈 줄 제거
    extractedContent = TrimExtraWhitespace(extractedContent)
    
    ExtractContentBetweenDelimiters = extractedContent
End Function

' 문자열의 앞뒤 공백, 빈 줄을 제거하는 향상된 함수
Private Function TrimExtraWhitespace(ByVal text As String) As String
    Dim result As String
    result = text
    
    ' 앞쪽 공백과 빈 줄 제거
    Do While Len(result) > 0
        If Left(result, 2) = vbCrLf Then
            result = Mid(result, 3)
        ElseIf Left(result, 1) = vbCr Or Left(result, 1) = vbLf Then
            result = Mid(result, 2)
        ElseIf Left(result, 1) = " " Or Left(result, 1) = vbTab Then
            result = Mid(result, 2)
        Else
            Exit Do
        End If
    Loop
    
    ' 뒤쪽 공백과 빈 줄 제거
    Do While Len(result) > 0
        If Right(result, 2) = vbCrLf Then
            result = Left(result, Len(result) - 2)
        ElseIf Right(result, 1) = vbCr Or Right(result, 1) = vbLf Then
            result = Left(result, Len(result) - 1)
        ElseIf Right(result, 1) = " " Or Right(result, 1) = vbTab Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop
    
    TrimExtraWhitespace = result
End Function

' Remove markdown formatting from text
Private Function CleanTextFromMarkdown(ByVal text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Remove code blocks (```...```)
    regEx.Pattern = "```[^`]*```"
    regEx.Global = True
    CleanTextFromMarkdown = regEx.Replace(text, "")
    
    ' Remove inline code (`...`)
    regEx.Pattern = "`[^`]*`"
    CleanTextFromMarkdown = regEx.Replace(CleanTextFromMarkdown, "")
    
    Set regEx = Nothing
End Function

' 문자열 앞부분에 있는 모든 줄 바꿈(CR, LF) 제거
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

' 응답을 처리하여 showThink 옵션에 따라 결과를 분리 반환
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

'===============================================================
' 베이스 LLM 함수들
'===============================================================

' 단일 프롬프트 방식 (기존 방식 유지하면서 중복 코드 제거)
Function LLM_Base_Simple(prompt As String, Optional temperature As Variant, Optional maxTokens As Variant, _
                        Optional model As Variant, Optional baseUrl As Variant, Optional apiKey As Variant) As String
    ' 모델명 설정
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_MODEL
    Else
        modelName = CStr(model)
    End If
    
    ' URL 설정
    Dim url As String
    If Not IsMissing(baseUrl) And Not IsEmpty(baseUrl) Then
        Dim baseStr As String
        baseStr = CStr(baseUrl)
        If Right(baseStr, 1) <> "/" Then baseStr = baseStr & "/"
        url = baseStr & "chat/completions"
    Else
        ' 모델명에 따라 baseUrl 자동 결정
        Dim autoBaseUrl As String
        autoBaseUrl = GetBaseUrlFromModel(modelName)
        
        If autoBaseUrl <> "" Then
            If Right(autoBaseUrl, 1) <> "/" Then autoBaseUrl = autoBaseUrl & "/"
            url = autoBaseUrl & "chat/completions"
        Else
            url = BASE_URL_DEFAULT
            If Right(url, 1) <> "/" Then url = url & "/"
            url = url & "chat/completions"
        End If
    End If
    
    On Error Resume Next
    ' API 키 확인 및 설정
    Dim finalApiKey As String
    finalApiKey = ResolveApiKey(url, modelName, apiKey)
    
    If Err.Number <> 0 Then
        LLM_Base_Simple = Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    ' JSON 페이로드 구성
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload_Simple(modelName, EscapeText(prompt), temperature, maxTokens)
    
    ' API 요청 및 응답 수신
    Dim response As String
    response = SendLLMRequest(url, jsonPayload, finalApiKey)
    
    If Left(response, 6) = "Error:" Then
        LLM_Base_Simple = response
        Exit Function
    End If
    
    ' 응답에서 콘텐츠 추출 및 반환
    LLM_Base_Simple = UnescapeText(ExtractContent(response))
End Function

' OpenAI 호환 채팅 완성 방식
Function LLM_Base_OpenAI(Optional systemPrompt As String = "", Optional userPrompt As String = "", _
                        Optional temperature As Variant, Optional maxTokens As Variant, _
                        Optional model As Variant, Optional baseUrl As Variant, Optional apiKey As Variant) As String
    
    ' Ensure we have at least one non-empty prompt
    If Trim(systemPrompt) = "" And Trim(userPrompt) = "" Then
        LLM_Base_OpenAI = "Error: At least one of systemPrompt or userPrompt must be provided"
        Exit Function
    End If
    
    ' Model name setup
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_MODEL
    Else
        modelName = CStr(model)
    End If
    
    ' URL setup with automatic base URL determination from model name if needed
    Dim url As String
    If Not IsMissing(baseUrl) And Not IsEmpty(baseUrl) Then
        Dim baseStr As String
        baseStr = CStr(baseUrl)
        ' Correct the URL formation
        If Right(baseStr, 1) <> "/" Then
            url = baseStr & "/"
        Else
            url = baseStr
        End If
    Else
        ' 모델명에 따라 baseUrl 자동 결정
        Dim autoBaseUrl As String
        autoBaseUrl = GetBaseUrlFromModel(modelName)
        
        If autoBaseUrl <> "" Then
            If Right(autoBaseUrl, 1) <> "/" Then
                url = autoBaseUrl & "/"
            Else
                url = autoBaseUrl
            End If
        Else
            url = BASE_URL_DEFAULT
            If Right(url, 1) <> "/" Then
                url = url & "/"
            End If
        End If
    End If
    
    ' 채팅 완성 엔드포인트 추가
    url = url & "chat/completions"
    
    On Error Resume Next
    ' API 키 확인 및 설정
    Dim finalApiKey As String
    finalApiKey = ResolveApiKey(url, modelName, apiKey)
    
    If Err.Number <> 0 Then
        LLM_Base_OpenAI = Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Create the JSON payload based on the API type
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload_OpenAI_chatcompletions(modelName, systemPrompt, userPrompt, temperature, maxTokens)
    
    ' Send the API request and get the response
    Dim response As String
    response = SendLLMRequest(url, jsonPayload, finalApiKey)
    
    If Left(response, 6) = "Error:" Then
        LLM_Base_OpenAI = response
        Exit Function
    End If
    
    ' Extract content from the response and return
    LLM_Base_OpenAI = UnescapeText(ExtractContent(response))
End Function

' Anthropic API 방식
Function LLM_Base_Anthropic(Optional systemPrompt As String = "", Optional userPrompt As String = "", _
                           Optional temperature As Variant, Optional maxTokens As Variant, _
                           Optional model As Variant, Optional apiKey As Variant) As String
    
    ' URL은 고정 (Anthropic API)
    Dim url As String
    url = ANTHROPIC_URL
    
    ' 모델명 설정
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_ANTHROPIC_MODEL
    Else
        modelName = CStr(model)
    End If
    
    On Error Resume Next
    ' API 키 확인 및 설정
    Dim finalApiKey As String
    finalApiKey = ResolveApiKey(url, modelName, apiKey)
    
    If Err.Number <> 0 Then
        LLM_Base_Anthropic = Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    ' JSON 페이로드 구성
    Dim jsonPayload As String
    jsonPayload = BuildJsonPayload_Anthropic(modelName, systemPrompt, userPrompt, temperature, maxTokens)
    
    ' API 요청 및 응답 수신
    Dim response As String
    response = SendLLMRequest(url, jsonPayload, finalApiKey, True)  ' Anthropic API임을 명시
    
    If Left(response, 6) = "Error:" Then
        LLM_Base_Anthropic = response
        Exit Function
    End If
    
    ' 응답에서 콘텐츠 추출 및 반환
    LLM_Base_Anthropic = UnescapeText(ExtractContent(response, True))  ' Anthropic 응답 파싱
End Function

'===============================================================
' 사용자 인터페이스 함수 섹션 - 중복 코드 제거
'===============================================================

Function LLM(prompt As String, Optional value As String = "", Optional temperature As Variant, _
             Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
             Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 프롬프트 구성
    Dim fullPrompt As String
    fullPrompt = prompt
    If value <> "" Then fullPrompt = fullPrompt & " " & value
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    response = LLM_Dispatcher(fullPrompt, "", temperature, maxTokens, model, baseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM = ProcessLLMResponse(response, showThink)
End Function

Function LLM_SUMMARIZE(text As String, Optional prompt As String = "", Optional temperature As Variant, _
                     Optional maxTokens As Variant, Optional model As Variant, Optional baseUrl As Variant, _
                     Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 프롬프트 구성
    If prompt = "" Then
        prompt = "Summarize in one line:"
    End If
    Dim fullPrompt As String
    fullPrompt = prompt & " " & text
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    response = LLM_Dispatcher(fullPrompt, "", temperature, maxTokens, model, baseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_SUMMARIZE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_CODE(programDetails As String, programmingLanguage As String, _
                  Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 프롬프트 구성
    Dim prompt As String
    prompt = "Generate a " & programmingLanguage & " program that fulfills the following requirements:" & vbCrLf & programDetails
    
    ' LLM 디스패처를 통해 요청 처리 (코드 생성에는 낮은 temperature 사용)
    Dim response As String
    response = LLM_Dispatcher(prompt, "", 0.2, , model, baseUrl, apiKey)
    
    ' 응답 처리 및 반환
    LLM_CODE = ProcessLLMResponse(response, showThink)
End Function

Function LLM_LIST(prompt As String, Optional model As Variant, Optional baseUrl As Variant, _
                  Optional showThink As Boolean = False, Optional apiKey As Variant) As Variant
    ' 프롬프트 구성
    Dim listPrompt As String
    listPrompt = prompt & vbCrLf & _
                 "Example:" & vbCrLf & _
                 "<list><item>Apple</item><item>Banana</item><item>Cherry</item></list>" & vbCrLf & _
                 "Please output only the list items enclosed within <list> and <item> tags, exactly in the above format, with no additional commentary."
    
    ' LLM 디스패처를 통해 요청 처리
    Dim response As String
    response = LLM_Dispatcher(listPrompt, "", , , model, baseUrl, apiKey)
    
    ' 응답 처리
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
    
    ' showThink 옵션에 따라 결과 배열 반환
    If showThink Then
        Dim resultArray(1) As Variant
        resultArray(0) = thinkText
        resultArray(1) = items
        LLM_LIST = resultArray
    Else
        LLM_LIST = items
    End If
End Function
