Option Explicit

Const SERVER_URL As String = "http://localhost:1234/v1/chat/completions"
Const DEFAULT_MODEL As String = "exaone-3.5-7.8b-instruct"

Function LLM(prompt As String, Optional value As String = "", Optional temperature As Variant, Optional max_tokens As Variant, Optional model As Variant, Optional base_url As Variant) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim url As String
    ' base_url 인자가 제공되면 해당 주소에 "/v1/chat/completions"를 덧붙임
    If Not IsMissing(base_url) And Not IsEmpty(base_url) Then
        url = CStr(base_url) & "/v1/chat/completions"
    Else
        url = SERVER_URL
    End If

    ' 모델 설정: 지정되지 않으면 기본값 사용
    Dim modelName As String
    If IsMissing(model) Or IsEmpty(model) Then
        modelName = DEFAULT_MODEL
    Else
        modelName = CStr(model)
    End If

    ' prompt와 value를 결합
    Dim fullPrompt As String
    If value = "" Then
        fullPrompt = prompt
    Else
        fullPrompt = prompt & " " & value
    End If

    ' JSON 페이로드 생성
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

    ' HTTP POST 요청 전송
    On Error Resume Next
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonPayload

    ' 응답 처리
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText

        Dim startPos As Integer
        startPos = InStr(response, """content"": """) + Len("""content"": """)
        Dim endPos As Integer
        Dim quoteCount As Integer
        Dim i As Integer

        ' 이스케이프된 따옴표를 고려해 정확한 종료 위치 찾기
        quoteCount = 0
        i = startPos
        Do While i <= Len(response)
            If Mid(response, i, 1) = """" Then
                If i = 1 Or Mid(response, i - 1, 1) <> "\" Then
                    quoteCount = quoteCount + 1
                    If quoteCount = 2 Then
                        endPos = i
                        Exit Do
                    End If
                End If
            End If
            i = i + 1
        Loop

        If endPos > startPos Then
            Dim rawContent As String
            ' 필요에 따라 불필요한 뒷부분 길이(여기서는 -23)를 조절할 수 있음
            rawContent = Mid(response, startPos, endPos - startPos - 23)
            ' 이스케이프된 따옴표를 일반 따옴표로 변환
            LLM = Replace(rawContent, "\""", """")
        Else
            LLM = "Error: 응답 파싱 실패"
        End If
    Else
        LLM = "Error: " & http.Status & " " & http.statusText
    End If

    On Error GoTo 0
    Set http = Nothing
End Function
