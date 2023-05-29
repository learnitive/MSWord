' This is a macro that will communicate with the Learnitive API in Microsoft Word.
' Just copy/paste this macro into Word following instructions in the Readme.md file.
' Don't forget to change the API key to your own.
' Original Author: Johann Dowa
' Updated by: Learnitive 2023
' MIT license

Function UnescapeString(ByVal str As String) As String
    Dim i As Integer
    Dim output As String
    For i = 1 To Len(str)
        If Mid(str, i, 2) = "\\" Then
            output = output & "\"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\/" Then
            output = output & "/"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\n" Then
            output = output & vbCrLf
            i = i + 1
        ElseIf Mid(str, i, 2) = "\r" Then
            output = output & vbCr
            i = i + 1
        ElseIf Mid(str, i, 2) = "\t" Then
            output = output & vbTab
            i = i + 1
        ElseIf Mid(str, i, 2) = "\" & Chr(34) Then
            output = output & """"
            i = i + 1
        Else
            output = output & Mid(str, i, 1)
        End If
    Next i
    UnescapeString = output
End Function

Sub Learnitive()
    '
    ' Learnitive AI Macro
    '
    
    Dim apiUrl As String
    Dim requestPayload As String
    Dim apiKey As String
    Dim httpRequest As Object
    Dim responseText As String
    Dim content As String
    Dim startIndex As Integer
    Dim endIndex As Integer
    
    content = InputBox("Enter your input", "Content")
    

    ' Change the API key to your own
    apiKey = "YOUR-KEY-HERE"
    
    ' Modify the API parameter as required
    model = "expert"
    max_tokens = "100"
    temperature = ".7"
    
    ' API endpoint url, learn more at https://www.learnitive.com/ai-api
    apiUrl = "https://www.learnitive.com/api/v1/contents"
      
    requestPayload = "{""input"":""" & content & """, ""model"":""" & model & """, ""max_tokens"":""" & max_tokens & """, ""temperature"":""" & temperature & """ }"
    
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "POST", apiUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    
    httpRequest.setRequestHeader "api-key", "" & apiKey
    On Error Resume Next
    httpRequest.send requestPayload
    On Error GoTo 0
    
    If httpRequest.Status <> 200 Then
        MsgBox "Error: " & httpRequest.Status & " " & httpRequest.StatusText
        Exit Sub
    End If
        
    responseText = httpRequest.responseText
    startPos = InStr(responseText, """text"":""") + 8
    
    endPos = InStr(responseText, """id"":""") - 2

    responseText = Trim(UnescapeString(Mid(responseText, startPos, endPos - startPos + 1)))
    
    Dim wdApp As Word.Application
    Set wdApp = GetObject(, "Word.Application")
    wdApp.Selection.Text = responseText
    Set httpRequest = Nothing
    

End Sub
