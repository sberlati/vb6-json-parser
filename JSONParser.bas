Option Explicit

Private Type JSONSTATE
    json As String
    position As Long
End Type

Private state As JSONSTATE

' Main parsing function - entry point
Public Function ParseJSON(ByVal jsonString As String) As Object
    state.json = jsonString
    state.position = 1
    
    SkipWhitespace
    
    Select Case Mid(state.json, state.position, 1)
        Case "{"
            Set ParseJSON = ParseObject
        Case "["
            Set ParseJSON = ParseArray
        Case Else
            Err.Raise vbObjectError + 1, "ParseJSON", "Invalid JSON string"
    End Select
End Function

' Parse a JSON object
Private Function ParseObject() As Dictionary
    Dim dict As New Dictionary
    Dim key As String
    
    ' Skip opening brace
    state.position = state.position + 1
    
    Do
        SkipWhitespace
        
        ' Check for empty object or end of object
        If Mid(state.json, state.position, 1) = "}" Then
            state.position = state.position + 1
            Set ParseObject = dict
            Exit Function
        End If
        
        ' Parse key
        If Mid(state.json, state.position, 1) <> """" Then
            Err.Raise vbObjectError + 2, "ParseObject", "Expected property name"
        End If
        key = ParseString
        
        SkipWhitespace
        
        ' Expect colon
        If Mid(state.json, state.position, 1) <> ":" Then
            Err.Raise vbObjectError + 3, "ParseObject", "Expected ':'"
        End If
        state.position = state.position + 1
        
        ' Parse value
        dict.Add key, ParseValue
        
        SkipWhitespace
        
        ' Check for more properties
        Select Case Mid(state.json, state.position, 1)
            Case "}"
                state.position = state.position + 1
                Set ParseObject = dict
                Exit Function
            Case ","
                state.position = state.position + 1
            Case Else
                Err.Raise vbObjectError + 4, "ParseObject", "Expected ',' or '}'"
        End Select
    Loop
End Function

' Parse a JSON array
Private Function ParseArray() As Collection
    Dim arr As New Collection
    
    ' Skip opening bracket
    state.position = state.position + 1
    
    Do
        SkipWhitespace
        
        ' Check for empty array or end of array
        If Mid(state.json, state.position, 1) = "]" Then
            state.position = state.position + 1
            Set ParseArray = arr
            Exit Function
        End If
        
        ' Parse value
        arr.Add ParseValue
        
        SkipWhitespace
        
        ' Check for more elements
        Select Case Mid(state.json, state.position, 1)
            Case "]"
                state.position = state.position + 1
                Set ParseArray = arr
                Exit Function
            Case ","
                state.position = state.position + 1
            Case Else
                Err.Raise vbObjectError + 5, "ParseArray", "Expected ',' or ']'"
        End Select
    Loop
End Function

' Parse a JSON value
Private Function ParseValue() As Variant
    SkipWhitespace
    
    Select Case Mid(state.json, state.position, 1)
        Case "{"
            Set ParseValue = ParseObject
        Case "["
            Set ParseValue = ParseArray
        Case """"
            ParseValue = ParseString
        Case "t"
            ParseValue = ParseTrue
        Case "f"
            ParseValue = ParseFalse
        Case "n"
            ParseValue = ParseNull
        Case "-", "0" To "9"
            ParseValue = ParseNumber
        Case Else
            Err.Raise vbObjectError + 6, "ParseValue", "Invalid value"
    End Select
End Function

' Parse a JSON string
Private Function ParseString() As String
    Dim result As String
    Dim char As String
    
    ' Skip opening quote
    state.position = state.position + 1
    
    Do While state.position <= Len(state.json)
        char = Mid(state.json, state.position, 1)
        
        Select Case char
            Case """"
                state.position = state.position + 1
                ParseString = result
                Exit Function
            Case "\"
                state.position = state.position + 1
                char = Mid(state.json, state.position, 1)
                
                Select Case char
                    Case """", "\", "/"
                        result = result & char
                    Case "b"
                        result = result & vbBack
                    Case "f"
                        result = result & vbFormFeed
                    Case "n"
                        result = result & vbNewLine
                    Case "r"
                        result = result & vbCr
                    Case "t"
                        result = result & vbTab
                    Case "u"
                        ' Unicode escape sequence
                        Dim hexCode As String
                        hexCode = Mid(state.json, state.position + 1, 4)
                        result = result & ChrW$(CLng("&H" & hexCode))
                        state.position = state.position + 4
                    Case Else
                        Err.Raise vbObjectError + 7, "ParseString", "Invalid escape sequence"
                End Select
            Case Else
                result = result & char
        End Select
        
        state.position = state.position + 1
    Loop
    
    Err.Raise vbObjectError + 8, "ParseString", "Unterminated string"
End Function

' Parse a JSON number
Private Function ParseNumber() As Variant
    Dim numStr As String
    Dim char As String
    
    Do While state.position <= Len(state.json)
        char = Mid(state.json, state.position, 1)
        
        If InStr("0123456789+-.eE", char) > 0 Then
            numStr = numStr & char
            state.position = state.position + 1
        Else
            Exit Do
        End If
    Loop
    
    If InStr(1, numStr, ".", vbTextCompare) > 0 Or _
       InStr(1, numStr, "e", vbTextCompare) > 0 Or _
       InStr(1, numStr, "E", vbTextCompare) > 0 Then
        ParseNumber = CDbl(numStr)
    Else
        ParseNumber = CLng(numStr)
    End If
End Function

' Parse JSON true value
Private Function ParseTrue() As Boolean
    If Mid(state.json, state.position, 4) = "true" Then
        state.position = state.position + 4
        ParseTrue = True
    Else
        Err.Raise vbObjectError + 9, "ParseTrue", "Expected 'true'"
    End If
End Function

' Parse JSON false value
Private Function ParseFalse() As Boolean
    If Mid(state.json, state.position, 5) = "false" Then
        state.position = state.position + 5
        ParseFalse = False
    Else
        Err.Raise vbObjectError + 10, "ParseFalse", "Expected 'false'"
    End If
End Function

' Parse JSON null value
Private Function ParseNull() As Variant
    If Mid(state.json, state.position, 4) = "null" Then
        state.position = state.position + 4
        ParseNull = Null
    Else
        Err.Raise vbObjectError + 11, "ParseNull", "Expected 'null'"
    End If
End Function

' Skip whitespace characters
Private Sub SkipWhitespace()
    Dim char As String
    
    Do While state.position <= Len(state.json)
        char = Mid(state.json, state.position, 1)
        
        If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
            state.position = state.position + 1
        Else
            Exit Do
        End If
    Loop
End Sub

