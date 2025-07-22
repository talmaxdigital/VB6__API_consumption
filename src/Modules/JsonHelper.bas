Attribute VB_Name = "JsonHelper"
Option Explicit

Private Type JSONSTATE
    json As String
    position As Long
End Type

Private state As JSONSTATE

Public Function ParseJSON(ByVal jsonString As String) As Object
    ' Função principal para analisar uma string JSON e retornar um objeto
    '
    ' Args:
    '   jsonString (String): String JSON válida para ser analisada
    '
    ' Result:
    '   Object: Dictionary para objetos JSON ou Collection para arrays JSON
    '
    ' Example:
    '   Dim jsonObj As Object
    '   Set jsonObj = ParseJSON("{""nome"":""João"",""idade"":30}")
    '   Debug.Print jsonObj("nome") ' Output: João

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

Public Function BuildJSON(ByVal obj As Variant) As String
    ' Função principal para construir uma string JSON a partir de um objeto
    '
    ' Args:
    '   obj (Variant): Objeto VB6 (Dictionary, Collection, ou valor primitivo)
    '
    ' Result:
    '   String: String JSON válida representando o objeto
    '
    ' Example:
    '   Dim dict As New Dictionary
    '   dict.Add "nome", "João"
    '   dict.Add "idade", 30
    '   Debug.Print BuildJSON(dict) ' Output: {"nome":"João","idade":30}

    BuildJSON = BuildValue(obj)
End Function

Private Function BuildValue(ByVal value As Variant) As String
    ' Constrói qualquer valor JSON e retorna a string apropriada
    '
    ' Args:
    '   value (Variant): Valor a ser convertido (Dictionary, Collection, String, Number, Boolean, Null)
    '
    ' Result:
    '   String: Representação JSON do valor

    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            BuildValue = BuildObject(value)
        ElseIf TypeName(value) = "Collection" Then
            BuildValue = BuildArray(value)
        Else
            Err.Raise vbObjectError + 20, "BuildValue", "Unsupported object type: " & TypeName(value)
        End If
    ElseIf IsNull(value) Then
        BuildValue = "null"
    ElseIf VarType(value) = vbBoolean Then
        BuildValue = BuildBoolean(value)
    ElseIf IsNumeric(value) Then
        BuildValue = BuildNumber(value)
    Else
        BuildValue = BuildString(CStr(value))
    End If
End Function

Private Function BuildObject(ByVal dict As Dictionary) As String
    ' Constrói um objeto JSON a partir de um Dictionary
    '
    ' Args:
    '   dict (Dictionary): Dictionary contendo as propriedades do objeto
    '
    ' Result:
    '   String: String JSON representando o objeto

    Dim result As String
    Dim key As Variant
    Dim isFirst As Boolean

    result = "{"
    isFirst = True

    For Each key In dict.Keys
        If Not isFirst Then
            result = result & ","
        End If

        result = result & BuildString(CStr(key)) & ":" & BuildValue(dict(key))
        isFirst = False
    Next key

    result = result & "}"
    BuildObject = result
End Function

Private Function BuildArray(ByVal coll As Collection) As String
    ' Constrói um array JSON a partir de uma Collection
    '
    ' Args:
    '   coll (Collection): Collection contendo os elementos do array
    '
    ' Result:
    '   String: String JSON representando o array

    Dim result As String
    Dim item As Variant
    Dim i As Integer

    result = "["

    For i = 1 To coll.Count
        If i > 1 Then
            result = result & ","
        End If

        result = result & BuildValue(coll(i))
    Next i

    result = result & "]"
    BuildArray = result
End Function

Private Function BuildString(ByVal str As String) As String
    ' Constrói uma string JSON com escape de caracteres
    '
    ' Args:
    '   str (String): String a ser codificada
    '
    ' Result:
    '   String: String JSON com caracteres de escape processados

    Dim result As String
    Dim i As Integer
    Dim char As String

    result = """"

    For i = 1 To Len(str)
        char = Mid(str, i, 1)

        Select Case char
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case "/"
                result = result & "\/"
            Case vbBack
                result = result & "\b"
            Case vbFormFeed
                result = result & "\f"
            Case vbNewLine
                result = result & "\n"
            Case vbCr
                result = result & "\r"
            Case vbTab
                result = result & "\t"
            Case Else
                If Asc(char) < 32 Then
                    result = result & "\u" & Right("0000" & Hex(Asc(char)), 4)
                Else
                    result = result & char
                End If
        End Select
    Next i

    result = result & """"
    BuildString = result
End Function

Private Function BuildNumber(ByVal num As Variant) As String
    ' Constrói um número JSON
    '
    ' Args:
    '   num (Variant): Número a ser convertido
    '
    ' Result:
    '   String: Representação JSON do número

    If VarType(num) = vbSingle Or VarType(num) = vbDouble Then
        ' Para números decimais, usar formato com ponto decimal
        BuildNumber = Replace(CStr(num), ",", ".")
    Else
        BuildNumber = CStr(num)
    End If
End Function

Private Function BuildBoolean(ByVal bool As Boolean) As String
    ' Constrói um valor booliano JSON
    '
    ' Args:
    '   bool (Boolean): Valor booliano a ser convertido
    '
    ' Result:
    '   String: "true" ou "false"

    If bool Then
        BuildBoolean = "true"
    Else
        BuildBoolean = "false"
    End If
End Function

Public Function CreateJSONObject() As Dictionary
    ' Cria um novo objeto JSON (Dictionary) vazio
    '
    ' Result:
    '   Dictionary: Novo Dictionary vazio para construir objetos JSON
    '
    ' Example:
    '   Dim obj As Dictionary
    '   Set obj = CreateJSONObject()
    '   obj.Add "nome", "Maria"
    '   obj.Add "idade", 25

    Set CreateJSONObject = New Dictionary
End Function

Public Function CreateJSONArray() As Collection
    ' Cria um novo array JSON (Collection) vazio
    '
    ' Result:
    '   Collection: Nova Collection vazia para construir arrays JSON
    '
    ' Example:
    '   Dim arr As Collection
    '   Set arr = CreateJSONArray()
    '   arr.Add "item1"
    '   arr.Add "item2"

    Set CreateJSONArray = New Collection
End Function

Private Function ParseObject() As Dictionary
    ' Analisa um objeto JSON e retorna um Dictionary
    '
    ' Result:
    '   Dictionary: Objeto Dictionary contendo as propriedades do objeto JSON
    Dim dict As New Dictionary
    Dim key As String

    state.position = state.position + 1

    Do
        SkipWhitespace

        If Mid(state.json, state.position, 1) = "}" Then
            state.position = state.position + 1
            Set ParseObject = dict
            Exit Function
        End If

        If Mid(state.json, state.position, 1) <> """" Then
            Err.Raise vbObjectError + 2, "ParseObject", "Expected property name"
        End If
        key = ParseString

        SkipWhitespace

        If Mid(state.json, state.position, 1) <> ":" Then
            Err.Raise vbObjectError + 3, "ParseObject", "Expected ':'"
        End If
        state.position = state.position + 1

        dict.Add key, ParseValue

        SkipWhitespace

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

Private Function ParseArray() As Collection
    ' Analisa um array JSON e retorna uma Collection
    '
    ' Result:
    '   Collection: Collection contendo os elementos do array JSON
    Dim arr As New Collection

    state.position = state.position + 1

    Do
        SkipWhitespace

        If Mid(state.json, state.position, 1) = "]" Then
            state.position = state.position + 1
            Set ParseArray = arr
            Exit Function
        End If

        arr.Add ParseValue

        SkipWhitespace

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

Private Function ParseValue() As Variant
    ' Analisa qualquer valor JSON e retorna o tipo apropriado
    '
    ' Result:
    '   Variant: Valor analisado que pode ser Dictionary, Collection, String, Number, Boolean ou Null
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

Private Function ParseString() As String
    ' Analisa uma string JSON com escape de caracteres
    '
    ' Result:
    '   String: String decodificada com caracteres de escape processados
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

Private Function ParseNumber() As Variant
    ' Analisa um número JSON e retorna Long ou Double conforme apropriado
    '
    ' Result:
    '   Variant: Long para inteiros ou Double para números decimais/científicos
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

Private Function ParseTrue() As Boolean
    ' Analisa o valor literal "true" do JSON
    '
    ' Result:
    '   Boolean: Sempre retorna True se válido
    If Mid(state.json, state.position, 4) = "true" Then
        state.position = state.position + 4
        ParseTrue = True
    Else
        Err.Raise vbObjectError + 9, "ParseTrue", "Expected 'true'"
    End If
End Function

Private Function ParseFalse() As Boolean
    ' Analisa o valor literal "false" do JSON
    '
    ' Result:
    '   Boolean: Sempre retorna False se válido
    If Mid(state.json, state.position, 5) = "false" Then
        state.position = state.position + 5
        ParseFalse = False
    Else
        Err.Raise vbObjectError + 10, "ParseFalse", "Expected 'false'"
    End If
End Function

Private Function ParseNull() As Variant
    ' Analisa o valor literal "null" do JSON
    '
    ' Result:
    '   Variant: Retorna Null se válido
    If Mid(state.json, state.position, 4) = "null" Then
        state.position = state.position + 4
        ParseNull = Null
    Else
        Err.Raise vbObjectError + 11, "ParseNull", "Expected 'null'"
    End If
End Function

Private Sub SkipWhitespace()
    ' Pula caracteres de espaço em branco na string JSON
    '
    ' Avança a posição até encontrar um caractere não-branco
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