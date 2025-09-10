Attribute VB_Name = "JsonHelper"
Option Explicit

Option Explicit

' ====================================================================
' JsonHelper Module - Parser e Builder JSON nativo para VB6
' Implementa��o completa de an�lise e gera��o de JSON sem depend�ncias
' ====================================================================

Private Type JSONSTATE
    Json As String
    position As Long
End Type

Private state As JSONSTATE

' ====================================================================
' FUN��ES PRINCIPAIS - PARSE E BUILD
' ====================================================================

Public Function ParseJSON(ByVal jsonString As String) As Object
    ' Analisa uma string JSON e retorna um objeto VB6 equivalente
    '
    ' Args:
    '   jsonString (String): String JSON v�lida para ser analisada
    '
    ' Result:
    '   Object: Dictionary para objetos JSON ou Collection para arrays JSON
    '
    ' Raises:
    '   vbObjectError + 1: String JSON inv�lida ou malformada
    '
    ' Example:
    '   Dim user As Object
    '   Set user = ParseJSON("{""nome"":""Jo�o"",""idade"":30,""ativo"":true}")
    '   Debug.Print user("nome")  ' Output: Jo�o
    '   Debug.Print user("idade") ' Output: 30
    '   Debug.Print user("ativo") ' Output: True
    '
    '   Dim array As Object
    '   Set array = ParseJSON("[""item1"",""item2"",123]")
    '   Debug.Print array(1) ' Output: item1

    state.Json = jsonString
    state.position = 1

    SkipWhitespace

    Select Case Mid(state.Json, state.position, 1)
        Case "{"
            Set ParseJSON = ParseObject
        Case "["
            Set ParseJSON = ParseArray
        Case Else
            Err.Raise vbObjectError + 1, "ParseJSON", "String JSON inv�lida - deve come�ar com '{' ou '['"
    End Select
End Function

Public Function BuildJSON(ByVal obj As Variant) As String
    ' Constr�i uma string JSON a partir de um objeto VB6
    '
    ' Args:
    '   obj (Variant): Objeto VB6 (Dictionary, Collection, ou valor primitivo)
    '
    ' Result:
    '   String: String JSON v�lida representando o objeto
    '
    ' Supported Types:
    '   - Dictionary: Convertido para objeto JSON
    '   - Collection: Convertido para array JSON
    '   - String: Convertido para string JSON com escape
    '   - Numeric: Convertido para n�mero JSON
    '   - Boolean: Convertido para true/false
    '   - Null: Convertido para null
    '
    ' Example:
    '   Dim user As Dictionary
    '   Set user = CreateJSONObject()
    '   user.Add "nome", "Jo�o"
    '   user.Add "idade", 30
    '   user.Add "ativo", True
    '   Debug.Print BuildJSON(user)
    '   ' Output: {"nome":"Jo�o","idade":30,"ativo":true}
    '
    '   Dim items As Collection
    '   Set items = CreateJSONArray()
    '   items.Add "item1"
    '   items.Add 123
    '   items.Add True
    '   Debug.Print BuildJSON(items)
    '   ' Output: ["item1",123,true]

    BuildJSON = BuildValue(obj)
End Function

' ====================================================================
' FUN��ES DE CRIA��O DE OBJETOS JSON
' ====================================================================

Public Function CreateJSONObject() As Dictionary
    ' Cria um novo objeto JSON vazio (Dictionary)
    '
    ' Result:
    '   Dictionary: Novo Dictionary vazio pronto para uso como objeto JSON
    '
    ' Example:
    '   Dim produto As Dictionary
    '   Set produto = CreateJSONObject()
    '   produto.Add "id", 1
    '   produto.Add "nome", "Notebook"
    '   produto.Add "preco", 2500.99
    '   produto.Add "disponivel", True
    '
    '   Dim jsonString As String
    '   jsonString = BuildJSON(produto)
    '   Debug.Print jsonString
    '   ' Output: {"id":1,"nome":"Notebook","preco":2500.99,"disponivel":true}

    Set CreateJSONObject = New Dictionary
End Function

Public Function CreateJSONArray() As Collection
    ' Cria um novo array JSON vazio (Collection)
    '
    ' Result:
    '   Collection: Nova Collection vazia pronta para uso como array JSON
    '
    ' Example:
    '   Dim categorias As Collection
    '   Set categorias = CreateJSONArray()
    '   categorias.Add "Eletr�nicos"
    '   categorias.Add "Inform�tica"
    '   categorias.Add "Acess�rios"
    '
    '   Dim jsonString As String
    '   jsonString = BuildJSON(categorias)
    '   Debug.Print jsonString
    '   ' Output: ["Eletr�nicos","Inform�tica","Acess�rios"]

    Set CreateJSONArray = New Collection
End Function

' ====================================================================
' FUN��ES AUXILIARES DE CONSTRU��O (BUILD)
' ====================================================================

Private Function BuildValue(ByVal value As Variant) As String
    ' Constr�i a representa��o JSON de qualquer tipo de valor
    '
    ' Args:
    '   value (Variant): Valor a ser convertido (Dictionary, Collection, String, Number, Boolean, Null)
    '
    ' Result:
    '   String: Representa��o JSON v�lida do valor
    '
    ' Raises:
    '   vbObjectError + 20: Tipo de objeto n�o suportado

    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            BuildValue = BuildObject(value)
        ElseIf TypeName(value) = "Collection" Then
            BuildValue = BuildArray(value)
        Else
            Err.Raise vbObjectError + 20, "BuildValue", "Tipo de objeto n�o suportado: " & TypeName(value)
        End If
    ElseIf IsNull(value) Then
        BuildValue = "null"
    ElseIf VarType(value) = vbBoolean Then
        BuildValue = BuildBoolean(value)
    ElseIf VarType(value) = vbString Then
        BuildValue = BuildString(CStr(value))
    ElseIf IsNumeric(value) Then
        BuildValue = BuildNumber(value)
    Else
        BuildValue = BuildString(CStr(value))
    End If
End Function

Private Function BuildObject(ByVal dict As Dictionary) As String
    ' Constr�i um objeto JSON a partir de um Dictionary
    '
    ' Args:
    '   dict (Dictionary): Dictionary contendo as propriedades do objeto
    '
    ' Result:
    '   String: String JSON representando o objeto no formato {"key":"value",...}

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
    ' Constr�i um array JSON a partir de uma Collection
    '
    ' Args:
    '   coll (Collection): Collection contendo os elementos do array
    '
    ' Result:
    '   String: String JSON representando o array no formato [value1,value2,...]

    Dim result As String
    Dim item As Variant
    Dim i As Integer

    result = "["

    For i = 1 To coll.count
        If i > 1 Then
            result = result & ","
        End If

        result = result & BuildValue(coll(i))
    Next i

    result = result & "]"
    BuildArray = result
End Function

Private Function BuildString(ByVal str As String) As String
    ' Constr�i uma string JSON com escape adequado de caracteres especiais
    '
    ' Args:
    '   str (String): String a ser codificada
    '
    ' Result:
    '   String: String JSON com caracteres de escape processados e aspas delimitadoras

    Dim result As String
    Dim i As Integer
    Dim char As String

    ' CORRE��O: Usar aspas duplas escapadas para VB6
    result = Chr(34) ' Equivale a "

    For i = 1 To Len(str)
        char = Mid(str, i, 1)

        Select Case char
            Case Chr(34): ' Aspas duplas
                result = result & "\" & Chr(34)
            Case "\":
                result = result & "\\"
            Case "/":
                result = result & "\/"
            Case vbBack:
                result = result & "\b"
            Case vbFormFeed:
                result = result & "\f"
            Case vbNewLine:
                result = result & "\n"
            Case vbCr:
                result = result & "\r"
            Case vbTab:
                result = result & "\t"
            Case Else:
                If Asc(char) < 32 Then
                    result = result & "\u" & Right("0000" & Hex(Asc(char)), 4)
                Else
                    result = result & char
                End If
        End Select
    Next i

    result = result & Chr(34) ' Equivale a "
    BuildString = result
End Function

Private Function BuildNumber(ByVal num As Variant) As String
    ' Constr�i a representa��o JSON de um n�mero
    '
    ' Args:
    '   num (Variant): N�mero a ser convertido (Integer, Long, Single, Double)
    '
    ' Result:
    '   String: Representa��o JSON do n�mero (formato americano com ponto decimal)

    If VarType(num) = vbSingle Or VarType(num) = vbDouble Then
        ' Para n�meros decimais, usar formato com ponto decimal (padr�o JSON)
        BuildNumber = Replace(CStr(num), ",", ".")
    Else
        BuildNumber = CStr(num)
    End If
End Function

Private Function BuildBoolean(ByVal bool As Boolean) As String
    ' Constr�i a representa��o JSON de um valor booleano
    '
    ' Args:
    '   bool (Boolean): Valor booleano a ser convertido
    '
    ' Result:
    '   String: "true" ou "false" (formato JSON padr�o)

    If bool Then
        BuildBoolean = "true"
    Else
        BuildBoolean = "false"
    End If
End Function

' ====================================================================
' FUN��ES AUXILIARES DE AN�LISE (PARSE)
' ====================================================================

Private Function ParseObject() As Dictionary
    ' Analisa um objeto JSON e retorna um Dictionary equivalente
    '
    ' Result:
    '   Dictionary: Dictionary contendo as propriedades do objeto JSON
    '
    ' Raises:
    '   vbObjectError + 2: Nome de propriedade esperado
    '   vbObjectError + 3: Dois pontos ':' esperados
    '   vbObjectError + 4: V�rgula ',' ou chave de fechamento '}' esperados

    Dim dict As New Dictionary
    Dim key As String

    state.position = state.position + 1 ' Skip opening '{'

    Do
        SkipWhitespace

        If Mid(state.Json, state.position, 1) = "}" Then
            state.position = state.position + 1
            Set ParseObject = dict
            Exit Function
        End If

        If Mid(state.Json, state.position, 1) <> """" Then
            Err.Raise vbObjectError + 2, "ParseObject", "Nome de propriedade esperado (string)"
        End If
        key = ParseString

        SkipWhitespace

        If Mid(state.Json, state.position, 1) <> ":" Then
            Err.Raise vbObjectError + 3, "ParseObject", "Dois pontos ':' esperados ap�s nome da propriedade"
        End If
        state.position = state.position + 1

        dict.Add key, ParseValue

        SkipWhitespace

        Select Case Mid(state.Json, state.position, 1)
            Case "}"
                state.position = state.position + 1
                Set ParseObject = dict
                Exit Function
            Case ",":
                state.position = state.position + 1
            Case Else
                Err.Raise vbObjectError + 4, "ParseObject", "V�rgula ',' ou chave de fechamento '}' esperados"
        End Select
    Loop
End Function

Private Function ParseArray() As Collection
    ' Analisa um array JSON e retorna uma Collection equivalente
    '
    ' Result:
    '   Collection: Collection contendo os elementos do array JSON
    '
    ' Raises:
    '   vbObjectError + 5: V�rgula ',' ou colchete de fechamento ']' esperados

    Dim arr As New Collection

    state.position = state.position + 1 ' Skip opening '['

    Do
        SkipWhitespace

        If Mid(state.Json, state.position, 1) = "]" Then
            state.position = state.position + 1
            Set ParseArray = arr
            Exit Function
        End If

        arr.Add ParseValue

        SkipWhitespace

        Select Case Mid(state.Json, state.position, 1)
            Case "]"
                state.position = state.position + 1
                Set ParseArray = arr
                Exit Function
            Case ",":
                state.position = state.position + 1
            Case Else
                Err.Raise vbObjectError + 5, "ParseArray", "V�rgula ',' ou colchete de fechamento ']' esperados"
        End Select
    Loop
End Function

Private Function ParseValue() As Variant
    ' Analisa qualquer valor JSON e retorna o tipo VB6 apropriado
    '
    ' Result:
    '   Variant: Valor analisado que pode ser Dictionary, Collection, String, Number, Boolean ou Null
    '
    ' Raises:
    '   vbObjectError + 6: Valor JSON inv�lido

    SkipWhitespace

    Select Case Mid(state.Json, state.position, 1)
        Case "{":
            Set ParseValue = ParseObject
        Case "[":
            Set ParseValue = ParseArray
        Case """":
            ParseValue = ParseString
        Case "t":
            ParseValue = ParseTrue
        Case "f":
            ParseValue = ParseFalse
        Case "n":
            ParseValue = ParseNull
        Case "-", "0" To "9":
            ParseValue = ParseNumber
        Case Else:
            Err.Raise vbObjectError + 6, "ParseValue", "Valor JSON inv�lido na posi��o " & state.position
    End Select
End Function

Private Function ParseString() As String
    ' Analisa uma string JSON processando caracteres de escape
    '
    ' Result:
    '   String: String decodificada com caracteres de escape processados
    '
    ' Supported Escapes:
    '   \"  \\  \/  \b  \f  \n  \r  \t  \uXXXX
    '
    ' Raises:
    '   vbObjectError + 7: Sequ�ncia de escape inv�lida
    '   vbObjectError + 8: String n�o terminada (falta aspas de fechamento)

    Dim result As String
    Dim char As String

    state.position = state.position + 1 ' Skip opening quote

    Do While state.position <= Len(state.Json)
        char = Mid(state.Json, state.position, 1)

        Select Case char
            Case """":
                state.position = state.position + 1
                ParseString = result
                Exit Function
            Case "\":
                state.position = state.position + 1
                char = Mid(state.Json, state.position, 1)

                Select Case char
                    Case """", "\", "/":
                        result = result & char
                    Case "b":
                        result = result & vbBack
                    Case "f":
                        result = result & vbFormFeed
                    Case "n":
                        result = result & vbNewLine
                    Case "r":
                        result = result & vbCr
                    Case "t":
                        result = result & vbTab
                    Case "u":
                        Dim hexCode As String
                        hexCode = Mid(state.Json, state.position + 1, 4)
                        result = result & ChrW$(CLng("&H" & hexCode))
                        state.position = state.position + 4
                    Case Else
                        Err.Raise vbObjectError + 7, "ParseString", "Sequ�ncia de escape inv�lida: \" & char
                End Select
            Case Else
                result = result & char
        End Select

        state.position = state.position + 1
    Loop

    Err.Raise vbObjectError + 8, "ParseString", "String n�o terminada - aspas de fechamento n�o encontradas"
End Function

Private Function ParseNumber() As Variant
    ' Analisa um n�mero JSON e retorna Long ou Double conforme apropriado
    '
    ' Result:
    '   Variant: Long para n�meros inteiros ou Double para decimais/cient�ficos
    '
    ' Supported Formats:
    '   Inteiros: 123, -456
    '   Decimais: 123.45, -67.89
    '   Cient�ficos: 1.23e10, -4.56E-7

    Dim numStr As String
    Dim char As String

    Do While state.position <= Len(state.Json)
        char = Mid(state.Json, state.position, 1)

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
    '   Boolean: Sempre retorna True se a sequ�ncia for v�lida
    '
    ' Raises:
    '   vbObjectError + 9: Literal "true" esperado

    If Mid(state.Json, state.position, 4) = "true" Then
        state.position = state.position + 4
        ParseTrue = True
    Else
        Err.Raise vbObjectError + 9, "ParseTrue", "Literal 'true' esperado"
    End If
End Function

Private Function ParseFalse() As Boolean
    ' Analisa o valor literal "false" do JSON
    '
    ' Result:
    '   Boolean: Sempre retorna False se a sequ�ncia for v�lida
    '
    ' Raises:
    '   vbObjectError + 10: Literal "false" esperado

    If Mid(state.Json, state.position, 5) = "false" Then
        state.position = state.position + 5
        ParseFalse = False
    Else
        Err.Raise vbObjectError + 10, "ParseFalse", "Literal 'false' esperado"
    End If
End Function

Private Function ParseNull() As Variant
    ' Analisa o valor literal "null" do JSON
    '
    ' Result:
    '   Variant: Retorna valor Null se a sequ�ncia for v�lida
    '
    ' Raises:
    '   vbObjectError + 11: Literal "null" esperado

    If Mid(state.Json, state.position, 4) = "null" Then
        state.position = state.position + 4
        ParseNull = Null
    Else
        Err.Raise vbObjectError + 11, "ParseNull", "Literal 'null' esperado"
    End If
End Function

Private Sub SkipWhitespace()
    ' Pula caracteres de espa�o em branco na string JSON
    '
    ' Advances position until a non-whitespace character is found
    ' Whitespace characters: space, tab, carriage return, line feed

    Dim char As String

    Do While state.position <= Len(state.Json)
        char = Mid(state.Json, state.position, 1)

        If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
            state.position = state.position + 1
        Else
            Exit Do
        End If
    Loop
End Sub


