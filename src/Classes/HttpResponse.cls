VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' HttpResponse Class - Encapsula respostas HTTP
' ====================================================================

Private m_StatusCode As Long
Private m_StatusText As String
Private m_ResponseText As String
Private m_ResponseHeaders As Dictionary
Private m_JsonData As Object

Public Sub Initialize(ByVal xmlHttp As Object)
    ' Inicializa a resposta a partir do objeto XMLHTTP
    '
    ' Args:
    '   xmlHttp (Object): Objeto XMLHTTP com a resposta

    m_StatusCode = xmlHttp.status
    m_StatusText = xmlHttp.StatusText
    m_ResponseText = xmlHttp.responseText

    Set m_ResponseHeaders = CreateJSONObject()
    ParseResponseHeaders xmlHttp.getAllResponseHeaders()
End Sub

Public Property Get StatusCode() As Long
    ' Código de status HTTP da resposta
    StatusCode = m_StatusCode
End Property

Public Property Get StatusText() As String
    ' Texto do status HTTP da resposta
    StatusText = m_StatusText
End Property

Public Property Get Text() As String
    ' Corpo da resposta como texto
    Text = m_ResponseText
End Property

Public Property Get IsSuccess() As Boolean
    ' Indica se a requisição foi bem-sucedida (status 200-299)
    IsSuccess = (m_StatusCode >= 200 And m_StatusCode <= 299)
End Property

Public Property Get Json() As Object
    ' Retorna o corpo da resposta parseado como JSON
    '
    ' Result:
    '   Object: Dictionary ou Collection com dados JSON parseados

    If m_JsonData Is Nothing Then
        If Len(m_ResponseText) > 0 Then
            Set m_JsonData = ParseJSON(m_ResponseText)
        End If
    End If

    Set Json = m_JsonData
End Property

Public Function GetHeader(ByVal headerName As String) As String
    ' Obtém o valor de um header específico
    '
    ' Args:
    '   headerName (String): Nome do header
    '
    ' Result:
    '   String: Valor do header ou string vazia se não encontrado

    If m_ResponseHeaders.Exists(headerName) Then
        GetHeader = m_ResponseHeaders(headerName)
    Else
        GetHeader = ""
    End If
End Function

Private Sub ParseResponseHeaders(ByVal headersText As String)
    ' Analisa os headers da resposta
    '
    ' Args:
    '   headersText (String): Texto com todos os headers

    Dim lines As Variant
    Dim i As Integer
    Dim colonPos As Integer
    Dim headerName As String
    Dim headerValue As String

    lines = Split(headersText, vbCrLf)

    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            colonPos = InStr(lines(i), ":")
            If colonPos > 0 Then
                headerName = Trim(Left(lines(i), colonPos - 1))
                headerValue = Trim(Mid(lines(i), colonPos + 1))

                If Not m_ResponseHeaders.Exists(headerName) Then
                    m_ResponseHeaders.Add headerName, headerValue
                End If
            End If
        End If
    Next i
End Sub
