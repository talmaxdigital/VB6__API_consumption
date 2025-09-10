VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' HttpRequest Class - Wrapper para XMLHTTP
' ====================================================================

Private m_XmlHttp As Object
Private m_Timeout As Long

Private Sub Class_Initialize()
    ' Inicializa o objeto XMLHTTP
    Set m_XmlHttp = CreateObject("MSXML2.XMLHTTP")
    m_Timeout = 30000 ' 30 segundos padrão
End Sub

Public Sub Open_(ByVal method As String, ByVal url As String, Optional ByVal async As Boolean = False)
    ' Abre uma conexão HTTP
    '
    ' Args:
    '   method (String): Método HTTP (GET, POST, PUT, DELETE, etc.)
    '   url (String): URL completa da requisição
    '   async (Boolean): Se True, requisição assíncrona (padrão: False)

    m_XmlHttp.Open method, url, async
End Sub

Public Sub SetRequestHeader(ByVal headerName As String, ByVal headerValue As String)
    ' Define um header da requisição
    '
    ' Args:
    '   headerName (String): Nome do header
    '   headerValue (String): Valor do header

    m_XmlHttp.SetRequestHeader headerName, headerValue
End Sub

Private Function GetRequestHeader(ByVal headerName As String) As String
    ' Função auxiliar para verificar se header já foi definido

    GetRequestHeader = ""
End Function

Public Sub Send(Optional ByVal body As Variant)
    ' Envia a requisição HTTP
    '
    ' Args:
    '   body (Variant): Corpo da requisição (opcional)

    ' CORREÇÃO: Garantir Content-Type correto antes de enviar
    If Not IsMissing(body) And VarType(body) = vbString Then
        If Len(CStr(body)) > 0 Then
            ' Verificar se é JSON pela primeira e última caractere
            Dim bodyStr As String
            bodyStr = CStr(body)
            If (Left(bodyStr, 1) = "{" And Right(bodyStr, 1) = "}") Or _
               (Left(bodyStr, 1) = "[" And Right(bodyStr, 1) = "]") Then
                ' Garantir Content-Type para JSON
                m_XmlHttp.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
            End If
        End If
    End If

    If IsMissing(body) Then
        m_XmlHttp.Send
    Else
        m_XmlHttp.Send body
    End If

    ' Aguardar resposta se síncrono
    Do While m_XmlHttp.readyState <> 4
        DoEvents
    Loop
End Sub

Public Sub SetTimeout(ByVal timeoutMs As Long)
    ' Define o timeout da requisição
    '
    ' Args:
    '   timeoutMs (Long): Timeout em milissegundos

    m_Timeout = timeoutMs
    ' Note: XMLHTTP pode não suportar timeout diretamente
    ' Implementação pode variar conforme a versão
End Sub

Public Property Get status() As Long
    ' Código de status HTTP da resposta
    status = m_XmlHttp.status
    Debug.Print "Status: " & status
End Property

Public Property Get StatusText() As String
    ' Texto do status HTTP da resposta
    StatusText = m_XmlHttp.StatusText
    Debug.Print "StatusText: " & StatusText
End Property

Public Property Get responseText() As String
    ' Corpo da resposta como texto
    responseText = m_XmlHttp.responseText
    Debug.Print "responseText: " & responseText
End Property

Public Function getAllResponseHeaders() As String
    ' Retorna todos os headers da resposta
    getAllResponseHeaders = m_XmlHttp.getAllResponseHeaders()
End Function

Public Property Get readyState() As Long
    ' Estado atual da requisição
    readyState = m_XmlHttp.readyState
End Property
