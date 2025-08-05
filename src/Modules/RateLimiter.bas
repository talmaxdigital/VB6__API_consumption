Attribute VB_Name = "RateLimiter"
Option Explicit

' ====================================================================
' RateLimiter Module - Controle de taxa de requisições para APIs
' ====================================================================

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Função para limitar taxa de requisições
' Use esta função antes de fazer requisições API que precisam de limitação de taxa
Public Sub LimitRequestRate(ByRef timestamps As Collection, ByVal maxRequestsPerSecond As Long)
    ' Limita a taxa de requisições por segundo
    '
    ' Args:
    '   timestamps (Collection): Coleção com timestamps das requisições
    '   maxRequestsPerSecond (Long): Número máximo de requisições permitidas por segundo

    Const ONE_SECOND As Double = 1# / (24# * 60# * 60#) ' 1 segundo em formato de data do VB
    Dim currentTime As Date
    currentTime = Now

    ' Remove timestamps mais antigos que 1 segundo
    Do While timestamps.Count > 0
        If DateDiff("s", timestamps(1), currentTime) >= 1 Then
            timestamps.Remove 1
        Else
            Exit Do
        End If
    Loop

    ' Se atingiu o limite de requisições por segundo, aguarda
    If timestamps.Count >= maxRequestsPerSecond Then
        Dim sleepTime As Double
        sleepTime = ONE_SECOND - DateDiff("s", timestamps(1), currentTime) / 86400#

        ' Converte para milissegundos e dorme
        Sleep CLng(sleepTime * 86400# * 1000#)
    End If

    ' Adiciona novo timestamp
    timestamps.Add Now
End Sub

