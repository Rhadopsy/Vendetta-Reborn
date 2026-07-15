Attribute VB_Name = "ModCompat"
Option Explicit

' Centralise les composants COM facultatifs afin que leur absence ne bloque pas
' le lancement du jeu sous Windows moderne ou Proton.

Public Function CompatCreerMoteurScript() As Object
    On Error Resume Next
    Set CompatCreerMoteurScript = CreateObject("MSScriptControl.ScriptControl")
    If Not CompatCreerMoteurScript Is Nothing Then
        CompatCreerMoteurScript.Language = "VBScript"
        CompatCreerMoteurScript.Reset
        CompatCreerMoteurScript.AllowUI = False
    End If
    On Error GoTo 0
End Function

Public Function CompatCreerDocumentXML() As Object
    On Error Resume Next
    Set CompatCreerDocumentXML = CreateObject("MSXML2.DOMDocument.6.0")
    If CompatCreerDocumentXML Is Nothing Then
        Set CompatCreerDocumentXML = CreateObject("MSXML2.DOMDocument.3.0")
    End If
    On Error GoTo 0
End Function

Private Function CompatCreerRequeteHttp() As Object
    On Error Resume Next
    Set CompatCreerRequeteHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    If CompatCreerRequeteHttp Is Nothing Then
        Set CompatCreerRequeteHttp = CreateObject("MSXML2.XMLHTTP.3.0")
    End If
    On Error GoTo 0
End Function

Public Function CompatHttpGet(ByVal URL As String) As String
    Dim Requete As Object

    On Error GoTo Erreur
    Set Requete = CompatCreerRequeteHttp()
    If Requete Is Nothing Then Exit Function
    Requete.Open "GET", URL, False
    Requete.Send

    If Requete.Status >= 200 And Requete.Status < 300 Then
        CompatHttpGet = Requete.responseText
    End If

    Set Requete = Nothing
    Exit Function

Erreur:
    CompatHttpGet = ""
    Set Requete = Nothing
End Function
