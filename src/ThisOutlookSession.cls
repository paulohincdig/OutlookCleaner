Dim excluidos As Long
Const OBJECT_ID as String = "00000000F1699FF7A6B40F4E9300335B19D9BE5F82810000"
Const URL_HOST as String = "http://URL"
Const SENDER_NAME as String = "MIS"
Const FILE_TO_EXPORT as String = "c:\received.txt"

Sub ExcluirEmailsAntigos()

Dim inicio As Date
inicio = Now

excluidos = 0

Do

encontrado = False

Set mails = Session.GetFolderFromID(OBJECT_ID).Items

mails.Sort "[ReceivedTime]", True

For Each mail In mails
    
    If mail.SenderName <> SENDER_NAME Then
        GoTo continue
    End If
    
    filtro = "[Subject] = """ & mail.Subject & """"

    Set encontrados = mails.Restrict(filtro)

    For Each i In encontrados

        entryID = Empty
    
        On Error Resume Next
        
        If i.ReceivedTime < mail.ReceivedTime Then
            Debug.Print Format(Now, "dd/mm/yyyy HH:mm:ss") & " " & i.Subject & " (" & i.ReceivedTime & ") - Excluído."
            DoEvents
            i.Delete
            excluidos = excluidos + 1
            encontrado = True
        Else
            Debug.Print Format(Now, "dd/mm/yyyy HH:mm:ss") & " " & i.Subject & " (" & i.ReceivedTime & ") - Último."
        End If
        
        On Error GoTo 0
        
    Next

continue:

Next

Loop While encontrado

Debug.Print Format(Now, "dd/mm/yyyy HH:mm:ss") & " - Concluído. Tempo total : " & Format(Now - inicio, "hh:mm:ss")

MsgBox excluidos & " e-mails excluídos." & vbNewLine & "Tempo total : " & Format(Now - inicio, "hh:mm:ss")

End Sub

Sub ExportarBase()

Dim fSO As New Scripting.FileSystemObject
Dim arquivo As TextStream
Dim campos()

tmp = Environ$("temp") & "\" & fSO.GetTempName

Set mails = Session.GetFolderFromID(OBJECT_ID).Items
Set arquivo = fSO.OpenTextFile(tmp, ForWriting, True)

For Each mail In mails
    
    ReDim campos(2)
    
    campos(0) = GetIdFromBody(mail.body)
    campos(1) = mail.Subject
    campos(2) = mail.ReceivedTime
    
    arquivo.WriteLine Join(campos, vbTab)

Next

arquivo.Close

If fSO.FileExists(FILE_TO_EXPORT) Then fSO.DeleteFile FILE_TO_EXPORT, True

fSO.MoveFile tmp, FILE_TO_EXPORT

Shell "notepad.exe " & FILE_TO_EXPORT

End Sub

Function GetIdFromBody(ByVal body As String) As String

GetIdFromBody = 0

i = InStr(1, LCase(body), URL_HOST)

If i = 0 Then Exit Function

f = InStr(i, LCase(body), "<")

If f = 0 Then f = InStr(i, LCase(body), Chr(10))
If f = 0 Then f = InStr(i, LCase(body), Chr(34))

If f = 0 Then
    Exit Function
End If

retorno = Mid(body, i, f - i)
c = InStrRev(retorno, "=", -1)

If c = 0 Then Exit Function

GetIdFromBody = ConvertToLong(Mid(retorno, c + 1, Len(retorno)))

End Function

Function ConvertToLong(ByVal texto As String) As Long

retorno = Empty

    For i = 1 To Len(texto)
        If InStr(1, "0123456789", Mid(texto, i, 1)) > 0 Then
            retorno = retorno & Mid(texto, i, 1)
        End If
    Next

ConvertToLong = retorno

End Function
