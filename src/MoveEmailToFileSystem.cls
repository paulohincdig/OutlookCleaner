'processo de postagem de arquivos recebidos por e-mail em uma pasta compartilhada na rede intranet
Sub MoverArquivo()

Dim fso As New Scripting.FileSystemObject
Dim encontrado As Outlook.mailitem
Dim file_msg_path As String
Const TEMP_FOLDER As String = "c:\temp\"
Const DESTINATION_FOLDER As String = "\\AL0003NAMP001\ArquivosVivo\"
Const ENTRYID_INBOX As String = "00000000DC33679E8835C84F9132082DA4824D840100E93584A254A9B04D97ACA3A73AE152160000021D683E0000"
Const ENTRYID_BASE As String = "00000000A1C9801B19064141967D6542B9F99FBE42810000"
Const BASE_EXTENSION As String = "txt"
Const ASSUNTO As String = "[ASSUNTO DO E-MAIL PROCURADO]"

Do

continua = False
Set mails = Session.GetFolderFromID(ENTRYID_INBOX).Items

filtro = "[Subject] = """ & ASSUNTO & """"

    Set encontrados = mails.Restrict(filtro)
    
    For Each encontrado In encontrados
        
        continua = True
        
        Set arquivo = encontrado.Attachments.Item(1)
        
        Prefix = Format(encontrado.ReceivedTime, "YYYYMMDDHHnnss_")
        
        file_msg_path = TEMPO_FOLDER & Prefix & arquivo.FileName
        
        If fso.FileExists(file_msg_path) Then
            fso.DeleteFile file_msg_path
        End If
        
        arquivo.SaveAsFile file_msg_path
        
        Set novo = Session.OpenSharedItem(file_msg_path)
        Set txt = novo.Attachments.Item(1)
        
        If LCase(txt.FileName) Like "*." & LCase(BASE_EXTENSION) Then
        
            file_txt_path = TEMP_FOLDER & Prefix & txt.FileName
            
            txt.SaveAsFile file_txt_path
            
            file_destination_path = Replace(file_txt_path, TEMP_FOLDER, DESTINATION_FOLDER)
            
            If Not fso.FileExists(file_almawave_path) Then
                fso.MoveFile file_txt_path, file_destination_path
                encontrado.Move Session.GetFolderFromID(ENTRYID_BASE)
                Debug.Print "Arquivo " & file_msg_path & " movido para a pasta de bases"
            Else
                '
                If FileEquals(fso, file_txt_path, file_destination_path) Then
                    Debug.Print "Arquivo " & file_msg_path & " já foi postado.", vbCritical
                Else
                    MsgBox "Um erro ocorreu ao tentar postar o arquivo " & file_msg_path, vbCritical
                End If
            End If
        End If

        novo.Close 1
        
        Set novo = Nothing
        Set txt = Nothing
        
        DeleteFileIfExists fso, file_msg_path
        
    Next
    
    Set mails = Nothing
    
Loop While continua

End Sub
'
Function FileEquals(ByRef fso As Scripting.FileSystemObject, ByVal a As String, ByVal b As String) As Boolean

Dim arquivo1 As file, arquivo2 As file

Set arquivo1 = fso.GetFile(a)
Set arquivo2 = fso.GetFile(b)

FileEquals = arquivo1.Name = arquivo2.Name _
        And arquivo1.Size = arquivo2.Size _
        And arquivo1.DateCreated = arquivo2.DateCreated

End Function
'
Sub DeleteFileIfExists(ByRef fso As Scripting.FileSystemObject, ByVal file As String)
If fso.FileExists(file) Then
    fso.DeleteFile file
End If
End Sub
