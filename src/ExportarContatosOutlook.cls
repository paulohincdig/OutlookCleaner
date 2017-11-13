'
Sub ExportarContatos()

On Error Resume Next

Dim fso As New Scripting.FileSystemObject
Dim arquivo As TextStream
Dim campos() As String
Dim contato As ExchangeUser

Const DESTINATION_PATH As String = "c:\temp\emails.txt"

Set arquivo = fso.OpenTextFile(DESTINATION_PATH, ForWriting, True)

ReDim campos(5)

campos(0) = "Cadastro"
campos(1) = "Login"
campos(2) = "Nome"
campos(3) = "Email"
campos(4) = "Cargo"
campos(5) = "Departamento"

arquivo.WriteLine Join(campos, vbTab)
            
For Each i In Session.AddressLists

    'If i.Name = "All Users" or i.Name =  Then
    
        For Each o In i.AddressEntries
            
            ReDim campos(5)
            
            Set contato = o.GetExchangeUser
                       
            campos(0) = contato.OfficeLocation
            campos(1) = contato.Alias
            campos(2) = contato.Name
            campos(3) = contato.PrimarySmtpAddress
            campos(4) = contato.JobTitle
            campos(5) = contato.Department
                    
            arquivo.WriteLine Join(campos, vbTab)
            
            If Err.Number <> 0 Then
                Debug.Print Err.Description
                Err.Clear
            End If

        Next
    
    'End If

Next

arquivo.Close

Set arquivo = Nothing

End Sub
