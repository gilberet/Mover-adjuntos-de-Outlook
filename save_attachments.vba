Option Explicit

Private WithEvents oItems As Outlook.Items
Private Sub Application_Startup()
    Dim OutlookApp As Outlook.Application
    Dim oNameSpace As Outlook.NameSpace
    
    
    Set OutlookApp = Outlook.Application
    Set oNameSpace = Outlook.GetNamespace("MAPI") ' Messaging Application Programaming Interface
    Set oItems = oNameSpace.GetDefaultFolder(olFolderInbox).Items '' Tomamos los correos que llegan a la Bandeja de Entreda solamente
    
    Debug.Print "Desencadenador iniciado " & VBA.Now
End Sub

Private Sub oItems_ItemAdd(ByVal Item As Object)
    Dim MyMail As Outlook.MailItem
    Dim oAtt As Outlook.Attachment
    Dim objShell, objFSO, filesInzip As Object
    Dim saveFolder, FullFileName As String
    
    
    '' Iniciacion variables
    Set objShell = CreateObject("Shell.Application")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    saveFolder = "\\172.17.0.160\c$\ReportesWMS\"
    
    Debug.Print "Inicio adjunto " & VBA.Now
    If VBA.TypeName(Item) = "MailItem" Then
        Set MyMail = Item
        
        '' Reporte de cajas **************************************
        Debug.Print "Recorriendo los correos " & VBA.Now
        If MyMail.Subject Like "*" & "Reporte de cajas" & "*" Then
            For Each oAtt In MyMail.Attachments
                oAtt.SaveAsFile (saveFolder & oAtt.DisplayName)
                Debug.Print "Se guardo Reporte de Cajas en la Carpeta"
                Set oAtt = Nothing
            Next oAtt
        End If
        
        '' Ordenes Enviadas **************************************
        If MyMail.Subject Like "*" & "Ordenes Enviadas" & "*" Then
            For Each oAtt In MyMail.Attachments
                If InStr(oAtt.FileName, ".csv") Then
                    oAtt.SaveAsFile saveFolder & oAtt.FileName
                    Debug.Print "Se guardo Ordenes Enviadas en la Carpeta"
                Else
                    If ((InStr(UCase(oAtt.DisplayName), ".ZIP"))) Then
                        
                        Debug.Print "Archivo Zip " & oAtt.FileName
                        FullFileName = saveFolder & oAtt.FileName
                        oAtt.SaveAsFile (FullFileName)
                        Set filesInzip = objShell.NameSpace(FullFileName & "\").Items
                        objShell.NameSpace(saveFolder).CopyHere filesInzip
                        objFSO.DeleteFile (FullFileName)
                        Debug.Print "Se guardo Ordenes Enviadas en la Carpeta"
                    End If
                End If
            
                
                Set oAtt = Nothing
            Next oAtt
        End If
        
        
        Set MyMail = Nothing
    End If
    Debug.Print "Finalizo adjunto " & VBA.Now
End Sub
