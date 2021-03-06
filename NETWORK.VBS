' ------------------------------------------------------------------------                                    
' Autor:             Isac Canedo                                                 
' Data Cria��o:      01/09/2009  
' ------------------------------------------------------------------------
'
' Este exemplo demonstra como usar o objeto WSHNetwork.
' Ele l� as propriedades de rede (nome do usu�rio e do computador), 
' conecta, desconecta e enumera as unidades de rede.


L_Welcome_MsgBox_Message_Text   = "Este script demonstra como usar o objeto WSHNetwork."
L_Welcome_MsgBox_Title_Text     = "Exemplo do Windows Scripting Host"
Call Welcome()

' ********************************************************************************
' *
' * Objeto WSH Network.
' *

Dim WSHNetwork
Dim colDrives, SharePoint
Dim CRLF

CRLF = Chr(13) & Chr(10)
Set WSHNetwork = WScript.CreateObject("WScript.Network")


Function Ask(strAction)

   ' Esta fun��o pergunta se o usu�rio quer executar uma determinada "A��o"
   ' e define um c�digo de retorno ou encerra a execu��o do script, dependendo 
   ' do bot�o que o usu�rio pressionar. Esta fun��o � chamada em v�rias
   ' partes no script abaixo.

    Dim intButton
    intButton = MsgBox(strAction,                   _
                       vbQuestion + vbYesNo,        _
                       L_Welcome_MsgBox_Title_Text )
    Ask = intButton = vbYes
End Function

' **************************************************
' *
' * Mostrar as propriedades do objeto WSHNetwork
' *
' *
MsgBox "UserDomain"     & Chr(9) & "= " & WSHNetwork.UserDomain  & CRLF &   _
       "UserName"       & Chr(9) & "= " & WSHNetwork.UserName    & CRLF &   _
       "ComputerName"   & Chr(9) & "= " & WSHNetwork.ComputerName,          _
       vbInformation + vbOKOnly,                                            _
       "Propriedades do WSHNetwork"

' **************************************************
' *
' * WSHNetwork.AddNetworkDrive
' *
' *

Function TryMapDrive(intDrive, strShare)
    Dim strDrive
    strDrive = Chr(intDrive + 64) & ":"
    On Error Resume Next
    WSHNetwork.MapNetworkDrive strDrive, strShare
    TryMapDrive = Err.Number = 0
End Function

If Ask("Deseja conectar-se a uma unidade de rede?") Then
    strShare = InputBox("Digite o compartilhamento de rede ao qual deseja conectar-se ")
    For intDrive = 26 To 5 Step -1
        If TryMapDrive(intDrive, strShare) Then Exit For
    Next

    If intDrive <= 5 Then
        MsgBox "N�o foi poss�vel conectar ao compartilhamento de rede. "                        & _
               "N�o h� letras de unidades dispon�veis para usar. "    & _
               CRLF                                                          & _
               "Desconecte uma de suas conex�es de rede existentes " & _
               "e tente este script novamente. ",                                  _
               vbExclamation + vbOkOnly,        _
               L_Welcome_MsgBox_Title_Text
    Else
        strDrive = Chr(intDrive + 64) & ":"
        MsgBox " " & strShare & " conectada � unidade " & strDrive,   _
               vbInformation + vbOkOnly,                            _
               L_Welcome_MsgBox_Title_Text

        If Ask("Deseja desconectar-se da unidade de rede que voc� acabou de criar?") Then
            WSHNetwork.RemoveNetworkDrive strDrive

            MsgBox "Unidade desconectada " & strDrive,        _
                   vbInformation + vbOkOnly,                _
                   L_Welcome_MsgBox_Title_Text
        End If
    End If
End If


' **************************************************
' *
' * WSHNetwork.EnumNetworkDrive
' *
' *
'Perguntar se o usu�rio quer enumerar as unidades de rede
If Ask("Deseja enumerar as unidades de rede conectadas?") Then
    'Enumera as unidades de rede em um objeto collection do tipo WshCollection
    Set colDrives = WSHNetwork.EnumNetworkDrives

    'Informa o usu�rio se nenhuma unidade de rede foi enumerada ou exibe as 
    'unidades enumeradas
    If colDrives.Count = 0 Then
        MsgBox "N�o h� unidades para enumerar.",     _
               vbInformation + vbOkOnly,                _
               L_Welcome_MsgBox_Title_Text
    Else
        strMsg = "Conex�es de unidade de rede atuais: " & CRLF
        For i = 0 To colDrives.Count - 1 Step 2
            strMsg = strMsg & CRLF & colDrives(i) & Chr(9) & colDrives(i + 1)
        Next
        
        MsgBox strMsg,                                  _
               vbInformation + vbOkOnly,                _
               L_Welcome_MsgBox_Title_Text

    End If
End If

' ********************************************************************************
' *
' * Bem-vindo
' *
Sub Welcome()
    Dim intDoIt

    intDoIt =  MsgBox(L_Welcome_MsgBox_Message_Text, _
                      vbOKCancel + vbInformation,    _
                      L_Welcome_MsgBox_Title_Text )
    If intDoIt = vbCancel Then
        WScript.Quit
    End If
End Sub


