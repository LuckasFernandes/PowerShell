#Script Atribuição de Assinatura em massa com scripts em Powershell
#Criado por: Luckas Fernandes - https://www.linkedin.com/in/luckas-fernandes-295a8482/
#
#Objetivo: Demonstrar de forma didática a criação de um script para atualização de Assinaturas de e-mail em massa via Outlook Office 365
#Pré-requisitos: Instalação do modulo msonline

try{
    $cred = Get-Credential                                                                  #Armazena na variável $cred os dados de acesso fornecidos pelo usuário mediante ao PopUp exibido
    Connect-MsolService -Credential $cred                                                   #Estabelece conexão com o Office365
    
    #Conecta com o tenant do Office365
    $session       = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $session -AllowClobber
        
    $list = Get-MsolUser -DomainName DOMINIO_SUA_EMPRESA -All                               #Lista todos os usuários cadastrado no Office365 no respectivo domínio da sua empresa
    
    foreach($user in $list) {
        $userPrincipalName     = $user.UserPrincipalName                                    #Armazena na variável $userPricipalName o e-mail do usuário
        $fullName              = $user.DisplayName                                          #Armazena na variável $displayName o nome completo do usuário
        $depart                = $user.Department                                           #Armazena na variável $depart o departamento que o usuário trabalha
        $phone                 = $user.PhoneNumber                                          #Armazena na variável $phone o telefone corporativo do usuário

        $outputFile            = $env:TEMP + "\" + $userPrincipalName + ".htm"              #Atribui o nome e o local aonde o arquivo html contendo a assinatura do usuário será armazenado temporariamente
        
        #Gera o arquivo HTML da assinatura com os padrões predefinidos
        "<html><head><title>Assinatura de e-mail - " + $fullName + "</title></head><body>
        <table>
            <tr>
            <td> 
                <a target=`"_blank`" rel=`"noopener noreferrer`"><img src=`"https://www.meiofiltrante.com.br/Arquivos/Noticias/56567/vulkan-do-brasil-conquista-selo-internacional-great-place-to-work.jpg`" style=`"width: 40px; padding-right: 15px;`"></a>
            </td>
            <td align=`"left`" valign=`"center`" style=`"border-left: 5px solid; color:rgb(79,79,79); padding-left: 10px`">
                <p style=`"font-family: Arial, Helvetica, Verdana;font-size:x-small;color:rgb(0,79,79)`">
                <strong>" + $fullName + "</strong><br>" + $depart + "<br>" + $location + "<br>" + $phone + " | <a href=`"SITE`" target=_blank>www.DOMINIO_EMPRESA.com.br</a><br>
                </p>
            </td>
            </tr>
        </table></body></html>" | Out-File $outputFile
                        
        #Registra nova assinatura na caixa de e-mail do Outlook 365
        set-mailboxmessageconfiguration -identity $userPrincipalName -signaturehtml (get-content $outputFile) -autoaddsignature $true 
    
        #Exibe no Prompt a mensagem de que a operação foi concluída com sucesso
        echo "A assinatura de e-mail do usuário $userPrincipalName foi atribuída ao Outlook."
    } 
} catch {  
    #Exibe no Prompt a mensagem de que a operação não foi concluída com sucesso
    echo "Foi identificado algum erro ao tentar atribuir ao Outlook a assinatura de e-mail do usuário $userPrincipalName."
}