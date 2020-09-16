#Script Interface Gráfica com scripts em Powershell
#Criado por: Luckas Fernandes - https://www.linkedin.com/in/luckas-fernandes-295a8482/
#
#Objetivo: Demonstrar de forma didática a criação de uma GUI para automatização de tarefas com Powershell

Add-Type -AssemblyName System.Windows.Forms                                                            #Importa a biblioteca "System.Windows.Forms"

#Configurando o formulário principal da GUI
$gui               = New-Object System.Windows.Forms.Form                                              #Cria o objeto do formulário principal
$gui.Text          = "Script GUI"                                                                      #Atribui um título a janela
$gui.Size          = New-Object System.Drawing.Size(255,180)                                           #Atribui o valor da altura e largura respectivamente a janela
$gui.StartPosition = 'CenterScreen'                                                                    #Indica ao sistema que inicie a janela no centro da tela

#Atribuindo componentes ao formulário
#
#Atrubuindo Label
$label             = New-Object System.Windows.Forms.Label                                             #Cria o objeto do label
$label.Text        = "Processo MySQL:"                                                                 #Atribui o texto ao componente label
$label.Size        = New-Object System.Drawing.Size(220,30)                                            #Atribui o valor da largura e altura respectivamente ao componente label
$label.Location    = New-Object System.Drawing.Point(10,10)                                            #Define o posicionamento do componente label atribuindo a distância da margem esquerda e do topo respectivamente
$label.TextAlign   = "TopCenter"                                                                       #Define o alinhamento do texto dentro do objeto
$gui.Controls.Add($label)                                                                              #Adiciona o item ao formulário

#Atrubuindo Button1
$button1           = New-Object System.Windows.Forms.Button                                            #Cria o objeto do Button1
$button1.Text      = "Start"                                                                           #Atribui o texto ao componente Button1
$button1.Size      = New-Object System.Drawing.Size(220,30)                                            #Atribui o valor da largura e altura respectivamente ao componente Button1
$button1.Location  = New-Object System.Drawing.Point(10,40)                                            #Define o posicionamento do componente Button1 atribuindo a distância da margem esquerda e do topo respectivamente
$button1.add_Click(${function:button1OnClick})                                                         #Atribuição a função $button1OnClick para gerar a interação com o usuário
$gui.Controls.Add($button1)                                                                            #Adiciona o item ao formulário

#Atrubuindo Button2
$button2           = New-Object System.Windows.Forms.Button                                            #Cria o objeto do Button2
$button2.Text      = "Stop"                                                                            #Atribui o texto ao componente Button2
$button2.Size      = New-Object System.Drawing.Size(220,30)                                            #Atribui o valor da largura e altura respectivamente ao componente Button2
$button2.Location  = New-Object System.Drawing.Point(10,70)                                            #Define o posicionamento do componente Button2 atribuindo a distância da margem esquerda e do topo respectivamente
$button2.add_Click(${function:button2OnClick})                                                         #Atribuição a função $button2OnClick para gerar a interação com o usuário
$gui.Controls.Add($button2)                                                                            #Adiciona o item ao formulário

#Atrubuindo Button3
$button3           = New-Object System.Windows.Forms.Button                                            #Cria o objeto do Button3
$button3.Text      = "Restart"                                                                         #Atribui o texto ao componente Button3
$button3.Size      = New-Object System.Drawing.Size(220,30)                                            #Atribui o valor da largura e altura respectivamente ao componente Button3
$button3.Location  = New-Object System.Drawing.Point(10,100)                                           #Define o posicionamento do componente Button3 atribuindo a distância da margem esquerda e do topo respectivamente
$button3.add_Click(${function:button3OnClick})                                                         #Atribuição a função $button2OnClick para gerar a interação com o usuário
$gui.Controls.Add($button3)                                                                            #Adiciona o item ao formulário

#Atribuindo a ação que o script executará mediante a interação do usuário
#
#Função responsável por executar a ação que o correspondente ao button1
Function button1OnClick {
    try{
        if(Get-Service -Name mysql -ComputerName WinServer){                                           #Verifica se o serviço mysql existe no WinServer
            Get-Service -Name mysql -ComputerName WinServer | Start-Service                            #cmdlet responsável por iniciar o serviço MySql no servidor WinServer
            [System.Windows.Forms.MessageBox]::Show("Serviço MySQL iniciado.", "SERVIÇO INICIADO")     #PopUp informando que a operação foi concluída
        } else {
            [System.Windows.Forms.MessageBox]::Show("Serviço MySql não localizado.", "ERROR")          #PopUp informando que o serviço MySql não foi localizado          
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao iniciar o serviço MySql.", "ERROR")          #PopUp informando erro na operação          
    }
}

#Função responsável por executar a ação que o correspondente ao button2
Function button2OnClick {
    try{
        if(Get-Service -Name mysql -ComputerName WinServer){                                           #Verifica se o serviço MySql existe no WinServer 
            Get-Service -Name mysql -ComputerName WinServer | Stop-Service                             #cmdlet responsável por parar o serviço MySql no servidor WinServer
            [System.Windows.Forms.MessageBox]::Show("Serviço MySQL parado.", "SERVIÇO PARADO")         #PopUp informando que a operação foi concluída
        } else {
            [System.Windows.Forms.MessageBox]::Show("Serviço MySql não localizado.", "ERROR")          #PopUp informando que o serviço MySql não foi localizado           
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao parar o serviço MySql.", "ERROR")            #PopUp informando erro na operação               
    }
}

#Função responsável por executar a ação que o correspondente ao button3
Function button3OnClick {
    try{
        if(Get-Service -Name mysql -ComputerName WinServer){                                           #Verifica se o serviço mysql existe no WinServer
            Get-Service -Name mysql -ComputerName WinServer | Restart-Service                          #cmdlet responsável por Reiniciar o serviço MySql no servidor WinServer
            [System.Windows.Forms.MessageBox]::Show("Serviço MySQL reiniciado.", "SERVIÇO REINICIADO") #PopUp informando que a operação foi concluída
        } else {
            [System.Windows.Forms.MessageBox]::Show("Serviço MySql não localizado.", "ERROR")          #PopUp informando que o serviço MySql não foi localizado          
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao reiniciar o serviço MySql.", "ERROR")        #PopUp informando erro na operação         
    }
}

# Inicia o formulario
$gui.ShowDialog()                                                                                      # Desenha na tela todos os componentes adicionados ao formulario
