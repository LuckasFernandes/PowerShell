#Script Tábuas de Marés com scripts em Powershell
#Criado por: Luckas Fernandes - https://www.linkedin.com/in/luckas-fernandes-295a8482/
#
#Objetivo: Automatizar o envio de relatórios semanais com a Tábuas de Marés acima de 0.9m

#Request HTTP solicitando dados de Paraty no ano de 2020 ao site ClimaTempo
$httpRequest = Invoke-WebRequest -Uri "https://www.climatempo.com.br/json/previsao-mare" `
-Method "POST" `
-Headers @{
"method"="POST"
  "authority"="www.climatempo.com.br"
  "scheme"="https"
  "path"="/json/previsao-mare"
  "content-length" = "24"
  "sec-ch-ua"="`"Google Chrome`";v=`"87`", `" Not;A Brand`";v=`"99`", `"Chromium`";v=`"87`""
  "sec-ch-ua-mobile"="?0"
  "user-agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"
  "content-type" = "application/x-www-form-urlencoded"
  "accept"="*/*"
  "origin"="https://www.climatempo.com.br"
  "sec-fetch-site"="same-origin"
  "sec-fetch-mode"="cors"
  "sec-fetch-dest"="empty"
  "referer"="https://www.climatempo.com.br/tabua-de-mares"
  #"accept-encoding"="gzip, deflate, br"
  "accept-language"="pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7"
  "cookie"="__cfduid=dca5f2be3e023f87d1c321b2e3d80fb211608611237; _ga=GA1.3.61479118.1608611238; _gid=GA1.3.589515134.1608611238; _hjTLDTest=1; _hjid=d1efbe39-dd4e-484f-a6d9-27c95973478c; ___ws-sr=https://www.google.com/; __qca=P0-1096655403-1608611241415; tt.u=0100007F7FF6F95E8D069FBF02E53203; ortcsession-w5tlOg-s=fc2ca00415a7c744; cookieConsent=true; _hjIncludedInSessionSample=1; nvg56295=bc534bf094199b72ef550907109|0_359; ___ws_vis=DF67CD28290914A1.1608695847410; ___ws_vis_sec=4414:1608695847410; FCCDCF=[[`"AKsRol_bDAxrLjJzyN1xZiT5A1QUdXiBDci4dNkrW_o4R0L-XOOT0GrMG5AEBJr4Bzmlf-rl8kEgElkLAw7P8oo4iIHeomyw1CoJuhTeQlXLTObFLbpBIbzQzJFjribDhSz5Yi3nNt6isixveZ-XznaYD_femgTx_g==`"],null,[`"[[],[],[],[],null,null,true]`",1608695847636]]; _ttuu.s=1608695848234; _ttdmp=E:1|S:47,51|A:3|X:3|LS:72,24,34|CA:CA18539; ortcsession-w5tlOg=fc2ca00415a7c744"
} `
-ContentType "application/x-www-form-urlencoded" `
-Body "idLocale=80810&year=2020"

#Recebe dados
$base = $httpRequest.Content

#Trata dados
$base = $base.Replace(',threshold:null,dateBegin:2020-12-23 00:00:00,dateEnd:2020-12-31 23:59:59,idStation:null]','')
$base = $base.Replace('{"success":true,"message":null,"time":null,"totalRows":null,"totalPages":null,"page":null,"data":[{"locale":{"id":80810,"name":"Parati","idStation":null,"latitude":null,"longitude":null,"type":{"id":43,"name":"LOCALIDADE MAR\u00c9"},"acronym":null,"region":null,"city":null,"uf":null,"country":null,"delivery":null,"geometricArea":null},"weather":[','')
$base = $base.Replace('"},{"','"};{"')
$base = $base -split ";"
$base = $base.Replace('{"tide":','')
$base = $base.Replace('"height":','')
$base = $base.Replace('"date":"','')
$base = $base.Replace('}','')
$base = $base.Replace(']','')
$base = $base.Replace('"','')

#Cabeçalho do array das marés
$Header = "tide","height","date"
#cria array com dados das marés
$array = @($base)
$mares = $array | ConvertFrom-Csv -Header $Header

#Obtém data atual
$data = Get-Date -Format yyyy-MM-dd
$dia = Get-Date -Format dddd

#Obtém dada que o script irá concluir a análise
$datafim = (Get-Date).AddDays(6) 
$datafim = Get-Date $datafim -Format yyyy-MM-dd  

#Inicia varável $maresAltas
$maresAltas = ""

#Inicia automação com condicional se for terça-feira
if($dia -eq "segunda-feira"){
    #inicia loop
    foreach($mare in $mares){
        #cria array com data e horário das marés
        $d = $mare.date
        $dheader = "date","time"
        $darray = @($d.Replace(" ",","))
        $dataArray = $darray | ConvertFrom-Csv -Header $dheader 

        #Obtém o valor a data de dentro do loop
        $dataLoop = $dataArray.date

        #Cria gatilho condional para iniciar script
        if($data -eq $dataLoop){
            $x =1
        }

        #Valida o inicio do gatilho
        if($x -eq 1){
            #trata altura para somente duas casas decimais 
            $altura = $mare.height
            $altura = [math]::Round($altura,2)   
            #valida se a maré será alta e se é maior que 0.90          
            if($mare.tide -eq "a" -and $altura -ge 0.90){
                #registra valores na variável $maresAltas
                $maresAltas += $mare.tide + "," + $altura + "," + $mare.date + ";"
            }
        }

        #Encerra loop
        if($datafim -eq $dataLoop){
            break
        }
    }
}

#Converte em matriz de objetos
$maresAltas = $maresAltas -split ";" | ConvertFrom-Csv -Header $header

#Corpo do e-mail
$body = '<html><head><title>Mares Altas em Paraty</title><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body><p style="text-align: center;"><strong>Teremos essa semana os seguintes dias com mare alta na cidade de Paraty</strong><center><table border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; "><tr><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">Dia</td><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">Horario</td><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">Altura</td></tr>'

#Inicia loop
foreach($mare in $maresAltas){
    #Cria array com data e horário das marés
    $d = $mare.date
    $dheader = "date","time"
    $darray = @($d.Replace(" ",","))
    $dataArray = $darray | ConvertFrom-Csv -Header $dheader 
    
    #Trata data
    $dataMare = $dataArray.date
    $dataMare = Get-Date $dataMare -Format dd/MM/yyyy

    #Atribui dados ao corpo do e-mail
    $body += '<tr><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">'+$dataMare+'</td><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">'+$dataArray.time+'</td><td style="color: #4A423C; font-family: arial, helvetica, sans-serif; font-size: 18px; padding: 5px 5px;" align="center">'+$mare.height+'</td></tr>'
}

#Finaliza corpo do e-mail
$body += "</table></center></p></body></html>" 
$body | Out-File index.html

#Define parâmetros para envio de e-mail
$smtpServer                 = 'smtp.outlook.com'
$port                       = '587' # or '25' if not using TLS
$from                       = 'EMAIL_ORIGEM@hotmail.com'
$to                         = 'EMAIL_DESTINO@DOMÍNIO.com'
$subject                    = "Mares Altas da Semana"

#Pscredential
$passwd = ConvertTo-SecureString -String “SENHA_EMAIL_ORIGEM” -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList (“EMAIL_ORIGEM@hotmail.com”,$passwd)

#Envia e-mail de acordo com os parâmetros
Send-MailMessage -SmtpServer $smtpServer -Port $port -UseSsl -Credential $credential -From $from -to $to -Subject $subject -Body $body -BodyAsHtml