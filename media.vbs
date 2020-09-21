dim n1, n2, n3, media, situacao 'Declaracao de variaveis
dim audio, resp
'Cint - Converte String para numero inteiro
'CLng - Converte string numeros longos
'Cdbl - Converte string numeros decimais
'Ccur - Converte string valores 

call carregar_audio 'Chamando subrotina 

sub carregar_audio()
    set audio = createobject("SAPI.SPVOICE")'Criando objeto de voz
    audio.volume = 100
    audio.rate = 2 'velocidade da fala

    call calcular_rendimento
end sub

sub calcular_rendimento()
    n1=Cdbl(inputbox("Digite a nota 1","AVISO"))
    n2=Cdbl(inputbox("Digite a nota 2","AVISO"))
    n3=Cdbl(inputbox("Digite a nota 3","AVISO"))
    media=round((n1+n2+n3)/3,1)'round nao deixa ter dizima
    
    if media < 4 then
        situacao = "Reprovado"
    elseif media >= 7 then
        situacao = "Arovado"
    else
        situacao = "exame"
    end if

'Saida de dados por voz
audio.speak("Rendimento do Aluno"+ vbnewline &_
    "Media do aluno" &media &""+vbnewline &_
    "Situaçao do Aluno"&situacao&"")

'Saida de dados por msg

msgbox("======================================" +vbnewline &_
    "         RENDIMENTO DO ALUNO         "+vbnewline &_
    "Media do Aluno: "&media&""+vbnewline &_
    "Situaçao do Aluno: "&situacao&""),vbinformation + vbokonly,"AVISO"

    resp = msgbox("Deseja realmnte encerrar?",vbquestion + vbyesno,"ATENÇAO")
    if resp=vbyes then
        wscript.quit
        else
        call calcular_rendimento
    end if
end sub