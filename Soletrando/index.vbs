dim palavras(20),sorteio,nome, resp,pontos,nivel,palavra
dim audio

pontos = 0

call carregar_audio
    sub carregar_audio()
    set audio=createobject("SAPI.SPVOICE") 'Criando objeto de voz
    audio.volume=100
    audio.rate = 1 'Velocidade da fala
    call pegar_nome
end sub

sub pegar_nome()
    nome = inputbox("Qual seu nome")
    call iniciar_jogo
end sub

sub nivel_um()
    palavras(1)= "Manga"
    palavras(2)= "Uva"
    palavras(3)= "Laranja"
    palavras(4)= "Limao"
    palavras(5)= "Abacaxi"
end sub

sub nivel_dois()
    palavras(1)= "Corola"
    palavras(2)= "Cruze"
    palavras(3)= "Onix"
    palavras(4)= "Prisma"
    palavras(5)= "HB20"
end sub




sub iniciar_jogo()
    call nivel_um

    msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel 1")

	randomize(second(time))
    sorteio = int(rnd*4)+1
    
    audio.speak(palavras(sorteio))

    do 
    resp = msgbox("Repetir?",vbYesNo,"ATENCAO")
        audio.speak(palavras(sorteio))

        palavra=Cstr(inputbox("Digite o a palavra: ","AVISO"))

        if palavra = palavras(sorteio) then
            msgbox("Voce acertou!")
            pontos = pontos + 1
            if pontos >= 5 then 
                call nivel_dois
                msgbox("Voce passou de nivel com " &pontos& "acertos")
            end if



            call iniciar_jogo
        else
            msgbox("Voce errou!")
        end if    
    loop while resp = vbYes
    
end sub











