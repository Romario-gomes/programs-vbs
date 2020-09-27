dim palavras(20), arrays(50),sorteio,nome, resp,pontos,nivel,palavra
dim audio

pontos = 0

call carregar_audio'funcao de audio
    sub carregar_audio()
    set audio=createobject("SAPI.SPVOICE") 'Criando objeto de voz
    audio.volume=100
    audio.rate = 1 'Velocidade da fala
    call pegar_nome
end sub

sub pegar_nome()
    nome = inputbox("Qual seu nome")'Pegando nome do usuario
    call iniciar_jogo
end sub


sub sortear_palavra()
    randomize(second(time))
    sorteio = int(rnd*5)+1
end sub

sub verificar_palavra()
    'Do While 
        for n=1 to 5 step 1 
			if palavras(sorteio) <> arrays(n) AND arrays(n) = empty then
                arrays(n) = palavras(sorteio)
            else 
                call sortear_palavra
            end if   
            msgbox(arrays(n))
	    next
    'Loop
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
    call sortear_palavra
    call verificar_palavra

    msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel 1")



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











