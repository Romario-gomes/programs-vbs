dim palavras(20),sorteio,nome, resp,pontos,nivel,palavra
dim audio



call carregar_audio
    sub carregar_audio()
    set audio=createobject("SAPI.SPVOICE") 'Criando objeto de voz
    audio.volume=100
    audio.rate = 1 'Velocidade da fala
    call iniciar_jogo
end sub

sub iniciar_jogo()
    palavras(1)= "Manga"
    palavras(2)= "Uva"
    palavras(3)= "Laranja"
    palavras(4)= "Limão"
    palavras(5)= "Abacaxi"

    nome = inputbox("Qual seu nome")

    msgbox("Bem vindo ao jogo SOLETRANDO: Nivel 1")
	randomize(second(time))
    sorteio = int(rnd*4)+1
    
    audio.speak(palavras(sorteio))

    do 
    resp = msgbox("Repetir?",vbYesNo,"ATENÇÃO")
        audio.speak(palavras(sorteio))

        palavra=Cstr(inputbox("Digite o a palavra: ","AVISO"))

        if palavra = palavras(sorteio) then
            msgbox("Você acertou!") 
        else
            msgbox("Você errou!")
        end if    
    loop while resp = vbYes
    
end sub











