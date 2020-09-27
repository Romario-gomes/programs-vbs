dim palavras(20), palavra_sorteio,sorteio,nome, resp,pontos,nivel,palavra
dim audio, resp_sorteio,cont
dim nvl1(5), nvl2(5), nvl3(5), nvl4(5)

cont = 0

call carregar_audio'funcao de audio
    sub carregar_audio()
    set audio=createobject("SAPI.SPVOICE") 'Criando objeto de voz
    audio.volume=100
    audio.rate = 1 'Velocidade da fala
    call arrays
end sub

sub arrays()
    nvl1(1)= "manga"
    nvl1(2)= "uva"
    nvl1(3)= "laranja"
    nvl1(4)= "limao"
    nvl1(5)= "abacaxi"

    nvl2(1)= "corola"
    nvl2(2)= "cruze"
    nvl2(3)= "onix"
    nvl2(4)= "prisma"
    nvl2(5)= "hb20"

    nvl3(1)= "porta"
    nvl3(2)= "livro"
    nvl3(3)= "corrente"
    nvl3(4)= "abajur"
    nvl3(5)= "petiz"

    nvl4(1)= "permuta"
    nvl4(2)= "ruar"
    nvl4(3)= "veneta"
    nvl4(4)= "alcunha"
    nvl4(5)= "justapor"
    call pegar_nome
end sub

sub pegar_nome()
    nome = inputbox("Qual seu nome")'Pegando nome do usuario
    call iniciar_jogo
end sub

sub sortear_palavra()
    do
      resp_sorteio = true
            if cont < 3 then 
                randomize(second(time))
                sorteio=int(rnd * 5) + 1
                palavra_sorteio = nvl1(sorteio)
                nvl1(sorteio) = ""
            elseif cont < 6 then 
                randomize(second(time))
                sorteio=int(rnd * 5) + 1
                palavra_sorteio = nvl2(sorteio)
                nvl2(sorteio) = ""
            elseif cont < 9 then 
                randomize(second(time))
                sorteio=int(rnd * 5) + 1
                palavra_sorteio = nvl3(sorteio)
                nvl3(sorteio) = ""
            elseif cont < 12 then 
                randomize(second(time))
                sorteio=int(rnd * 5) + 1
                palavra_sorteio = nvl4(sorteio)
                nvl4(sorteio) = ""
            end if
        if palavra_sorteio = "" then
            resp_sorteio = false
        end if  
    loop while resp_sorteio = false
end sub






sub iniciar_jogo()
    call sortear_palavra
    if cont < 3 then
        msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel 1")
    elseif cont < 6 then
        msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel 2")
    elseif cont < 9 then
        msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel 3")
    elseif cont < 12 then
        msgbox("Bem vindo "& nome &" ao jogo SOLETRANDO: Nivel ")
    end if 

    do 
    
        audio.speak(palavra_sorteio)
        resp = msgbox("Repetir?",vbYesNo,"ATENCAO")
        
        palavra=Cstr(inputbox("Digite o a palavra: ","AVISO"))

        if palavra = palavra_sorteio then
            msgbox("Voce acertou!")
            cont = cont + 1
                call iniciar_jogo
            if cont > 3 then 
                msgbox("Voce passou de nivel com " &cont& "acertos")
                call iniciar_jogo
            end if
        else
            msgbox("Voce errou!")
        end if    
    loop while resp = vbYes
    
end sub











