dim a,b,c
dim audio, resp

call carregar_audio

sub carregar_audio()
    set audio = createobject("SAPI.SPVOICE")'Criando objeto de voz
    audio.volume = 100
    audio.rate = 2 'velocidade da fala

    call mostrar_num
end sub

sub mostrar_num()
    a=Cint(inputbox("Digite o primeiro numero ","AVISO"))
    b=Cint(inputbox("Digite o segundo numero ","AVISO"))
    c=Cint(inputbox("Digite o terceiro numero ","AVISO"))
    nome=(inputbox("Digite seu nome: ","AVISO"))

    audio.speak("ola, "&nome&"")

    if (a > b and b>c) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &c&""+vbnewline &_
        "Maior"&a&"") 
    end if

    if (a < b and b<c) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &a&""+vbnewline &_
        "Maior"&c&"") 
    end if

    if (a < b and b<c and c<a) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &c&""+vbnewline &_
        "Maior"&b&"") 
    end if
    if (a > b and b<c and c>a) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &b&""+vbnewline &_
        "Maior"&c&"") 
    end if
     if (a < b and b>c and c>a) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &a&""+vbnewline &_
        "Maior"&b&"") 
    end if
    if (a > b and b<c and c<a) then
       audio.speak("Vou falar qual e o maior e o menor numero"+ vbnewline &_
        "Menor: " &b&""+vbnewline &_
        "Maior"&a&"") 
    end if
end sub