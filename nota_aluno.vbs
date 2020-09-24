dim n1,n2,n3,media,situacao 'Declarando variáveis locais alfanumericas
dim audio,resp

call carregar_audio
sub carregar_audio()
set audio=createobject("SAPI.SPVOICE") 'Criando objeto de voz
audio.volume=100
audio.rate = 2 'Velocidade da fala
call calcular_rendimento
end sub

sub calcular_rendimento()
n1=cdbl(inputbox("Digite a nota 1","AVISO"))
n2=cdbl(inputbox("Digite a nota 2","AVISO"))
n3=cdbl(inputbox("Digite a nota 3","AVISO"))
media=round((n1+n2+n3)/3,1)
if media < 4 then
situacao="Reprovado"
elseif media >=7 then
situacao="Aprovado"
else
situacao="Exame"
end if

'Saida de Dados por voz
audio.speak ("Rendimento do Aluno" + vbnewline &_
"Média do Aluno "& media &"" + vbnewline & _
"Situação do Aluno "& situacao &"")
'Saida de Dados por msg
msgbox("===========================" + vbnewline &_
" RENDIMENTO DO ALUNO " + vbnewline & _
"===========================" + vbnewline &_
"Média do Aluno : "& media &"" + vbnewline &_
"Situação do Aluno: "& situacao &""),vbinformation + vbokonly,"AVISO"
resp=msgbox("Deseja realmente encerrar?",vbquestion + vbyesno,"ATENÇÃO")
if resp=vbyes then
wscript.quit
else
call calcular_rendimento
end if
end sub