dim numA, numB, total, resp,operacoes, resposta, opera, cont
cont = cdbl(0)
call sortear_num

sub sortear_num()
    randomize(second(time))
    numA=int(rnd * 10) + 1
    numB=int(rnd * 10) + 1
    operacoes=int(rnd * 4) + 1

    select case operacoes
        case 1:
            total = numA + numB
            opera = "+"
        case 2:
            total = numA - numB
            opera = "-"
        case 3:
            total = numA * numB
            opera = "*"
        case 4:
        total = numA / numB
            opera = "/"
    end select
    resposta=cdbl(inputbox("Qual o valor de "&numA&" "&opera&""&numB))
    

    if(resposta = total) then
        cont = cont+1
        msgbox("Resposta Correta!!"+vbnewline & _
        "Quantidade de acertos:"&cont)

        call sortear_num
        else
            msgbox("Você perdeu."+vbnewline & _
            "Quantidade de acertos:"&cont& +vbnewline)

            resp=msgbox("Deseja jogar novamente?",vbquestion + vbyesno, "ATENÇÃO")
                
            if resp=vbyes then
                cont=0
                call sortear_num()
                else
                    wscript.quit
                
            end if
    end if
end sub

    
