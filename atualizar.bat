@echo off
echo A atualizar ficheiros no GitHub...
git add .
git commit -m "Atualizacao automatica diaria de faturas e horas"
git push origin master
echo Atualizacao concluida com sucesso!
pause