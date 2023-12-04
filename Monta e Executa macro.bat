@echo off
cls
color 80
set mydate=%date%
set mytime=%time%
echo Data e Hora do inicio do processamento: %mydate% as %mytime%

echo Computador %computername%       Usuario:%username%  

:tarefas
echo                                                                                                                                                       .  
echo                   TAREFAS
echo    ------------------------------------------------
echo    1. Executar atualizacao por Macro 
echo    2. Criar executavel
echo    3. Sair
echo    ------------------------------------------------
echo                                                                                                                                                      . 

set /p opcao= Escolha uma opcao:
if %opcao% equ 1 goto opcao1
if %opcao% equ 2 goto opcao2
if %opcao% equ 3 goto opcao3
if %opcao% GEQ 4 goto opcao4

:opcao1
cls
echo
echo    ---------------------------------------------------------------------------------------
echo                                              ATENCAO
echo      Ao final das respostas do aplicativo, MINIMIZAR a tela de informacao mas nao FECHAR
echo    ---------------------------------------------------------------------------------------
echo
"C:\Users\acmdo\Desktop\dist\Despertar_Macro_V4.exe" "C:\Users\acmdo\Meu Drive\Python\PROGRAMAS\Guarda.xlsm" "MacroAcumula" "C:\Users\acmdo\Desktop\dist\Log despertar.txt"
cls
exit

:opcao2
cls
pyinstaller --onefile "C:\Users\acmdo\Meu Drive\Python\PROGRAMAS\Despertar_Macro_V4.py"
pause
cls
goto tarefas

:opcao3
cls
exit

:opcao4
echo    -----------------------------------------------------------------
echo    Tarefa inválida! Escolha uma opção das tarefas 
echo    ------------------------------------------------------------------
pause
cls
goto tarefas

