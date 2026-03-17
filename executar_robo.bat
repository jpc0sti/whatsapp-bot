@echo off
:: Define o título da janela do terminal
title Monitor de Automacao - WhatsApp

echo ======================================================
echo          INICIANDO ROBO DE WHATSAPP (OFFICE)
echo ======================================================
echo.

:: Verifica se a pasta node_modules existe (se o npm install foi feito)
if not exist "node_modules\" (
    echo [ERRO] Os motores do robo nao estao instalados.
    echo Rodando "npm install" para voce automaticamente...
    call npm install
)

echo [INFO] Iniciando o script JS...
echo [DICA] Mantenha esta janela aberta enquanto o robo trabalha.
echo.

:: Executa o seu script principal
node criar_do_excel.js

echo.
echo ======================================================
echo          PROCESSO FINALIZADO COM SUCESSO!
echo ======================================================
pause