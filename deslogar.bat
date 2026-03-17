@echo off
echo Removendo login atual do WhatsApp...
rd /s /q .wwebjs_auth
echo Pronto! Na proxima vez que rodar o script, ele pedira o QR Code.
pause