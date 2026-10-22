@echo off
setlocal
cd /d "%~dp0"
set "JURIDICO_PACKAGE_DIR=%~dp0"
echo Preparando os arquivos do controle...
powershell.exe -NoLogo -NoProfile -NonInteractive -Command "$ErrorActionPreference='Stop'; $names=@('Assistente-Juridico.ps1','Juridico.Core.psm1','Juridico.Cadastro.ps1','Juridico.Interface.ps1'); foreach($name in $names){$path=Join-Path $env:JURIDICO_PACKAGE_DIR $name; if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw ('Falta o arquivo '+$name+'. Copie todos os arquivos da atualizacao para a mesma pasta.')}; Unblock-File -LiteralPath $path}"
if errorlevel 1 goto prepare_error
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%~dp0Assistente-Juridico.ps1"
if errorlevel 1 goto open_error
exit /b 0
:prepare_error
echo.
echo Nao foi possivel preparar os arquivos. Confira a mensagem acima.
echo Copie todos os arquivos do pacote para a mesma pasta do controle.
echo Se houver bloqueio da empresa, solicite a liberacao a TI.
pause
exit /b 1
:open_error
echo.
echo Nao foi possivel abrir o controle. Confira a mensagem acima.
echo Este controle usa Windows e Excel instalado no computador.
echo Se a mensagem exigir assinatura digital, pode haver uma regra da empresa.
echo A politica permanente do Windows nao foi alterada.
pause
exit /b 1
