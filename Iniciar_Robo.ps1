Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "      INICIANDO AUTOMAÇÃO ESALVADOR       " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Aguarde enquanto o sistema prepara o robô..." -ForegroundColor Yellow

# Roda o script principal do Python
py main.py

Write-Host ""
Write-Host "Processo concluído. Pressione qualquer tecla para fechar a janela..." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
