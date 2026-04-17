@echo off
color 0A
echo ========================================================
echo   Enviando atualizacoes para o Github e Vercel...
echo ========================================================
git add .
git commit -m "Atualizacao rapida via script"
git push origin main
echo ========================================================
echo   Sucesso! As alteracoes ja estao subindo para a Vercel.
echo   Pressione qualquer tecla para fechar...
echo ========================================================
pause >nul
