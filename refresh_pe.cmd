@echo off
rem Refresh do P/E (Bloomberg DAPI) + push — roda só com o Terminal logado.
rem Agendado em dias úteis via Task Scheduler; também funciona por duplo-clique.
cd /d "%~dp0"
echo ===== %date% %time% ===== >> refresh_pe.log
python refresh_pe.py >> refresh_pe.log 2>&1
if errorlevel 1 (
  echo Terminal indisponivel ou erro — nada publicado. >> refresh_pe.log
  exit /b 1
)
git add p_e.xlsx >> refresh_pe.log 2>&1
git diff --cached --quiet && (
  echo Sem mudancas no p_e.xlsx — nada a publicar. >> refresh_pe.log
  exit /b 0
)
git commit -m "data: P/E refresh (Bloomberg DAPI)" >> refresh_pe.log 2>&1
git pull --rebase origin main >> refresh_pe.log 2>&1
git push origin main >> refresh_pe.log 2>&1
echo Publicado com sucesso. >> refresh_pe.log
