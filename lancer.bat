@echo off
title Plan Financier Hotel 5*
cd /d C:\Users\Iziboard2\hotel_plan_financier
echo ============================================
echo   Plan Financier Hotel 5* - Demarrage...
echo ============================================
echo.
python -m streamlit run app.py --server.port 8501 --server.headless true
pause
