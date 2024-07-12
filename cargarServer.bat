@echo off
cd C:\Users\cmunoz\Desktop\Web\GeneradorPPT

rem Iniciar el servidor Flask en segundo plano
start "" /b cmd /c "python TM.py"

rem Esperar unos segundos para asegurar que el servidor Flask se inicie
timeout /t 1 /nobreak >nul

rem Abrir la dirección en el navegador predeterminado
start http://127.0.0.1:5000

rem Finalizar el script sin pausar
exit