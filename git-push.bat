@echo off
echo 正在推送代码到GitHub...
"C:\Program Files\Git\bin\git.exe" add .
"C:\Program Files\Git\bin\git.exe" commit -m "Update: %date% %time%"
"C:\Program Files\Git\bin\git.exe" push origin main
echo 推送完成！
pause 