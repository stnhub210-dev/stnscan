@echo off 
chcp 65001 > nul 
cd /d C:\Users\User\Desktop\½ºÄµ 
git status --porcelain > C:\Users\User\AppData\Local\Temp\gitstatus.txt 
for %%%%A in (C:\Users\User\AppData\Local\Temp\gitstatus.txt) do if %%%%~zA==0 goto :nochange 
git add -A 
git commit -m "update" 
git push 
:nochange
