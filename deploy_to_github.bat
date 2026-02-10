@echo off
echo Initializing Git repository...
git init
git add .
git commit -m "Deploy: Web Surat Fixes"
git branch -M main
echo Adding remote origin...
git remote add origin https://github.com/fidyan1/fidyan.github.io.git
echo Pushing to GitHub...
git push -u origin main --force
pause
