@echo off
cd /d "D:\AI AGENT THUY\AI branding"
python -u "scripts\fb_personal_interactor.py" --count 20 --headless >> "scratch\fb_personal_interactor.log" 2>&1
