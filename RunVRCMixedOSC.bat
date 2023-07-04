@echo off
call UpdateScripts.bat
echo [%~n0] Installing requirements (be sure to have python installed and in PATH)
python -m pip install -r VRCMixedOSC/requirements.txt
python VRCMixedOSC/vrcmixedosc.py
pause