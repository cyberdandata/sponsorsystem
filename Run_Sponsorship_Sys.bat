node -v 
cd /d D:\/com
start /b node server
timeout /t 5
echo Opening App in the Browser in About 15seconds please wait if it opens but not load press F5
start http://localhost:3000/


