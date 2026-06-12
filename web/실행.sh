#!/bin/bash
echo "창고이동 검수 웹 프로그램 시작 중..."
pip install flask xlrd openpyxl -q
echo ""
echo "========================================"
echo " 브라우저에서 접속하세요: http://localhost:5000"
echo "========================================"
python3 -c "import webbrowser,threading,time; threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000')).start()"
python3 app.py
