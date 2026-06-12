#!/bin/bash
echo "창고이동 3단계 검수 프로그램 시작 중..."
pip install streamlit xlrd openpyxl -q
echo ""
echo "브라우저에서 자동으로 열립니다: http://localhost:8501"
echo ""
streamlit run app.py --server.headless false --browser.gatherUsageStats false
