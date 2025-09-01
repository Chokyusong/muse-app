파일 내보내기
1) 준비


<img width="791" height="501" alt="image" src="https://github.com/user-attachments/assets/9248ea40-15ed-4592-9d93-20be5f645ee4" />

(선택) 가상환경

py -m venv .venv
.venv\Scripts\activate


PyInstaller 설치

pip install pyinstaller

2) 기본 빌드 (콘솔 숨김 + 단일 파일)

desktop_app.py가 진입점이라고 가정합니다.

pyinstaller --onefile --noconsole --name MuseApp desktop_app.py


--onefile: 단일 exe

--noconsole: 콘솔창 숨김(팝업만 보임)

생성물: dist\MuseApp.exe

테스트:
dist\MuseApp.exe --expire-now
→ 만료 팝업이 뜨고 바로 종료되면 성공.

3) 아이콘/버전 정보(선택)

아이콘:

pyinstaller --onefile --noconsole --icon app.ico --name MuseApp desktop_app.py


버전 리소스 파일(선택, file_version.txt 등) 만들어 넣을 수도 있습니다. 필요하면 템플릿 드릴게요.

4) 외부 파일 동봉이 필요할 때

메시지 템플릿, 이미지, .env 같은 파일이 exe에 같이 들어가야 한다면 --add-data:

pyinstaller --onefile --noconsole --name MuseApp ^
  --add-data "message.txt;." ^
  --add-data ".env;." ^
  desktop_app.py


구분자: Windows는 ;, mac/Linux는 :

실행 시 파일 경로 얻는 방법(동봉/개발 환경 모두 대응):

import sys, os

def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)

msg_file = resource_path("message.txt")
with open(msg_file, "r", encoding="utf-8") as f:
    msg = f.read()

5) Selenium/웹드라이버 사용 시 팁

webdriver_manager를 쓰면 exe에 드라이버를 굳이 포함할 필요가 없습니다(최초 실행 시 자동 다운로드).

로컬에 직접 드라이버를 포함하려면 --add-binary로 추가하고 코드에서 해당 경로를 지정하세요.

6) 흔한 이슈 정리

팝업만 보이게 하고 싶다 → 이미 --noconsole로 해결.

무헤드/서버에서 tkinter가 안 뜨는 환경 → 제시해드린 코드가 자동으로 콘솔 메시지로 폴백 후 종료합니다.

안티바이러스/스마트스크린 경고 → 가능한 한

빌드 후 압축 해제형(--onedir)로 테스트해보고,

최종 배포는 코드서명(Authenticode)까지 하면 경고가 줄어듭니다.

상대경로 에러 → 위 resource_path 패턴 사용.

엔트리 가드 권장

if __name__ == "__main__":
    # 현재 메인 로직

7) 최종 체크리스트

만료 로직 블록이 파일 최상단에 있고,

FORCE_EXPIRE=False(배포 모드),

배포일/기간 값이 정확,

pyinstaller --onefile --noconsole로 빌드,

dist\MuseApp.exe --expire-now로 최종 테스트 OK.

필요하시면 PyInstaller .spec 파일까지 만들어서 add-data, 아이콘, 버전정보를 한 번에 고정하는 템플릿도 바로 드릴게요.

1) requirements.txt 준비

지금 만든 파일을 프로젝트 루트에 두세요.

예시:
<img width="217" height="158" alt="image" src="https://github.com/user-attachments/assets/916f93e2-f738-4f8d-a5af-1b35a2b7bbed" />


2) 가상환경(권장)
python -m venv .venv
.venv\Scripts\activate   # Windows
source .venv/bin/activate  # macOS/Linux

3) 한 번에 설치
pip install -r requirements.txt


→ 목록에 있는 모든 패키지가 자동 설치됩니다.

4) 개발/빌드 도구까지 같이 설치하고 싶을 때
pip install -r dev-requirements.txt

⚡ 버전까지 고정해서 재현하고 싶다면

지금 환경에서:

pip freeze > requirements.lock.txt


→ 나중에 설치할 땐:

pip install -r requirements.lock.txt


이러면 현재와 완전히 동일한 버전으로 세팅됩니다.

👉 정리하면:
한 줄 명령어

pip install -r requirements.txt


이게 “한 번에 설치”하는 공식 방법이에요 ✅

원하시나요? 제가 지금 코드 기준으로 requirements.txt 완성본을 직접 작성해드릴까요 (버전 고정 없이 최신 설치용)?
