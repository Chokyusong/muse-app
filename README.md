1) requirements.txt 준비

지금 만든 파일을 프로젝트 루트에 두세요.

예시:

requirements.txt
desktop_app.py
heart_aggregate.py
panda_dm_sender.py

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
