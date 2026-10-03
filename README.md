# 팀플 메이트

## 처음 쓰는 순서
1. 앱을 켜면 시작 페이지가 나와요. 팀원이 모두 모인 자리에서 주제, 팀원 이름과 성향을 넣고, 각자 자기 비밀번호를 입력해 팀을 만들어요.
2. 다음부터는 시작 페이지에서 '나는 누구인가요?'에서 이름을 고르고 비밀번호를 넣으면 들어갈 수 있어요.
3. 자료 올리기, 투표, 할 일 체크는 로그인한 본인 것만 할 수 있어요.

## 저장
- 모든 내용은 teamplay_data.json 파일에 자동으로 저장돼요. 이 파일을 지우면 처음 상태가 돼요.
- 사이드바 [설정] → '백업 파일 받기'로 파일을 받아 두면, 다른 컴퓨터에서 '백업 불러오기'로 이어서 쓸 수 있어요.
- 모두가 비밀번호를 잊었다면 앱을 끄고 teamplay_data.json을 지운 뒤 다시 켜면 돼요. 이때 모든 기록이 사라져요.

## 실행
    py -m pip install -r requirements.txt
    py -m streamlit run app.py

## AI 연결 - 무료: Gemini API 키
1. aistudio.google.com 에 구글 계정으로 로그인합니다.
2. 왼쪽 메뉴의 Get API key → Create API key 를 누릅니다.
3. AIza 로 시작하는 키를 복사합니다.
4. 앱 사이드바 [설정] → AI 종류 'Gemini (무료)' → Gemini API 키 칸에 붙여넣기 → [연결 테스트]

## AI 연결 - 유료: Claude API 키
1. platform.claude.com 에 가입하고 로그인합니다.
2. Billing(결제)에서 크레딧을 충전합니다. 5달러 정도면 연습용으로 충분합니다.
3. Settings → API keys → Create key 로 키를 만듭니다. sk-ant- 로 시작하며, 처음 한 번만 보이니 바로 복사해 두세요.
4. 둘 중 하나로 넣습니다.
   - 잠깐 쓸 때: 앱 사이드바 맨 아래 [설정] → Claude API 키 칸에 붙여넣기 (창을 닫으면 사라짐)
   - 계속 쓸 때: .streamlit/secrets.toml.example 파일 이름을 secrets.toml 로 바꾸고 키를 붙여넣기

키가 없어도 앱은 규칙 기반으로 모두 동작합니다.
키는 비밀번호와 같습니다. 남에게 보내거나 인터넷에 올리지 마세요.
