# MIGRATION — read-book (하루 한 권 책 정리)

## 실행 방법
자동화 루틴이 `CLAUDE.md` 지침대로 하루 한 권을 정리해 노션에 저장하고 카카오톡으로 알림. 진행 상황의 단일 진실원은 노션 (`reading-list.md` 는 보조 기록).
`책_요약/` 의 `build.py`, `build2.py` 는 책장 도서목록 xlsx 를 만드는 스크립트.

## 버전
Python 3.13.x, `openpyxl` (build 스크립트용, 필요 시 `pip install openpyxl`)

## 환경변수
없음. 인증은 Claude 앱의 Notion 커넥터 / 카카오 연동으로 처리.

## 외부 도구
Claude 앱 + Notion 커넥터, 카카오톡 알림 연동

## 스케줄러
이 PC 의 Windows 작업 스케줄러·Claude 로컬 예약 작업에는 등록돼 있지 않음 (원격/클라우드 루틴으로 추정 — 새 PC 에서 루틴 설정 확인 필요).

## git 에 안 올라가는 것
없음. 이 repo 는 **public** 이므로 키·토큰·개인정보를 절대 커밋하지 말 것.