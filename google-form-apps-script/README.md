# Google Form + Apps Script Setup

이 폴더에는 `주력 강의 재설계 과정` 신청서를 생성하는 Apps Script 코드가 들어 있습니다.

생성 범위:

- Google Form 생성
- 응답용 Google Sheet 생성 및 연결
- 제출 시 관리자 알림 메일 발송
- 제출 시 신청자 확인 메일 발송

## 준비

공식 문서 기준으로 아래 두 가지가 필요합니다.

- `clasp` 설치
- Google Apps Script API 활성화

참고:

- [clasp 공식 문서](https://developers.google.com/apps-script/guides/clasp)
- [Forms Service 공식 문서](https://developers.google.com/apps-script/reference/forms)

## 실행 순서

1. `clasp` 설치

```bash
npm install -g @google/clasp
```

2. Apps Script API 활성화

- `https://script.google.com/home/usersettings`
- `Google Apps Script API`를 켭니다.

3. Google 로그인

```bash
clasp login
```

4. 현재 폴더에서 Apps Script 프로젝트 생성

```bash
cd C:\코딩\공개과정\google-form-apps-script
clasp create --type standalone --title "주력 강의 재설계 과정 신청서 자동화" --rootDir .
```

5. 로컬 코드 업로드

```bash
clasp push -f
```

6. Apps Script 편집기에서 API 실행 배포 생성

- `clasp open`
- 우측 상단 `배포` > `새 배포`
- `유형 선택`에서 `API Executable` 추가
- 접근 권한은 `나만`으로 설정
- 배포를 생성합니다.

7. Apps Script 편집기 열기

```bash
clasp open
```

8. 편집기 또는 `clasp run`으로 `createCourseForm` 함수 1회 실행

실행 후 생성되는 항목:

- Google Form 1개
- 응답 Spreadsheet 1개
- 제출 트리거 1개

## 생성되는 문항

- 신청자 기본 정보
- 주력 강의 정보
- 우선 개선 과제
- 제출 가능 자료
- 확인 사항

## 알림 메일

기본 관리자 메일:

- `JongmokJ@gmail.com`

필요하면 `Code.js`의 `CONFIG.adminEmail` 값을 바꾸면 됩니다.
