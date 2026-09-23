# Battery Teardown Tracker (웹 공개 버전)

e81c-calculator-web과 동일한 방식입니다: **추가 설정이나 로그인 없이, 링크만 열면
누구나 바로 쓸 수 있는 정적 웹사이트**이며 GitHub Pages로 호스팅됩니다.

## 데이터는 어디에 저장되나요? (중요)

이 버전은 별도의 서버나 클라우드 데이터베이스를 쓰지 않습니다. 대신 **입력한 데이터와
첨부 이미지는 그 사람의 브라우저 안(IndexedDB)에만 저장**됩니다.

- 같은 사람이 같은 브라우저로 다시 접속하면 새로고침해도, 컴퓨터를 재부팅해도 데이터가
  그대로 남아 있습니다.
- **하지만 다른 사람이 이 링크를 열거나, 같은 사람이 다른 브라우저/다른 PC에서 열면
  빈 화면부터 시작합니다.** 데이터가 사람들 사이에 공유되지 않습니다 (계산기와 똑같은
  "정적 사이트" 방식의 자연스러운 한계입니다).
- 브라우저의 "인터넷 사용 기록/사이트 데이터 삭제"를 하면 이 사이트에 저장된 데이터도
  함께 지워집니다. 시크릿(비공개) 창에서는 창을 닫는 순간 데이터가 사라집니다.

**여러 사람·여러 PC가 같은 데이터를 실시간으로 함께 봐야 한다면** 이 방식으로는 안
되고, 별도의 공유 저장소(Firebase 등)를 연결한 버전이 필요합니다 — 필요하시면
말씀해주세요, 이전에 준비해드렸던 Firebase 연동 버전을 다시 꺼내드릴 수 있습니다.

### 그럼 각자 입력한 게 다 흩어지는 건데, 어떻게 모으나요?

화면 상단의 **"CSV 내보내기" / "Excel 내보내기"** 버튼으로 각자 작업한 내용을
파일로 저장한 뒤, 이메일이나 사내 공유 폴더로 모아서 하나의 파일로 합치는 방식을
권장합니다. (지금은 자동 병합 기능은 없습니다.)

---

## 1. 배포 방법 (제가 지금 바로 진행 가능)

이 컴퓨터에 GitHub CLI(`gh`)가 `sangnew` 계정으로 로그인되어 있어 아래 과정을 제가
대신 실행해 드릴 수 있습니다. 다음 명령으로 저장소를 만들고 GitHub Pages를 켭니다:

```
cd web
git init
git add .
git commit -m "Battery Teardown Tracker web app"
gh repo create battery-teardown-tracker-web --public --source=. --remote=origin --push
echo '{"source":{"branch":"main","path":"/"}}' | gh api -X POST repos/sangnew/battery-teardown-tracker-web/pages --input -
```

배포 후 아래 주소에서 바로 접속됩니다 (몇 분 정도 걸릴 수 있습니다):

```
https://sangnew.github.io/battery-teardown-tracker-web/
```

## 2. 이후 업데이트 방법

`web/` 폴더 파일을 수정한 뒤:

```
git add .
git commit -m "설명"
git push
```

GitHub Pages가 자동으로 새 버전을 반영합니다 (보통 1분 이내).

## 3. 로컬에서 미리 테스트하고 싶다면 (선택)

```
npx serve web
```

`file://`로 index.html을 직접 열면 일부 브라우저에서 정상 동작하지 않을 수 있으니
간단한 로컬 서버를 통해 열어주세요.

---

## 주요 기능

- 신규 등록 / 표에서 직접 셀 편집 / 행 추가·삭제
- 엑셀에서 여러 행·열을 복사해 표에 붙여넣기 (Ctrl+V)
- CSV/Excel 가져오기·내보내기
- Cell ID 검색, LOT/Grade/Operator/TD Status/TD Date 필터
- Cell ID 중복 시 경고 + 기존 기록 바로가기
- VHX/Image 다중 첨부 + 썸네일 클릭 확대보기
- TD Result/Further Analysis/SEM·EDS 긴 텍스트도 표에서 바로 입력 (Enter 저장, Shift+Enter 줄바꿈)
- 탭 2개: **Tear Down** / **Frozen IR · Spot 분석** (LOT, Cell ID, Grade, dOCV, Frozen IR Result, Pass/Fail, voltage drop, Dropped Layer, dOCV (V), Spot Found, Top/Back, x, y, Shape, SEM/EDS, Location, Long/Short side, Height)
- 가져오기·내보내기는 현재 보고 있는 탭 기준
- 빈 값은 0으로 바뀌지 않고 그대로 보존, OCV 소수점/IR "O.F." 등 원본 문자열 그대로 저장
- 연도 없는 TD Date(예: "1/26")는 임의로 채우지 않고 저장 시 사용자에게 연도를 확인
