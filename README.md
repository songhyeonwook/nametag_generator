# Nametag Generator (명찰 생성기)

Excel 또는 CSV 명단을 업로드하여 여러 명의 명찰이 포함된 PowerPoint 파일을 생성하는 앱입니다.

이 저장소에는 두 가지 버전이 있습니다.

- `index.html`, `style.css`, `script.js`: GitHub Pages에서 바로 사용할 수 있는 정적 웹앱
- `app.py`: 기존 Streamlit/Python 버전

## GitHub Pages 버전

GitHub Pages 버전은 서버 없이 브라우저 안에서만 동작합니다.

### 주요 기능

- Excel/CSV 파일 업로드
- PPTX 템플릿 업로드
- 업로드한 PPTX의 첫 번째 슬라이드 전체를 A4 용지 안에 축소 복제
- `{{company}}`, `{{department}}`, `{{name}}`, `{{dept_name}}` 플레이스홀더 치환
- `회사명`, `부서`, `이름` 컬럼 자동 매핑
- A4 가로/세로 방향 선택
- 1x1, 2x2, 3x3 등 명찰 배열 선택
- 인당 출력 개수 설정
- 브라우저에서 PPTX 파일 생성 및 다운로드

### 사용 방법

1. `index.html`을 브라우저에서 엽니다.
2. Excel 또는 CSV 명단을 업로드합니다.
3. PPTX 템플릿을 업로드합니다.
4. 컬럼 매핑과 출력 옵션을 확인합니다.
5. `명찰 PPTX 생성` 버튼을 눌러 파일을 다운로드합니다.

### 템플릿 작성 방법

PowerPoint에서 명찰 디자인을 만들고 첫 번째 슬라이드에 아래 플레이스홀더를 넣어주세요.

- `{{company}}`: 회사명
- `{{department}}`: 부서
- `{{name}}`: 이름
- `{{dept_name}}`: 부서와 이름을 합친 텍스트

생성 결과는 선택한 배열에 맞춰 A4 한 페이지에 여러 명찰을 배치한 PPTX입니다. 템플릿 첫 슬라이드의 도형, 텍스트, 이미지, 스타일을 복사하고 각 칸에 맞게 좌표와 크기를 축소합니다.

### GitHub Pages 배포 방법

이 저장소에는 GitHub Actions 배포 워크플로가 포함되어 있습니다.

1. 이 폴더의 파일을 GitHub 저장소에 커밋하고 `main` 브랜치로 푸시합니다.
2. GitHub 저장소에서 `Settings` > `Pages`로 이동합니다.
3. `Build and deployment`의 `Source`를 `GitHub Actions`로 선택합니다.
4. `Actions` 탭에서 `Deploy GitHub Pages` 워크플로가 성공하면 아래 주소로 접속할 수 있습니다.

```text
https://songhyeonwook.github.io/nametag_generator/
```

## Streamlit/Python 버전

업로드한 PPTX 템플릿의 첫 번째 슬라이드를 복제하고 `{{company}}`, `{{department}}`, `{{name}}`, `{{dept_name}}` 플레이스홀더를 치환하려면 Python 버전을 사용하세요.

### 로컬에서 실행하기

파이썬 환경이 구축되어 있다면 아래 명령어로 직접 앱을 실행할 수 있습니다.

```bash
# 가상환경 생성 및 활성화
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt

# 앱 실행
streamlit run app.py
```

브라우저가 열리며 `http://localhost:8501`에서 명찰 생성기를 사용할 수 있습니다.
