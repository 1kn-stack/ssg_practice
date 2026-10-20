# 모먼트룸 매출 비교 — 1단계

교육용 가상 데이터 480행을 제공하는 로컬 정적 앱입니다. 기존 운영대시보드 예제는 보존합니다.

저장소에는 소스코드와 실행 설정만 포함합니다. 원본 A~E 자료, 정규화 JSON, 보고서와 검사 기록은 포함하지 않습니다. 새로 복제한 환경에서는 원본 실습 자료를 로컬에 별도로 준비한 뒤 아래 생성 명령을 실행해야 합니다.

프로젝트 루트에서 실행:

```powershell
python scripts/prepare_sales.py
python -m http.server 8000 --bind 127.0.0.1 --directory results/sales-app
```

http://localhost:8000 에 접속합니다. Python 표준 라이브러리와 동봉한 JSZip을 사용하므로 별도 패키지 설치나 외부 CDN이 필요하지 않습니다. 서버는 앱 디렉터리만 제공합니다.

점포·채널·기준주 필터가 카드·차트·표·보고서에 함께 적용됩니다. 추가 Excel 파일은 열 이름으로 정규화하며 원본 파일/시트/행을 보관합니다. 중복 record_id는 기존 기록을 유지하고 추가 합산하지 않습니다. 필수 열 누락 파일은 거부하며 정상 데이터는 유지합니다.

메모는 입력 즉시 record_id별로 현재 브라우저의 로컬 저장소에 보관됩니다. 다른 사용자와 공유되지 않습니다. 브라우저 데이터 삭제 시 사라질 수 있습니다.

보고서는 04_보고양식.xlsx에서 읽은 16개 헤더 순서로 새 XLSX를 다운로드합니다. 브라우저 다운로드 설정의 위치에 저장되며 원본 양식은 수정하지 않습니다. 목표 0의 달성률은 화면에서 —, Excel에서는 빈 셀입니다.

원본 재대조:

```powershell
python scripts/verify_sales.py
```

검사 기록은 results/verification.md, 원본 해시는 results/source-hashes-before.json에 있습니다. 원본 A~E 폴더는 Git 제외 대상입니다. 이 폴더의 data/sales.json은 제공 교육용 자료에서 생성한 별도 사본입니다. 실제 영업 자료로 교체한 결과는 배포하거나 커밋하지 마세요.

GitHub 대상: https://github.com/1kn-stack/ssg_practice (커밋·푸시는 사용자가 요청할 때 수행).
