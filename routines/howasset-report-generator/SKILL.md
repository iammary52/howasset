---
name: howasset-report-generator
description: 하우자산운용 최소영업자본액 검토보고서 자동 생성 및 이메일 발송
---

## 최소영업자본액 검토보고서 자동 생성

Gmail에서 최신 .xlsx 첨부파일을 받아 최소영업자본액 검토보고서(Word + PDF)를 생성하고 이메일로 발송하는 작업입니다.

### 실행 단계

**1. 패키지 설치**
```bash
pip install openpyxl python-docx -q
```

**2. GitHub에서 스크립트 및 Word 템플릿 다운로드**
```bash
git clone https://{{GITHUB_PAT}}@github.com/iammary52/howasset.git /tmp/report
```

**3. Gmail 설정 파일 생성**

`/tmp/report/config.json` 파일을 아래 내용으로 작성:
```json
{
    "gmail_address": "{{GMAIL_ADDRESS}}",
    "gmail_app_password": "{{GMAIL_APP_PASSWORD}}"
}
```

**4. output 폴더 생성 후 스크립트 실행**
```bash
mkdir -p /tmp/report/output
cd /tmp/report
python generate_report.py --from-email --headless
```

**5. 결과 확인**
- output/ 폴더에 Word(.docx)와 PDF 파일이 생성되었는지 확인
- 이메일이 발송되었는지 확인
- 오류 발생 시 오류 내용을 출력하고 중단

### 성공 기준
- Gmail 받은편지함에서 .xlsx 첨부파일을 가진 미읽음 메일 발견
- 보고서 Word + PDF 생성 완료
- 이메일 발송 완료

### 실패 시 처리
- 미읽음 .xlsx 첨부 메일이 없으면 "처리할 새 메일이 없습니다" 출력 후 종료
- 그 외 오류는 상세 오류 메시지 출력
