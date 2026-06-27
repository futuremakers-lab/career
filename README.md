# FuturesImpactLAB Career Report

익명 학생번호를 입력하면 성찰일지, 학과 정보, 2026학년도 2학기 개설 예정 교과목, 교육과정 문서를 바탕으로 진로상담 리포트 초안을 생성하는 로컬 웹 애플리케이션입니다.

## 실행

```bash
pip install -r requirements.txt
python career_report_app.py --serve --host 127.0.0.1 --port 8765
```

브라우저에서 `http://127.0.0.1:8765/`로 접속합니다.

## 포함 데이터

- `성찰일지.xlsx`
- `2026-2 개설예정 교과목.xlsx`
- `교육과정.docx`
- `new01.png`, `new02.png`, `new03.png`

`analytics_events.jsonl`은 로컬 사용 로그이므로 저장소에 포함하지 않습니다.
