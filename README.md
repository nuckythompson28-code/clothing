# SOLTRI 피복관리

쏠트리 피복 지급/반납/재고 관리 PWA.

- **프론트**: GitHub Pages — https://nuckythompson28-code.github.io/clothing/
- **백엔드**: Google Apps Script + 구글시트 (`Code.gs`)
- **PDF 보관**: Google Drive `피복관리_PDF/YYYY-MM/`

NAS(192.168.0.60:8588) 독립 서버에서 이전한 버전. 와이파이가 안 터지는 곳에서도
LTE/5G로 접속 가능하고, 오프라인 입력은 IndexedDB 큐에 쌓였다가 자동 동기화된다.

## 구조

| 파일 | 역할 |
|---|---|
| `index.html` | 앱 전체 (단일 파일 PWA) |
| `sw.js` | 서비스 워커 — 오프라인 셸 캐싱 |
| `manifest.json` | PWA 매니페스트 (홈 화면 추가) |
| `Code.gs` | Apps Script 백엔드 소스 (구글시트에 붙여넣어 배포) |
| `피복관리_v7.html` | 구버전 (참고용) |

## 백엔드 설치 (최초 1회)

1. [sheets.new](https://sheets.new) — 새 스프레드시트 생성 (이름: `피복관리DB`)
2. 확장 프로그램 → **Apps Script** → 기본 코드 지우고 `Code.gs` 내용 전체 붙여넣기 → 저장
3. **배포 → 새 배포 → 유형: 웹 앱**
   - 실행 계정: **나**
   - 액세스 권한: **모든 사용자** (익명)
4. 발급된 `https://script.google.com/macros/s/.../exec` URL 복사
5. 앱(github.io 주소) 첫 화면의 설정란에 URL 붙여넣고 **연결**

시트(items/emps/history 등)는 첫 호출 때 자동 생성된다.

## 기존 데이터 이전 (NAS → 시트)

NAS의 `clothing_server` 폴더에서:

```
py -3 export_for_sheets.py
py -3 import_to_sheets.py "https://script.google.com/macros/s/.../exec"
```

## 코드 수정 후 재배포

- 프론트: 이 저장소에 push → Pages 자동 반영 (1~2분)
- 백엔드: Apps Script 편집기에서 수정 → **배포 → 배포 관리 → 새 버전** (URL 유지됨)

## 주의

- 공개 저장소이므로 Apps Script URL·데이터는 절대 커밋하지 말 것 (URL은 기기별 localStorage에만 저장)
- `clothing_export.json` (직원명단·서명 포함) 업로드 금지
