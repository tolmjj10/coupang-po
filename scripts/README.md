# 판매 데이터 자동 갱신

매일 아침 사내 oa_ops DB에서 판매 데이터를 끌어와 JSON 파일들을 갱신하고 git push.

## 파일

- `refresh_sales.py` — 메인 Python 스크립트
- `refresh_sales.bat` — Windows 작업 스케줄러용 배치 (python 호출 + 로그)
- `refresh_sales.log` — 실행 로그 (자동 생성)

## 갱신 대상

| 파일 | 조건 |
|---|---|
| `data/sales-dash.json` | **매번** (대시보드 KPI/6개월 차트) |
| `data/sales-sku/{YYYY-MM}.json` | DB(`crawl_coupang_daily`)에 로컬보다 새 데이터가 있을 때만 |
| `data/sales-sku/index.json` | 월별 totals 재계산 |

## 설치 (한 번만)

### 1. 의존성

```powershell
pip install requests
```

### 2. 수동 테스트

```powershell
cd "C:\Users\admin\Desktop\17. claude\po"
python scripts\refresh_sales.py
```

정상 출력:
```
=== refresh_sales 2026-10-30T09:00:12 ===
[dash] 당월 매출 1,773,708,289원  입고 1,162,708,849원  ...
[check] DB 최신 2026-10-28  로컬 최신 2026-10-28
[sku] 새 데이터 없음 — 생략
[git] 변경 없음 — push 생략
=== 완료 ===
```

### 3. Windows 작업 스케줄러 등록

**방법 A: PowerShell 1줄로 등록 (관리자 권한 필요 없음)**

```powershell
schtasks /create /tn "쿠팡 판매 데이터 갱신" /tr "C:\Users\admin\Desktop\17. claude\po\scripts\refresh_sales.bat" /sc daily /st 09:00 /f
```

**방법 B: GUI**
1. `작업 스케줄러` 열기 → `기본 작업 만들기`
2. 이름: `쿠팡 판매 데이터 갱신`
3. 트리거: **매일**, 시작 **09:00**
4. 동작: **프로그램 시작**, 프로그램: `C:\Users\admin\Desktop\17. claude\po\scripts\refresh_sales.bat`
5. 마침

### 4. 작동 확인

- 등록 후 작업 스케줄러에서 해당 작업 우클릭 → **실행** → `refresh_sales.log` 확인
- git push가 성공하려면 Windows 자격 증명 관리자에 GitHub 인증이 캐시돼 있어야 함 (이미 돼 있을 것)

## 주의

- PC가 **09:00 시점에 켜져 있어야** 작동. 꺼져 있으면 다음 켜졌을 때 밀린 작업으로 1회 실행 (작업 스케줄러 옵션 "예약된 시작 시간이 지난 후 작업을 가능한 한 빨리 시작" 체크 시)
- 사내망(192.168.30.x) 접속 가능해야 함 (= 사무실 PC 또는 VPN)
- 로그가 커지면 가끔 `scripts\refresh_sales.log` 삭제

## 환경 변수 (선택)

평소엔 수정 불필요. 바꾸고 싶으면 시스템 환경 변수에 설정:

| 변수 | 기본값 |
|---|---|
| `OA_OPS_MCP_URL` | `http://192.168.30.69:8790/mcp` |
| `OA_OPS_AUTH` | `Bearer oaops_ro_...` (oa_ops.mcp.team.json 참조) |
| `REPO_DIR` | `C:\Users\admin\Desktop\17. claude\po` |
