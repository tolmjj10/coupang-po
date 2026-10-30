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
=== refresh_sales 2026-10-30T08:35:12 ===
[dash] 당월 매출 1,773,708,289원  입고 1,162,708,849원  ...
[check] DB 최신 2026-10-28  로컬 최신 2026-10-28
[sku] 새 데이터 없음 — 생략
[git] 변경 없음 — push 생략
=== 완료 ===
```

### 3. Windows 작업 스케줄러 등록

**PowerShell로 등록** (경로에 공백 있으면 `schtasks` CLI가 못 다루니 반드시 아래 방식 사용):

```powershell
$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument '/c ""C:\Users\admin\Desktop\17. claude\po\scripts\refresh_sales.bat""'
$trigger = New-ScheduledTaskTrigger -Daily -At 8:35am
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskName "쿠팡 판매 데이터 갱신" -Action $action -Trigger $trigger -Settings $settings -Force
```

- `StartWhenAvailable`: PC가 08:35에 꺼져 있었다면 켜졌을 때 자동 실행
- `ExecutionTimeLimit 1시간`: 1시간 넘으면 강제 종료 (hang 방지)

**지금 바로 테스트:**
```powershell
Start-ScheduledTask -TaskName "쿠팡 판매 데이터 갱신"
# 20초 대기 후:
Get-ScheduledTaskInfo -TaskName "쿠팡 판매 데이터 갱신" | fl LastRunTime, LastTaskResult
Get-Content scripts\refresh_sales.log -Encoding utf8 -Tail 20
```
`LastTaskResult: 0` 이면 성공.

### 4. 작동 확인

- 등록 후 작업 스케줄러에서 해당 작업 우클릭 → **실행** → `refresh_sales.log` 확인
- git push가 성공하려면 Windows 자격 증명 관리자에 GitHub 인증이 캐시돼 있어야 함 (이미 돼 있을 것)

## 주의

- PC가 **08:35 시점에 켜져 있어야** 작동. 꺼져 있으면 다음 켜졌을 때 밀린 작업으로 1회 실행 (작업 스케줄러 옵션 "예약된 시작 시간이 지난 후 작업을 가능한 한 빨리 시작" 체크 시)
- 사내망(192.168.30.x) 접속 가능해야 함 (= 사무실 PC 또는 VPN)
- 로그가 커지면 가끔 `scripts\refresh_sales.log` 삭제

## 환경 변수 (선택)

평소엔 수정 불필요. 바꾸고 싶으면 시스템 환경 변수에 설정:

| 변수 | 기본값 |
|---|---|
| `OA_OPS_MCP_URL` | `http://192.168.30.69:8790/mcp` |
| `OA_OPS_AUTH` | `Bearer oaops_ro_...` (oa_ops.mcp.team.json 참조) |
| `REPO_DIR` | `C:\Users\admin\Desktop\17. claude\po` |
