# 재고매칭 프로그램

재고현황과 buyList의 데이터를 진벌리 제품위치 파일에 매칭/병합합니다.

## 다운로드

최신 버전: **[Releases 페이지](../../releases/latest)**

## 기능

- 재고현황(현재고-가용) → 진벌리 재고 컬럼 매칭
- buyList(소비기한) → 진벌리 유통기한 컬럼 매칭 (동일 품번이면 최신 날짜)
- 외부 데이터 연결 자동 제거 (한셀/엑셀 팝업 방지)
- 자동 업데이트 알림

## 사용법

1. exe 실행 → 파일 선택창에서 3개 파일 동시 선택
   - 재고현황조회Ⅱ_다운로드.xlsx
   - 진벌리 제품 위치(2024).xlsx
   - buyList_*.xlsx
2. 결과: `바탕화면\재고매칭\YYYY-MM-DD 오전(또는 오후).xlsx`
3. 재고현황 파일은 자동 삭제됨

## 릴리스 방법

```bash
echo 1.0.1 > version.txt
git add version.txt 재고매칭.py
git commit -m "v1.0.1"
git tag v1.0.1
git push origin main --tags
```
