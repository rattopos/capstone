# 데이터 파싱 오류 분석

## 개요
데이터 무결성 원칙에 위배되는 데이터 파싱 관련 오류들을 정리한 문서입니다.

---

## 🔴 1. 데이터 파싱 실패 시 기본값 사용 문제

### 1.1 `data_converter.py` - `_safe_float` 함수

**위치**: `data_converter.py:1004-1011`

**문제점**:
```python
def _safe_float(self, val) -> float:
    """안전하게 float으로 변환"""
    if pd.isna(val):
        return 0.0  # ❌ 결측치를 0.0으로 변환
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0  # ❌ 파싱 실패 시 0.0 반환
```

**문제**:
- 결측치(`pd.isna(val)`)를 0.0으로 변환 → 데이터 왜곡
- 파싱 실패 시 0.0 반환 → 실제 값과 구분 불가능
- 데이터 무결성 원칙 위배

**해결 방향**:
- 결측치는 `None` 반환
- 파싱 실패도 `None` 반환
- 템플릿에서 `None`을 'N/A' 또는 빈 값으로 표시

---

### 1.2 `services/grdp_service.py` - `safe_float` 함수

**위치**: `services/grdp_service.py:173-179`

**문제점**:
```python
def safe_float(val, default=0.0):
    try:
        if pd.isna(val):
            return default  # ❌ 기본값 0.0 반환
        return round(float(val), 1)
    except:
        return default  # ❌ 파싱 실패 시 기본값 0.0 반환
```

**문제**:
- 결측치와 파싱 실패를 기본값(0.0)으로 처리 → 데이터 왜곡
- 데이터 무결성 원칙 위배

**해결 방향**:
- `default` 파라미터 제거 또는 `None` 사용
- 결측치/파싱 실패 시 `None` 반환

---

## 🔴 2. 예외 처리로 인한 데이터 손실

### 2.1 `services/summary_data.py` - 예외 무시

**위치**: `services/summary_data.py:396-397`

**문제점**:
```python
for i, row in df.iterrows():
    try:
        # 데이터 추출 로직
        ...
    except Exception as e:
        continue  # ❌ 오류 발생 시 해당 행을 완전히 무시
```

**문제**:
- 예외 발생 시 해당 행의 데이터를 완전히 무시
- 오류 로깅 없이 조용히 실패
- 데이터 손실 발생

**해결 방향**:
- 예외 발생 시 로깅 추가
- 필요한 경우 부분 데이터라도 처리
- 명시적인 오류 처리

---

### 2.2 `services/grdp_service.py` - 예외 무시

**위치**: `services/grdp_service.py:91-92`

**문제점**:
```python
except:
    return False  # ❌ 모든 예외를 무시
```

**문제**:
- 모든 예외를 무시하고 `False` 반환
- 오류 원인 파악 불가능
- 데이터 손실 발생

**해결 방향**:
- 명시적인 예외 처리
- 오류 로깅 추가
- 실패 원인 추적 가능하도록 개선

---

## 🔴 3. 연도/분기 추출 실패 시 기본값 사용

### 3.1 `data_converter.py` - 기본값 2025, 2

**위치**: `data_converter.py:203-205`

**문제점**:
```python
# 3. 기본값 (실패 시)
self.year, self.quarter = 2025, 2
print(f"[경고] 연도/분기 추출 실패, 기본값 사용: {self.year}년 {self.quarter}분기")
```

**문제**:
- 연도/분기 추출 실패 시 하드코딩된 기본값 사용
- 데이터 무결성 원칙 위배
- 잘못된 연도/분기로 보도자료 생성 가능

**해결 방향**:
- 기본값 제거
- 추출 실패 시 `ValueError` 발생
- 사용자에게 명시적으로 입력 요청

---

### 3.2 `services/grdp_service.py` - 기본값 2025, 2

**위치**: `services/grdp_service.py:104-105`

**문제점**:
```python
# 연도/분기 기본값
target_year = year or 2025  # ❌
target_quarter = quarter or 2  # ❌
```

**문제**:
- `year`나 `quarter`가 `None`일 때 기본값 사용
- 데이터 무결성 원칙 위배

**해결 방향**:
- 기본값 제거
- `None` 체크 후 명시적 오류 처리

---

## 🔴 4. 데이터 검증 부족

### 4.1 `.get()` 메서드의 기본값 사용

**문제점**:
코드 전반에 걸쳐 `.get(key, default_value)` 패턴이 많이 사용되고 있습니다. 이 중 일부는 데이터 무결성 원칙에 위배될 수 있습니다.

**예시**:
- `grdp_data.get('report_info', {}).get('year', 2025)` - 연도 기본값
- `region.get('growth_rate', 0.0)` - 성장률 기본값 0.0
- `region.get('region', '-')` - 지역명 기본값 '-'

**해결 방향**:
- 기본값이 실제 데이터로 오인될 수 있는 경우 제거
- `None` 반환 후 템플릿에서 처리
- 필수 데이터는 검증 로직 추가

---

## 📋 요약

| 항목 | 위치 | 문제점 | 우선순위 |
|------|------|--------|---------|
| `_safe_float` 함수 | `data_converter.py:1004-1011` | 결측치/파싱 실패 시 0.0 반환 | 🔴 높음 |
| `safe_float` 함수 | `services/grdp_service.py:173-179` | 기본값 0.0 사용 | 🔴 높음 |
| 예외 무시 | `services/summary_data.py:396` | `except Exception: continue` | 🔴 높음 |
| 예외 무시 | `services/grdp_service.py:91` | `except: return False` | 🔴 높음 |
| 연도/분기 기본값 | `data_converter.py:204` | 기본값 2025, 2 | 🔴 높음 |
| 연도/분기 기본값 | `services/grdp_service.py:104-105` | 기본값 2025, 2 | 🔴 높음 |
| 데이터 검증 | 전반 | `.get()` 기본값 사용 | 🟡 중간 |

---

## 💡 해결 방향 (우선순위 순)

### 최우선 (데이터 왜곡 방지)
1. **`_safe_float` 및 `safe_float` 함수 수정**: 결측치/파싱 실패 시 `None` 반환
2. **예외 처리 개선**: 예외 무시 제거, 오류 로깅 추가
3. **연도/분기 기본값 제거**: 추출 실패 시 에러 발생

### 중요 (데이터 검증 강화)
4. **데이터 검증 로직 추가**: 필수 데이터 누락 시 명시적 오류 처리
5. **기본값 사용 검토**: 실제 데이터로 오인될 수 있는 기본값 제거

---

## 참고사항

- 모든 수정사항은 `.cursorrules`의 **데이터 무결성 원칙**을 준수해야 합니다
- 결측치는 절대 기본값, 추정값, 평균값으로 채우지 말 것
- 결측치는 반드시 'N/A', None, 또는 빈 값으로 표시
- 데이터에 어떤 왜곡도 들어가서는 안 됨
