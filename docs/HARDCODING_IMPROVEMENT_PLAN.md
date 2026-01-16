# 하드코딩 문제 해결 방안

## 현재 상태 분석

### ✅ 이미 해결된 부분
1. **`templates/base_generator.py`**: `find_target_col_index` 메서드 구현 완료
   - 연도/분기를 동적으로 찾는 Robust Dynamic Parsing System 구축
   - ValueError 발생으로 어설픈 진행 방지

2. **`templates/construction_generator.py`**: 동적 탐색 적용 완료
   - `ConstructionGenerator` 클래스로 리팩토링
   - `find_target_col_index` 사용

3. **`templates/mining_manufacturing_generator.py`**: 동적 탐색 적용 완료
   - 모든 하드코딩된 컬럼 인덱스 제거
   - `find_target_col_index` 사용

### ❌ 아직 해결되지 않은 부분

#### 1. `templates/domestic_migration_generator.py`
**문제점:**
```python
net_migration_2025_24 = safe_float(row[25], 0)  # 2025.2/4 순이동
net_migration_2025_14 = safe_float(row[24], 0)  # 2025.1/4
net_migration_2024_24 = safe_float(row[21], 0)  # 2024.2/4
net_migration_2023_24 = safe_float(row[17], 0)  # 2023.2/4
```

**해결 방안:**
- `BaseGenerator`를 상속받아 `find_target_col_index` 사용
- 각 연도/분기별로 동적으로 컬럼 인덱스 찾기
- 상대 위치 가정(`-1`, `-4`) 대신 명시적 탐색

#### 2. `templates/employment_rate_generator.py`
**문제점:**
```python
change = safe_float(nationwide_row[18] if len(nationwide_row) > 18 else None, 0)  # 하드코딩
employment_rate = safe_float(index_row[21] if len(index_row) > 21 else None, 60.0)  # 2025.2/4
rate_2024_2 = safe_float(nrow[17], 0)
rate_2025_2 = safe_float(nrow[21], 0)
```

**해결 방안:**
- `BaseGenerator`를 상속받아 `find_target_col_index` 사용
- 기본값 `60.0`을 설정 파일이나 계산된 값으로 대체
- 과거 분기 데이터도 동적으로 찾기

#### 3. 기타 Generator 파일들
- `templates/consumption_generator.py`
- `templates/service_industry_generator.py`
- `templates/price_trend_generator.py`
- `templates/export_generator.py`
- `templates/import_generator.py`
- `templates/unemployment_generator.py`

각 파일에서 하드코딩된 컬럼 인덱스 확인 필요

---

## 해결 전략

### Phase 1: BaseGenerator 상속 구조 확립
모든 Generator가 `BaseGenerator`를 상속받도록 리팩토링

### Phase 2: 동적 컬럼 탐색 적용
1. **현재 분기 찾기**: `find_target_col_index(header_row, year, quarter)`
2. **과거 분기 찾기**: 
   - 전분기: `find_target_col_index(header_row, year, quarter - 1)` 또는 `find_target_col_index(header_row, year, quarter - 1 if quarter > 1 else year - 1, 4)`
   - 전년동분기: `find_target_col_index(header_row, year - 1, quarter)`
   - 2년 전 동분기: `find_target_col_index(header_row, year - 2, quarter)`

### Phase 3: 상대 위치 가정 제거
```python
# ❌ 나쁜 예: 상대 위치 가정
prev_q_col = target_col - 1
prev_y_col = target_col - 4

# ✅ 좋은 예: 명시적 탐색
prev_q_col = self.find_target_col_index(header_row, prev_year, prev_quarter)
prev_y_col = self.find_target_col_index(header_row, year - 1, quarter)
```

### Phase 4: 기본값 하드코딩 제거
```python
# ❌ 나쁜 예
employment_rate = safe_float(index_row[21], 60.0)

# ✅ 좋은 예
employment_rate = safe_float(index_row[target_col], None)
if employment_rate is None:
    # 데이터가 없으면 계산하거나 명시적 오류 처리
    raise ValueError("고용률 데이터를 찾을 수 없습니다")
```

---

## 구체적 구현 예시

### 예시 1: domestic_migration_generator.py 개선

**Before:**
```python
net_migration_2025_24 = safe_float(row[25], 0)  # 2025.2/4
net_migration_2025_14 = safe_float(row[24], 0)  # 2025.1/4
net_migration_2024_24 = safe_float(row[21], 0)  # 2024.2/4
net_migration_2023_24 = safe_float(row[17], 0)  # 2023.2/4
```

**After:**
```python
class DomesticMigrationGenerator(BaseGenerator):
    def __init__(self, excel_path, year, quarter, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
    
    def _find_migration_column(self, header_row, target_year, target_quarter):
        """순인구이동 컬럼 찾기"""
        return self.find_target_col_index(header_row, target_year, target_quarter)
    
    def extract_sido_data(self, summary_df):
        """시도별 순인구이동 데이터 추출"""
        header_row = summary_df.iloc[2]  # 헤더 행 찾기
        
        # 동적으로 각 분기 컬럼 찾기
        col_2025_24 = self._find_migration_column(header_row, 2025, 2)
        col_2025_14 = self._find_migration_column(header_row, 2025, 1)
        col_2024_24 = self._find_migration_column(header_row, 2024, 2)
        col_2023_24 = self._find_migration_column(header_row, 2023, 2)
        
        for i in range(3, len(summary_df)):
            row = summary_df.iloc[i]
            sido = row[4]
            
            if sido in SIDO_ORDER:
                net_migration_2025_24 = safe_float(row[col_2025_24], 0)
                net_migration_2025_14 = safe_float(row[col_2025_14], 0)
                net_migration_2024_24 = safe_float(row[col_2024_24], 0)
                net_migration_2023_24 = safe_float(row[col_2023_24], 0)
                # ...
```

### 예시 2: employment_rate_generator.py 개선

**Before:**
```python
change = safe_float(nationwide_row[18], 0)  # 하드코딩
employment_rate = safe_float(index_row[21], 60.0)  # 2025.2/4
```

**After:**
```python
class EmploymentRateGenerator(BaseGenerator):
    def __init__(self, excel_path, year, quarter, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
    
    def get_nationwide_data(self, df_analysis, df_index):
        """전국 데이터 추출"""
        # 헤더 행 찾기
        header_row_idx = self._find_header_row(df_analysis)
        header_row = df_analysis.iloc[header_row_idx]
        
        # 동적으로 현재 분기 컬럼 찾기
        target_col = self.find_target_col_index(header_row, self.year, self.quarter)
        
        nationwide_row = df_analysis.iloc[3]
        change = safe_float(nationwide_row[target_col], None)
        
        if change is None:
            raise ValueError(f"{self.year}년 {self.quarter}분기 증감률 데이터를 찾을 수 없습니다")
        
        # 집계 시트에서도 동적으로 찾기
        index_header_row = df_index.iloc[header_row_idx]
        index_target_col = self.find_target_col_index(index_header_row, self.year, self.quarter)
        
        index_row = df_index.iloc[3]
        employment_rate = safe_float(index_row[index_target_col], None)
        
        if employment_rate is None:
            raise ValueError(f"{self.year}년 {self.quarter}분기 고용률 데이터를 찾을 수 없습니다")
        
        return {
            'employment_rate': employment_rate,
            'change': change
        }
```

---

## 우선순위

### High Priority (즉시 해결 필요)
1. ✅ `construction_generator.py` - 완료
2. ✅ `mining_manufacturing_generator.py` - 완료
3. 🔴 `domestic_migration_generator.py` - 하드코딩 다수
4. 🔴 `employment_rate_generator.py` - 하드코딩 다수

### Medium Priority
5. `consumption_generator.py`
6. `service_industry_generator.py`
7. `price_trend_generator.py`

### Low Priority (이미 부분적으로 적용됨)
8. `export_generator.py`
9. `import_generator.py`
10. `unemployment_generator.py`

---

## 체크리스트

각 Generator 파일에 대해:
- [ ] `BaseGenerator` 상속 확인
- [ ] 하드코딩된 컬럼 인덱스(`row[N]`, `iloc[N]`) 제거
- [ ] `find_target_col_index` 사용
- [ ] 상대 위치 가정(`-1`, `-4`) 제거
- [ ] 기본값 하드코딩 제거 또는 명시적 오류 처리
- [ ] 테스트: 다른 연도/분기 데이터로 검증

---

## 참고사항

1. **에러 처리**: `find_target_col_index`는 찾지 못하면 `ValueError` 발생
   - 어설프게 기본값 사용하지 말고 명시적 오류 처리
   
2. **성능**: 동적 탐색은 약간의 오버헤드가 있지만, 유지보수성과 정확성이 더 중요

3. **하위 호환성**: 기존 코드와의 호환성을 위해 점진적으로 리팩토링
