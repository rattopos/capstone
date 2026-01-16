#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
어휘 매핑 및 나레이션 패턴 단위 테스트

[문서 1] 어휘 매핑 규칙 검증
[문서 2] 4가지 나레이션 패턴 검증
[문서 3] 기여도 정렬 검증
"""

import sys
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from utils.text_utils import get_terms, get_josa


def test_vocabulary_mapping():
    """[문서 1] 어휘 매핑 규칙 검증"""
    print("=" * 60)
    print("[Test 1] 어휘 매핑 규칙 검증")
    print("=" * 60)
    
    # Type A (물량 지표): 광공업생산
    print("\n[Type A - 물량 지표]")
    
    # 증가
    cause, result = get_terms('manufacturing', 5.2)
    assert cause == '늘어', f"증가 원인 어휘 오류: {cause} (기대: 늘어)"
    assert result == '증가', f"증가 결과 어휘 오류: {result} (기대: 증가)"
    print(f"  ✓ 증가 (5.2%): {cause}, {result}")
    
    # 감소
    cause, result = get_terms('manufacturing', -3.1)
    assert cause == '줄어', f"감소 원인 어휘 오류: {cause} (기대: 줄어)"
    assert result == '감소', f"감소 결과 어휘 오류: {result} (기대: 감소)"
    print(f"  ✓ 감소 (-3.1%): {cause}, {result}")
    
    # 보합
    cause, result = get_terms('manufacturing', 0.0)
    assert cause is None, f"보합 원인 어휘 오류: {cause} (기대: None)"
    assert result == '보합', f"보합 결과 어휘 오류: {result} (기대: 보합)"
    print(f"  ✓ 보합 (0.0%): {cause}, {result}")
    
    # Type B (가격 지표): 물가, 고용률, 실업률
    print("\n[Type B - 가격/비율 지표]")
    
    # 상승
    cause, result = get_terms('price', 2.1)
    assert cause == '올라', f"상승 원인 어휘 오류: {cause} (기대: 올라)"
    assert result == '상승', f"상승 결과 어휘 오류: {result} (기대: 상승)"
    print(f"  ✓ 상승 (2.1%): {cause}, {result}")
    
    # 하락
    cause, result = get_terms('price', -1.5)
    assert cause == '내려', f"하락 원인 어휘 오류: {cause} (기대: 내려)"
    assert result == '하락', f"하락 결과 어휘 오류: {result} (기대: 하락)"
    print(f"  ✓ 하락 (-1.5%): {cause}, {result}")
    
    # 보합
    cause, result = get_terms('employment', 0.0)
    assert cause is None, f"보합 원인 어휘 오류: {cause} (기대: None)"
    assert result == '보합', f"보합 결과 어휘 오류: {result} (기대: 보합)"
    print(f"  ✓ 보합 (0.0%): {cause}, {result}")
    
    print("\n✅ 어휘 매핑 규칙 검증 통과!")


def test_josa_processing():
    """조사 처리 검증"""
    print("\n" + "=" * 60)
    print("[Test 2] 조사 처리 검증")
    print("=" * 60)
    
    # 은/는
    assert get_josa('서울', '은/는') == '은', "서울은 오류"
    assert get_josa('경기', '은/는') == '는', "경기는 오류"
    assert get_josa('부산', '은/는') == '은', "부산은 오류"
    assert get_josa('인천', '은/는') == '은', "인천은 오류"
    assert get_josa('대구', '은/는') == '는', "대구는 오류"
    
    print("  ✓ 서울은, 경기는, 부산은, 인천은, 대구는")
    
    # 이/가
    assert get_josa('업종', '이/가') == '이', "업종이 오류"
    assert get_josa('소비', '이/가') == '가', "소비가 오류"
    
    print("  ✓ 업종이, 소비가")
    
    print("\n✅ 조사 처리 검증 통과!")


def test_pattern_selection():
    """[문서 2] 패턴 선택 로직 검증"""
    print("\n" + "=" * 60)
    print("[Test 3] 나레이션 패턴 선택 검증")
    print("=" * 60)
    
    # Dummy generator 생성
    from templates.base_generator import BaseGenerator
    
    class DummyGenerator(BaseGenerator):
        def extract_all_data(self):
            return {}
    
    gen = DummyGenerator('dummy.xlsx', year=2025, quarter=2)
    
    # 패턴 A: 순접 (일반적인 증감)
    pattern = gen.select_narrative_pattern(growth_rate=5.2)
    assert pattern == 'pattern_a', f"패턴 A 선택 오류: {pattern}"
    print(f"  ✓ 패턴 A (순접): growth_rate=5.2 → {pattern}")
    
    # 패턴 B: 역접 (상반된 업종 혼재)
    pattern = gen.select_narrative_pattern(growth_rate=5.2, has_contrast_industries=True)
    assert pattern == 'pattern_b', f"패턴 B 선택 오류: {pattern}"
    print(f"  ✓ 패턴 B (역접): growth_rate=5.2, contrast=True → {pattern}")
    
    # 패턴 C: 보합
    pattern = gen.select_narrative_pattern(growth_rate=0.0)
    assert pattern == 'pattern_c', f"패턴 C 선택 오류: {pattern}"
    print(f"  ✓ 패턴 C (보합): growth_rate=0.0 → {pattern}")
    
    # 패턴 D: 방향 전환
    pattern = gen.select_narrative_pattern(growth_rate=5.2, prev_rate=-3.1)
    assert pattern == 'pattern_d', f"패턴 D 선택 오류: {pattern}"
    print(f"  ✓ 패턴 D (방향 전환): growth_rate=5.2, prev_rate=-3.1 → {pattern}")
    
    print("\n✅ 패턴 선택 로직 검증 통과!")


def test_narrative_generation():
    """[문서 2] 나레이션 생성 검증"""
    print("\n" + "=" * 60)
    print("[Test 4] 나레이션 생성 검증")
    print("=" * 60)
    
    from templates.base_generator import BaseGenerator
    
    class DummyGenerator(BaseGenerator):
        def extract_all_data(self):
            return {}
    
    gen = DummyGenerator('dummy.xlsx', year=2025, quarter=2)
    
    # 패턴 A: 순접
    narrative = gen.generate_narrative(
        pattern='pattern_a',
        region='서울',
        growth_rate=5.2,
        prev_rate=None,
        main_industries=['반도체·전자부품', '자동차·트레일러'],
        report_id='manufacturing'
    )
    print(f"\n  패턴 A (순접):")
    print(f"    {narrative}")
    assert '늘어' in narrative, "원인 어휘 누락"
    assert '증가' in narrative, "결과 어휘 누락"
    assert '서울은' in narrative, "조사 오류"
    
    # 패턴 B: 역접
    narrative = gen.generate_narrative(
        pattern='pattern_b',
        region='경기',
        growth_rate=3.5,
        prev_rate=None,
        main_industries=['반도체·전자부품'],
        contrast_industries=['식료품', '섬유제품'],
        report_id='manufacturing'
    )
    print(f"\n  패턴 B (역접):")
    print(f"    {narrative}")
    assert '줄었으나' in narrative or '늘었으나' in narrative, "역접 어휘 누락"
    assert '경기는' in narrative, "조사 오류"
    
    # 패턴 C: 보합
    narrative = gen.generate_narrative(
        pattern='pattern_c',
        region='대전',
        growth_rate=0.0,
        prev_rate=None,
        main_industries=['반도체·전자부품'],
        contrast_industries=['식료품'],
        report_id='manufacturing'
    )
    print(f"\n  패턴 C (보합):")
    print(f"    {narrative}")
    assert '보합' in narrative, "보합 어휘 누락"
    assert '늘었으나' in narrative and '줄어' in narrative, "보합 패턴 오류"
    
    # 패턴 D: 방향 전환
    narrative = gen.generate_narrative(
        pattern='pattern_d',
        region='부산',
        growth_rate=4.2,
        prev_rate=-2.5,
        main_industries=['자동차·트레일러'],
        report_id='manufacturing'
    )
    print(f"\n  패턴 D (방향 전환):")
    print(f"    {narrative}")
    assert '전분기' in narrative, "전분기 언급 누락"
    assert '감소하였으나' in narrative or '증가하였으나' in narrative, "방향 전환 어휘 누락"
    
    print("\n✅ 나레이션 생성 검증 통과!")


def test_contribution_ranking():
    """[문서 3] 기여도 정렬 검증"""
    print("\n" + "=" * 60)
    print("[Test 5] 기여도 정렬 검증")
    print("=" * 60)
    
    from templates.base_generator import BaseGenerator
    
    class DummyGenerator(BaseGenerator):
        def extract_all_data(self):
            return {}
    
    gen = DummyGenerator('dummy.xlsx', year=2025, quarter=2)
    
    # 테스트 데이터: 증감률은 크지만 가중치가 작은 업종 vs 증감률은 작지만 가중치가 큰 업종
    industries = [
        {'name': '식료품', 'change_rate': 10.0, 'weight': 50},  # 기여도: 500
        {'name': '반도체·전자부품', 'change_rate': 5.0, 'weight': 300},  # 기여도: 1500 (최고)
        {'name': '섬유제품', 'change_rate': 15.0, 'weight': 20},  # 기여도: 300
        {'name': '자동차·트레일러', 'change_rate': 3.0, 'weight': 200},  # 기여도: 600
    ]
    
    ranked = gen.rank_by_contribution(industries, top_n=3)
    
    print(f"\n  입력:")
    for ind in industries:
        contrib = abs(ind['change_rate'] * ind['weight'])
        print(f"    - {ind['name']:20s}: 증감률={ind['change_rate']:5.1f}%, 가중치={ind['weight']:4d}, 기여도={contrib:7.1f}")
    
    print(f"\n  출력 (기여도 순):")
    for i, ind in enumerate(ranked, 1):
        print(f"    {i}. {ind['name']:20s}: 기여도={ind['contribution']:7.1f}")
    
    # 검증: 반도체가 1위여야 함 (기여도 1500)
    assert ranked[0]['name'] == '반도체·전자부품', f"1위 오류: {ranked[0]['name']} (기대: 반도체·전자부품)"
    assert ranked[1]['name'] == '자동차·트레일러', f"2위 오류: {ranked[1]['name']} (기대: 자동차·트레일러)"
    assert ranked[2]['name'] == '식료품', f"3위 오류: {ranked[2]['name']} (기대: 식료품)"
    
    print("\n✅ 기여도 정렬 검증 통과!")


def test_integration():
    """통합 테스트: 전체 플로우"""
    print("\n" + "=" * 60)
    print("[Test 6] 통합 테스트")
    print("=" * 60)
    
    from templates.base_generator import BaseGenerator
    
    class DummyGenerator(BaseGenerator):
        def extract_all_data(self):
            return {}
    
    gen = DummyGenerator('dummy.xlsx', year=2025, quarter=2)
    
    # 시나리오: 광공업생산 증가 (5.2%)
    industries = [
        {'name': '반도체·전자부품', 'change_rate': 8.5, 'weight': 300},
        {'name': '자동차·트레일러', 'change_rate': 4.2, 'weight': 200},
        {'name': '식료품', 'change_rate': -2.1, 'weight': 50},
    ]
    
    # 1. 기여도 정렬
    top_increase = [i for i in industries if i['change_rate'] > 0]
    ranked = gen.rank_by_contribution(top_increase, top_n=2)
    
    print(f"\n  [Step 1] 기여도 정렬:")
    for i, ind in enumerate(ranked, 1):
        print(f"    {i}. {ind['name']} (기여도: {ind['contribution']:.1f})")
    
    # 2. 패턴 선택
    pattern = gen.select_narrative_pattern(
        growth_rate=5.2,
        prev_rate=None,
        has_contrast_industries=False
    )
    print(f"\n  [Step 2] 패턴 선택: {pattern}")
    
    # 3. 나레이션 생성
    main_industry_names = [ind['name'] for ind in ranked]
    narrative = gen.generate_narrative(
        pattern=pattern,
        region='전국',
        growth_rate=5.2,
        prev_rate=None,
        main_industries=main_industry_names,
        report_id='manufacturing'
    )
    
    print(f"\n  [Step 3] 생성된 나레이션:")
    print(f"    \"{narrative}\"")
    
    # 검증
    assert '전국은' in narrative, "조사 오류"
    assert '늘어' in narrative, "Type A 원인 어휘 오류"
    assert '증가' in narrative, "Type A 결과 어휘 오류"
    assert '5.2%' in narrative, "수치 누락"
    assert '반도체·전자부품' in narrative, "주요 업종 누락"
    
    # 금지어 체크 (Type A는 상승/하락/올라/내려 사용 불가)
    forbidden_words = ['상승', '하락', '올라', '내려']
    for word in forbidden_words:
        assert word not in narrative, f"금지어 사용: {word}"
    
    print("\n✅ 통합 테스트 통과!")


def main():
    """전체 테스트 실행"""
    print("\n")
    print("╔" + "═" * 58 + "╗")
    print("║" + " " * 15 + "어휘 매핑 리팩토링 검증" + " " * 16 + "║")
    print("╚" + "═" * 58 + "╝")
    
    try:
        test_vocabulary_mapping()
        test_josa_processing()
        test_pattern_selection()
        test_narrative_generation()
        test_contribution_ranking()
        test_integration()
        
        print("\n" + "=" * 60)
        print("🎉 모든 테스트 통과!")
        print("=" * 60)
        print("\n[요약]")
        print("  ✓ 어휘 매핑: Type A (물량) / Type B (가격) 분리 완료")
        print("  ✓ 조사 처리: 받침 유무에 따른 동적 선택 완료")
        print("  ✓ 패턴 선택: 4가지 패턴 분기 완료")
        print("  ✓ 나레이션 생성: 엄격한 어휘 매핑 준수")
        print("  ✓ 기여도 정렬: |증감률 × 가중치| 순 정렬 완료")
        print("\n[다음 단계]")
        print("  → 실제 엑셀 데이터로 mining_manufacturing_generator.py 테스트")
        print("  → 나머지 8개 generator에 동일 패턴 적용")
        
        return 0
        
    except AssertionError as e:
        print(f"\n❌ 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return 1
    except Exception as e:
        print(f"\n❌ 테스트 오류: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
