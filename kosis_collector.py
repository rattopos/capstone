#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KOSIS API 기초자료 수집 모듈

국가데이터처 KOSIS Open API를 통해 통계 데이터를 수집하고
기초자료 수집표 엑셀 파일을 생성합니다.

사용법:
    # CLI로 실행
    python kosis_collector.py --config kosis_config.json --output 기초자료_수집표.xlsx
    
    # Python에서 import
    from kosis_collector import KOSISCollector
    collector = KOSISCollector('kosis_config.json')
    collector.fetch_all()
    collector.to_excel('output.xlsx')
"""

import os
import json
import argparse
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime

import requests
import pandas as pd
from dotenv import load_dotenv

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


class KOSISCollector:
    """
    KOSIS Open API를 통해 통계 데이터를 수집하는 클래스
    
    Attributes:
        config: JSON 설정 파일에서 로드한 설정
        api_key: KOSIS API 키
        collected_data: 수집된 데이터를 저장하는 딕셔너리
    """
    
    # KOSIS API 기본 엔드포인트
    BASE_URL = "https://kosis.kr/openapi/Param/statisticsParameterData.do"
    
    def __init__(self, config_path: str, env_path: Optional[str] = None):
        """
        초기화
        
        Args:
            config_path: JSON 설정 파일 경로
            env_path: .env 파일 경로 (기본값: 프로젝트 루트의 .env)
        """
        self.config_path = Path(config_path)
        self.collected_data: Dict[str, pd.DataFrame] = {}
        
        # .env 파일 로드
        if env_path:
            load_dotenv(env_path)
        else:
            load_dotenv()
        
        # API 키 로드
        self.api_key = os.getenv('KOSIS_API_KEY')
        if not self.api_key:
            raise ValueError(
                "KOSIS_API_KEY가 설정되지 않았습니다. "
                ".env 파일에 KOSIS_API_KEY=your_api_key 형식으로 설정해주세요."
            )
        
        # 설정 파일 로드
        self.config = self._load_config()
        logger.info(f"설정 파일 로드 완료: {config_path}")
        logger.info(f"등록된 통계표: {list(self.config.get('statistics', {}).keys())}")
    
    def _load_config(self) -> Dict:
        """설정 파일 로드"""
        if not self.config_path.exists():
            raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {self.config_path}")
        
        with open(self.config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _build_request_url(self, stat_config: Dict) -> str:
        """
        API 요청 URL 구성
        
        Args:
            stat_config: 개별 통계표 설정
            
        Returns:
            완성된 API 요청 URL
        """
        # 사용자가 직접 URL을 지정한 경우
        if 'url' in stat_config and stat_config['url']:
            url = stat_config['url']
            
            # {api_key} 플레이스홀더 치환
            if '{api_key}' in url:
                url = url.replace('{api_key}', self.api_key)
            # URL에 apiKey가 없으면 추가
            elif 'apiKey=' not in url:
                separator = '&' if '?' in url else '?'
                url = f"{url}{separator}apiKey={self.api_key}"
            return url
        
        # 기본 파라미터로 URL 구성
        params = {
            'method': 'getList',
            'apiKey': self.api_key,
            'format': 'json',
            'jsonVD': 'Y',
            'userStatsId': stat_config.get('userStatsId', ''),
            'orgId': stat_config.get('orgId', ''),
            'tblId': stat_config.get('tblId', ''),
        }
        
        # 추가 파라미터 병합
        if 'params' in stat_config:
            params.update(stat_config['params'])
        
        # URL 파라미터 문자열 생성
        param_str = '&'.join(f"{k}={v}" for k, v in params.items() if v)
        return f"{self.BASE_URL}?{param_str}"
    
    def fetch_data(self, stat_id: str) -> Optional[pd.DataFrame]:
        """
        단일 통계표 데이터 수집
        
        Args:
            stat_id: 통계표 ID (설정 파일의 키)
            
        Returns:
            수집된 데이터의 DataFrame, 실패 시 None
        """
        if stat_id not in self.config.get('statistics', {}):
            logger.error(f"통계표 '{stat_id}'가 설정 파일에 없습니다.")
            return None
        
        stat_config = self.config['statistics'][stat_id]
        
        # URL이 비어있는지 확인
        if not stat_config.get('url') and not stat_config.get('tblId'):
            logger.warning(f"'{stat_id}': URL 또는 tblId가 설정되지 않았습니다. 건너뜁니다.")
            return None
        
        url = self._build_request_url(stat_config)
        logger.info(f"'{stat_id}' 데이터 수집 중...")
        
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            
            # 에러 응답 확인
            if isinstance(data, dict) and 'err' in data:
                logger.error(f"'{stat_id}' API 에러: {data.get('err', '알 수 없는 에러')}")
                return None
            
            # 빈 응답 확인
            if not data:
                logger.warning(f"'{stat_id}': 데이터가 없습니다.")
                return None
            
            # DataFrame 변환
            df = pd.DataFrame(data)
            self.collected_data[stat_id] = df
            
            logger.info(f"'{stat_id}' 수집 완료: {len(df)} 행")
            return df
            
        except requests.exceptions.Timeout:
            logger.error(f"'{stat_id}' 요청 시간 초과")
        except requests.exceptions.RequestException as e:
            logger.error(f"'{stat_id}' 요청 실패: {e}")
        except json.JSONDecodeError:
            logger.error(f"'{stat_id}' JSON 파싱 실패")
        except Exception as e:
            logger.error(f"'{stat_id}' 처리 중 오류: {e}")
        
        return None
    
    def fetch_all(self) -> Dict[str, pd.DataFrame]:
        """
        전체 통계표 데이터 일괄 수집
        
        Returns:
            수집된 데이터 딕셔너리 {stat_id: DataFrame}
        """
        statistics = self.config.get('statistics', {})
        
        if not statistics:
            logger.warning("수집할 통계표가 없습니다.")
            return {}
        
        logger.info(f"총 {len(statistics)}개 통계표 수집 시작")
        
        success_count = 0
        for stat_id in statistics:
            result = self.fetch_data(stat_id)
            if result is not None:
                success_count += 1
        
        logger.info(f"수집 완료: {success_count}/{len(statistics)}개 성공")
        return self.collected_data
    
    def to_excel(self, output_path: str, include_metadata: bool = True) -> str:
        """
        수집된 데이터를 엑셀 파일로 저장
        
        Args:
            output_path: 출력 파일 경로
            include_metadata: 메타데이터 시트 포함 여부
            
        Returns:
            저장된 파일 경로
        """
        if not self.collected_data:
            logger.warning("저장할 데이터가 없습니다. 먼저 fetch_data() 또는 fetch_all()을 호출하세요.")
            return ""
        
        output_path = Path(output_path)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 메타데이터 시트 추가
            if include_metadata:
                metadata = {
                    '항목': ['생성일시', '수집 통계표 수', 'API 버전', '설정 파일'],
                    '값': [
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        len(self.collected_data),
                        'KOSIS Open API v1.0',
                        str(self.config_path)
                    ]
                }
                pd.DataFrame(metadata).to_excel(writer, sheet_name='_메타데이터', index=False)
            
            # 각 통계표 데이터 저장
            for stat_id, df in self.collected_data.items():
                # 시트명 결정 (설정에 지정된 이름 또는 stat_id 사용)
                stat_config = self.config['statistics'].get(stat_id, {})
                sheet_name = stat_config.get('sheet_name', stat_id)
                
                # 시트명 길이 제한 (Excel 최대 31자)
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"시트 '{sheet_name}' 저장 완료")
        
        logger.info(f"엑셀 파일 저장 완료: {output_path}")
        return str(output_path)
    
    def get_statistics_list(self) -> List[str]:
        """설정된 통계표 목록 반환"""
        return list(self.config.get('statistics', {}).keys())
    
    def get_collected_data(self, stat_id: str) -> Optional[pd.DataFrame]:
        """특정 통계표의 수집된 데이터 반환"""
        return self.collected_data.get(stat_id)
    
    def clear_collected_data(self):
        """수집된 데이터 초기화"""
        self.collected_data.clear()
        logger.info("수집 데이터 초기화 완료")


def main():
    """CLI 메인 함수"""
    parser = argparse.ArgumentParser(
        description='KOSIS API 기초자료 수집기',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  # 전체 통계표 수집
  python kosis_collector.py --config kosis_config.json --output 기초자료_수집표.xlsx
  
  # 특정 통계표만 수집
  python kosis_collector.py --config kosis_config.json --stat 광공업생산 --output output.xlsx
  
  # 여러 통계표 수집
  python kosis_collector.py --config kosis_config.json --stat 광공업생산 서비스업생산 --output output.xlsx
  
  # 설정된 통계표 목록 확인
  python kosis_collector.py --config kosis_config.json --list
        """
    )
    
    parser.add_argument(
        '--config', '-c',
        required=True,
        help='JSON 설정 파일 경로'
    )
    parser.add_argument(
        '--output', '-o',
        default='기초자료_수집표.xlsx',
        help='출력 엑셀 파일 경로 (기본값: 기초자료_수집표.xlsx)'
    )
    parser.add_argument(
        '--stat', '-s',
        nargs='+',
        help='수집할 통계표 ID (지정하지 않으면 전체 수집)'
    )
    parser.add_argument(
        '--list', '-l',
        action='store_true',
        help='설정된 통계표 목록만 출력'
    )
    parser.add_argument(
        '--env',
        help='.env 파일 경로 (기본값: 현재 디렉토리의 .env)'
    )
    parser.add_argument(
        '--no-metadata',
        action='store_true',
        help='메타데이터 시트 제외'
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='상세 로그 출력'
    )
    
    args = parser.parse_args()
    
    # 로깅 레벨 조정
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        collector = KOSISCollector(args.config, args.env)
        
        # 목록 출력 모드
        if args.list:
            print("\n등록된 통계표 목록:")
            print("-" * 40)
            for stat_id in collector.get_statistics_list():
                stat_config = collector.config['statistics'][stat_id]
                sheet_name = stat_config.get('sheet_name', stat_id)
                has_url = bool(stat_config.get('url') or stat_config.get('tblId'))
                status = "✓" if has_url else "✗ (URL 미설정)"
                print(f"  {status} {stat_id} → {sheet_name}")
            return
        
        # 데이터 수집
        if args.stat:
            # 특정 통계표만 수집
            for stat_id in args.stat:
                collector.fetch_data(stat_id)
        else:
            # 전체 수집
            collector.fetch_all()
        
        # 엑셀 저장
        if collector.collected_data:
            collector.to_excel(args.output, include_metadata=not args.no_metadata)
            print(f"\n✓ 수집 완료: {args.output}")
        else:
            print("\n✗ 수집된 데이터가 없습니다.")
            
    except FileNotFoundError as e:
        print(f"오류: {e}")
        return 1
    except ValueError as e:
        print(f"설정 오류: {e}")
        return 1
    except Exception as e:
        print(f"예상치 못한 오류: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1
    
    return 0


if __name__ == '__main__':
    exit(main())

