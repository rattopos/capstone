"""
계산 엔진 모듈
데이터에 대한 계산 수행 (합계, 평균, 증감률 등)
"""

from typing import List, Any, Union, Optional
import statistics


class Calculator:
    """데이터 계산을 수행하는 클래스"""
    
    def __init__(self):
        """계산기 초기화"""
        pass
    
    @staticmethod
    def _to_numeric(value: Any) -> Optional[float]:
        """
        값을 숫자로 변환합니다.
        
        Args:
            value: 변환할 값
            
        Returns:
            숫자로 변환된 값 (변환 불가능하면 None)
        """
        if value is None:
            return None
        
        if isinstance(value, (int, float)):
            return float(value)
        
        if isinstance(value, str):
            # 문자열에서 숫자 추출 (콤마 제거 등)
            cleaned = value.replace(',', '').strip()
            try:
                return float(cleaned)
            except ValueError:
                return None
        
        return None
    
    @staticmethod
    def _ensure_list(values: Union[Any, List[Any]]) -> List[Any]:
        """
        값을 리스트로 변환합니다.
        
        Args:
            values: 단일 값 또는 리스트
            
        Returns:
            리스트
        """
        if not isinstance(values, list):
            return [values]
        return values
    
    def sum(self, values: Union[Any, List[Any]]) -> float:
        """
        값들의 합계를 계산합니다.
        
        Args:
            values: 단일 값 또는 값들의 리스트
            
        Returns:
            합계
            
        Raises:
            ValueError: 계산 가능한 숫자가 없을 때
        """
        values_list = self._ensure_list(values)
        numeric_values = [self._to_numeric(v) for v in values_list]
        numeric_values = [v for v in numeric_values if v is not None]
        
        if not numeric_values:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        return sum(numeric_values)
    
    def average(self, values: Union[Any, List[Any]]) -> float:
        """
        값들의 평균을 계산합니다.
        
        Args:
            values: 단일 값 또는 값들의 리스트
            
        Returns:
            평균값
            
        Raises:
            ValueError: 계산 가능한 숫자가 없을 때
        """
        values_list = self._ensure_list(values)
        numeric_values = [self._to_numeric(v) for v in values_list]
        numeric_values = [v for v in numeric_values if v is not None]
        
        if not numeric_values:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        return statistics.mean(numeric_values)
    
    def max_value(self, values: Union[Any, List[Any]]) -> float:
        """
        값들 중 최대값을 구합니다.
        
        Args:
            values: 단일 값 또는 값들의 리스트
            
        Returns:
            최대값
            
        Raises:
            ValueError: 계산 가능한 숫자가 없을 때
        """
        values_list = self._ensure_list(values)
        numeric_values = [self._to_numeric(v) for v in values_list]
        numeric_values = [v for v in numeric_values if v is not None]
        
        if not numeric_values:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        return max(numeric_values)
    
    def min_value(self, values: Union[Any, List[Any]]) -> float:
        """
        값들 중 최소값을 구합니다.
        
        Args:
            values: 단일 값 또는 값들의 리스트
            
        Returns:
            최소값
            
        Raises:
            ValueError: 계산 가능한 숫자가 없을 때
        """
        values_list = self._ensure_list(values)
        numeric_values = [self._to_numeric(v) for v in values_list]
        numeric_values = [v for v in numeric_values if v is not None]
        
        if not numeric_values:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        return min(numeric_values)
    
    def growth_rate(self, old_value: Any, new_value: Any, percentage: bool = True) -> float:
        """
        증감률을 계산합니다.
        
        Args:
            old_value: 이전 값 (기준값)
            new_value: 새로운 값
            percentage: True면 퍼센트로 반환, False면 소수로 반환
            
        Returns:
            증감률 (퍼센트 또는 소수)
            
        Raises:
            ValueError: 기준값이 0이거나 계산 불가능한 값일 때
        """
        old_num = self._to_numeric(old_value)
        new_num = self._to_numeric(new_value)
        
        if old_num is None or new_num is None:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        if old_num == 0:
            raise ValueError("기준값이 0이어서 증감률을 계산할 수 없습니다.")
        
        rate = (new_num - old_num) / old_num
        
        if percentage:
            return rate * 100
        
        return rate
    
    def growth_amount(self, old_value: Any, new_value: Any) -> float:
        """
        증감액을 계산합니다.
        
        Args:
            old_value: 이전 값
            new_value: 새로운 값
            
        Returns:
            증감액 (new_value - old_value)
            
        Raises:
            ValueError: 계산 가능한 숫자가 없을 때
        """
        old_num = self._to_numeric(old_value)
        new_num = self._to_numeric(new_value)
        
        if old_num is None or new_num is None:
            raise ValueError("계산 가능한 숫자가 없습니다.")
        
        return new_num - old_num
    
    def calculate(self, operation: str, *args) -> float:
        """
        계산 연산을 수행합니다.
        
        Args:
            operation: 연산 이름 ('sum', 'average', 'max', 'min', 'growth_rate', 'growth_amount')
            *args: 연산에 필요한 인자들
            
        Returns:
            계산 결과
            
        Raises:
            ValueError: 알 수 없는 연산이거나 인자가 잘못되었을 때
        """
        operation = operation.lower().strip()
        
        if operation in ['sum', '합계']:
            if len(args) == 1:
                return self.sum(args[0])
            return self.sum(args)
        
        elif operation in ['average', 'avg', '평균']:
            if len(args) == 1:
                return self.average(args[0])
            return self.average(args)
        
        elif operation in ['max', '최대값', '최대']:
            if len(args) == 1:
                return self.max_value(args[0])
            return self.max_value(args)
        
        elif operation in ['min', '최소값', '최소']:
            if len(args) == 1:
                return self.min_value(args[0])
            return self.min_value(args)
        
        elif operation in ['growth_rate', '증감률', '증가율']:
            if len(args) < 2:
                raise ValueError("증감률 계산에는 이전 값과 새로운 값이 필요합니다.")
            percentage = args[2] if len(args) > 2 else True
            return self.growth_rate(args[0], args[1], percentage)
        
        elif operation in ['growth_amount', '증감액', '증가액']:
            if len(args) < 2:
                raise ValueError("증감액 계산에는 이전 값과 새로운 값이 필요합니다.")
            return self.growth_amount(args[0], args[1])
        
        else:
            raise ValueError(f"알 수 없는 연산: {operation}")
    
    def calculate_from_cell_refs(self, operation: str, values: List[Any]) -> float:
        """
        셀 참조에서 추출된 값들로 계산을 수행합니다.
        
        Args:
            operation: 연산 이름
            values: 셀 값들의 리스트 (단일 값일 수도 있음)
            
        Returns:
            계산 결과
        """
        # 연산이 'growth_rate' 또는 'growth_amount'인 경우
        if operation in ['growth_rate', '증감률', '증가율', 'growth_amount', '증감액', '증가액']:
            if len(values) < 2:
                raise ValueError(f"{operation} 계산에는 최소 2개의 값이 필요합니다.")
            return self.calculate(operation, values[0], values[1])
        
        # 그 외의 경우 (sum, average, max, min 등)
        return self.calculate(operation, values)

