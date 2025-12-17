"""
Word 문서를 PDF로 변환하는 모듈
"""

import os
import subprocess
import sys
from pathlib import Path
from typing import Optional


class WordToPDFConverter:
    """Word 문서를 PDF로 변환하는 클래스"""
    
    def __init__(self):
        """Word to PDF 변환기 초기화"""
        self.libreoffice_available = self._check_libreoffice()
        self.docx2pdf_available = self._check_docx2pdf()
    
    def _check_libreoffice(self) -> bool:
        """LibreOffice가 설치되어 있는지 확인"""
        try:
            result = subprocess.run(
                ['soffice', '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            return result.returncode == 0
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return False
    
    def _check_docx2pdf(self) -> bool:
        """docx2pdf가 사용 가능한지 확인"""
        try:
            import docx2pdf
            return True
        except ImportError:
            return False
    
    def convert_word_to_pdf(self, word_path: str, output_path: Optional[str] = None) -> str:
        """
        Word 문서를 PDF로 변환합니다.
        
        Args:
            word_path: Word 파일 경로 (.docx)
            output_path: 출력 PDF 파일 경로 (None이면 자동 생성)
            
        Returns:
            생성된 PDF 파일 경로
        """
        word_file = Path(word_path)
        if not word_file.exists():
            raise FileNotFoundError(f"Word 파일을 찾을 수 없습니다: {word_path}")
        
        if output_path is None:
            output_path = str(word_file.with_suffix('.pdf'))
        
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        # LibreOffice 사용 (Mac/Linux)
        if self.libreoffice_available:
            return self._convert_with_libreoffice(word_path, str(output_file))
        
        # docx2pdf 사용 (Windows 또는 대체)
        elif self.docx2pdf_available:
            return self._convert_with_docx2pdf(word_path, str(output_file))
        
        else:
            raise RuntimeError(
                "PDF 변환 도구가 설치되어 있지 않습니다.\n"
                "다음 중 하나를 설치해주세요:\n"
                "1. LibreOffice (Mac/Linux): brew install --cask libreoffice\n"
                "2. docx2pdf (Windows): pip install docx2pdf"
            )
    
    def _convert_with_libreoffice(self, word_path: str, output_path: str) -> str:
        """LibreOffice를 사용하여 변환"""
        word_file = Path(word_path)
        output_file = Path(output_path)
        output_dir = output_file.parent
        
        try:
            # LibreOffice를 headless 모드로 실행하여 변환
            cmd = [
                'soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(output_dir),
                str(word_file)
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice 변환 실패: {result.stderr}")
            
            # LibreOffice는 원본 파일명을 사용하여 PDF 생성
            generated_pdf = output_dir / word_file.with_suffix('.pdf').name
            
            if generated_pdf.exists():
                # 원하는 경로로 이동
                if generated_pdf != output_file:
                    generated_pdf.rename(output_file)
                return str(output_file)
            else:
                raise RuntimeError(f"PDF 파일이 생성되지 않았습니다: {generated_pdf}")
                
        except subprocess.TimeoutExpired:
            raise RuntimeError("PDF 변환 시간 초과")
        except Exception as e:
            raise RuntimeError(f"PDF 변환 중 오류 발생: {str(e)}")
    
    def _convert_with_docx2pdf(self, word_path: str, output_path: str) -> str:
        """docx2pdf를 사용하여 변환"""
        try:
            import docx2pdf
            docx2pdf.convert(word_path, output_path)
            return output_path
        except Exception as e:
            raise RuntimeError(f"docx2pdf 변환 중 오류 발생: {str(e)}")

