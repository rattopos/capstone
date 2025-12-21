#!/bin/bash

# 웹 애플리케이션 데모 비디오 생성 배쉬 스크립트

# 색상 정의
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo "============================================================"
echo "🎥 웹 애플리케이션 데모 비디오 생성기"
echo "============================================================"
echo ""

# 기본값 설정
DEFAULT_DIR="$HOME/Desktop"

# 날짜와 시간 생성 (YYYY-MM-DD_HH-MM-SS 형식)
TIMESTAMP=$(date +"%Y-%m-%d_%H-%M-%S")
DEFAULT_FILENAME="demo_video_${TIMESTAMP}.mp4"

# 저장 위치 입력 받기
echo -e "${BLUE}📁 저장 위치를 입력하세요 (Enter: 기본값 '~/Desktop')${NC}"
read -p "저장 위치: " OUTPUT_DIR

# 저장 위치가 비어있으면 기본값 사용
if [ -z "$OUTPUT_DIR" ]; then
    OUTPUT_DIR="$DEFAULT_DIR"
fi

# ~를 홈 디렉토리로 확장
OUTPUT_DIR="${OUTPUT_DIR/#\~/$HOME}"

# 디렉토리 생성
mkdir -p "$OUTPUT_DIR"

# 파일명 입력 받기
echo ""
echo -e "${BLUE}📝 파일명을 입력하세요 (Enter: 기본값 '${DEFAULT_FILENAME}')${NC}"
read -p "파일명: " FILENAME

# 파일명이 비어있으면 기본값 사용
if [ -z "$FILENAME" ]; then
    FILENAME="$DEFAULT_FILENAME"
fi

# 파일명에 확장자가 없으면 .mp4 추가
if [[ ! "$FILENAME" =~ \.(mp4|webm|avi|mov)$ ]]; then
    FILENAME="${FILENAME}.mp4"
fi

# 전체 경로 생성
if [ "$OUTPUT_DIR" = "." ] || [ "$OUTPUT_DIR" = "./" ]; then
    OUTPUT_PATH="$FILENAME"
else
    OUTPUT_PATH="${OUTPUT_DIR}/${FILENAME}"
fi

echo ""
echo -e "${GREEN}✅ 설정 완료:${NC}"
echo "   저장 위치: $OUTPUT_DIR"
echo "   파일명: $FILENAME"
echo "   전체 경로: $OUTPUT_PATH"
echo ""

# 고급 데모 옵션 확인
echo -e "${YELLOW}고급 데모를 실행하시겠습니까? (여러 템플릿 테스트)${NC}"
read -p "고급 데모 실행? (y/N): " ADVANCED

ADVANCED_FLAG=""
if [[ "$ADVANCED" =~ ^[Yy]$ ]]; then
    ADVANCED_FLAG="--advanced"
    echo -e "${GREEN}고급 데모 모드로 실행합니다.${NC}"
else
    echo -e "${GREEN}기본 데모 모드로 실행합니다.${NC}"
fi

echo ""
echo "============================================================"
echo "🚀 데모 비디오 생성을 시작합니다..."
echo "============================================================"
echo ""

# Python 스크립트 실행
python3 create_demo_video.py $ADVANCED_FLAG --output "$OUTPUT_PATH"

# 실행 결과 확인
if [ $? -eq 0 ]; then
    echo ""
    echo "============================================================"
    echo -e "${GREEN}✨ 데모 비디오 생성이 완료되었습니다!${NC}"
    echo "============================================================"
    echo ""
    echo -e "${BLUE}생성된 파일:${NC} $OUTPUT_PATH"
    
    # 파일 크기 표시 (파일이 존재하는 경우)
    if [ -f "$OUTPUT_PATH" ]; then
        FILE_SIZE=$(du -h "$OUTPUT_PATH" | cut -f1)
        echo -e "${BLUE}파일 크기:${NC} $FILE_SIZE"
    fi
else
    echo ""
    echo "============================================================"
    echo -e "${YELLOW}⚠️  데모 비디오 생성 중 오류가 발생했습니다.${NC}"
    echo "============================================================"
    exit 1
fi

