# hwp2hwpx

HWP → HWPX 변환 라이브러리 (Python)

한글과컴퓨터(한컴)의 워드프로세서 "한글"의 `.hwp` (바이너리 OLE2) 파일을 `.hwpx` (ZIP/XML) 파일로 변환합니다.

> 원본 Java 구현: [neolord0/hwp2hwpx](https://github.com/neolord0/hwp2hwpx) — 이 Python 버전은 Java 의존성 없이 동일한 변환을 수행합니다.

## 설치 (Installation)

```bash
pip install pyhwp olefile lxml
```

## 사용법 (Usage)

### CLI

```bash
python3 -m hwp2hwpx input.hwp                    # → input.hwpx
python3 -m hwp2hwpx input.hwp -o output.hwpx      # 출력 경로 지정
python3 -m hwp2hwpx input.hwp -v                   # 상세 출력
```

### Python API

```python
from hwp2hwpx import convert_file

# 파일 변환
output_path = convert_file("input.hwp", "output.hwpx")

# 또는 바이트 딕셔너리로 변환
from hwp2hwpx import convert
from hwp2hwpx.reader import HWPReader

with HWPReader("input.hwp") as reader:
    files = convert(reader)  # dict: filepath → bytes
```

## 아키텍처 (Architecture)

```
hwp2hwpx/
├── __init__.py          # 패키지 진입점
├── __main__.py          # CLI (argparse)
├── converter.py         # HWPX ZIP 조립 (mimetype, OPF, manifest, sections)
├── reader.py            # HWP 파일 읽기 (pyhwp + olefile 래퍼)
├── header_converter.py  # DocInfo → header.xml (글꼴, 문자속성, 문단속성, 스타일)
├── section_converter.py # BodyText → section[N].xml (문단, 표, 컬럼, 구역속성)
├── xml_builder.py       # lxml 유틸리티, HWPX 네임스페이스 정의
└── value_maps.py        # HWP 바이너리 값 → HWPX XML 열거형 변환 테이블
```

### 변환 흐름

1. `HWPReader`가 pyhwp의 xmlmodel API로 HWP 레코드 스트림을 파싱
2. `header_converter`가 DocInfo 레코드를 `header.xml`로 변환 (글꼴, 테두리, 문자속성, 문단속성, 스타일)
3. `section_converter`가 각 섹션의 레코드를 `section[N].xml`로 변환 (문단 → run → 텍스트/컨트롤)
4. `converter`가 모든 파일을 HWPX ZIP으로 패키징 (mimetype, OPF, manifest, header, sections, bindata, preview)

## 테스트 현황 (Test Status)

**41/41 파일 변환 성공** (33 테스트 케이스 + 8 실제 HWP 파일)

### 지원 기능
- 문단 텍스트 및 문자 속성 (charShape)
- 문단 속성 (paraShape, 정렬, 간격, 들여쓰기)
- 표 (table, 셀 병합, 테두리, 배경)
- 다단 (column definition)
- 구역 속성 (section properties, 용지 크기, 여백)
- 글꼴 (fontface), 스타일, 탭, 번호 매기기
- 바이너리 데이터 (이미지 파일 포함)
- 미리보기 텍스트/이미지

### 미구현 (TODO)
- GSO (도형, 이미지, 그리기 개체, OLE)
- 머리글/바닥글 내용
- 각주/미주 내용
- 필드 (하이퍼링크, 책갈피, 쪽 번호)
- 수식

## 참고 자료

- [한글 문서 파일 구조 5.0](http://www.hancom.com/etc/hwpDownload.do?gnb0=269&gnb1=271)
  > "본 제품은 한글과컴퓨터의 HWP 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다."
- [OWPML 문서](http://www.hancom.com/etc/hwpDownload.do?gnb0=269&gnb1=271)
- 원본 Java 라이브러리: [hwplib](https://github.com/neolord0/hwplib), [hwpxlib](https://github.com/neolord0/hwpxlib)
- Python HWP 라이브러리: [pyhwp](https://github.com/mete0r/pyhwp)

## 라이선스 (License)

Apache-2.0 (원본 프로젝트와 동일)

## 변경 이력 (Changelog)

### 2026-03-30 — Python 리팩터링
- Java 의존성 제거, 순수 Python 구현
- pyhwp + olefile + lxml 기반
- 41개 테스트 파일 100% 변환 성공

### 2026-01-20
- pull-request 7: hwp의 ShadowInfo 객체가 null일 때 오류 해결

### 2025-11-14
- 이슈 3: 표 셀의 배경색 설정 오류 해결

### 2025-03-10
- 이슈 1: 포함된 이미지파일의 이름을 생성하는 루틴 변경
- hwpxlib 1.0.5 버전으로 변경
