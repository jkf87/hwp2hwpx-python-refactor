# hwp2hwpx-python-refactor

`.hwp` 파일을 `.hwpx`로 변환하는 **순수 Python 기반 변환기**입니다.

순수 Python 구현으로 `.hwp`를 `.hwpx`로 변환합니다.

## 설치

```bash
pip install -r requirements.txt
```

## 사용법

### 기본 변환

```bash
python3 -m hwp2hwpx input.hwp
```

같은 폴더에 `input.hwpx`가 생성됩니다.

### 출력 파일명 지정

```bash
python3 -m hwp2hwpx input.hwp -o output.hwpx
```

### 상세 로그 출력

```bash
python3 -m hwp2hwpx input.hwp -v
```

## Python 코드에서 사용

```python
from hwp2hwpx import convert_file

convert_file("input.hwp", "output.hwpx")
```

## 참고

- 입력: `.hwp`
- 출력: `.hwpx`
- 주요 의존성: `pyhwp`, `olefile`, `lxml`

## 라이선스

라이선스 검토 중
