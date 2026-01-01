# 📚 프로젝트 용어 사전

> 지역경제동향 보고서 자동 생성 시스템에서 사용되는 용어들을 정리한 사전입니다.

---

## 1. 경제/통계 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **광공업생산지수** (IIP) | 광업과 제조업의 생산 활동 수준을 나타내는 지수. 기준 연도(2020년=100)를 100으로 설정하고 생산량의 변화를 측정한다. | [한국은행 경제용어](https://www.bok.or.kr/portal/ecEdu/ecWordDicary/search.do?menuNo=200688) |
| **서비스업생산지수** | 서비스업 분야의 생산 활동 수준을 측정하는 지수. 도소매, 운수, 숙박음식, 정보통신 등의 서비스업 생산량 변화를 나타낸다. | [e-나라지표](https://www.index.go.kr/unity/potal/indicator/IndexInfo.do?clasCd=10&idxCd=F0053) |
| **소매판매액지수** | 소매업체의 판매액 변화를 측정하는 지수. 소비 동향을 파악하는 대표적인 지표로, 가계의 소비 지출 수준을 반영한다. | [국가통계포털 KOSIS](https://kosis.kr/statHtml/statHtml.do?orgId=101&tblId=DT_1KE1001) |
| **GRDP** (지역내총생산) | Gross Regional Domestic Product의 약자. 일정 기간 동안 특정 지역 내에서 생산된 모든 최종 재화와 서비스의 시장 가치 합계. GDP의 지역 버전이다. | [위키백과 - 지역내총생산](https://ko.wikipedia.org/wiki/지역내총생산) |
| **고용률** | 15세 이상 생산가능인구 중 취업자가 차지하는 비율. 수식: (취업자 수 / 생산가능인구) × 100 | [위키백과 - 고용률](https://ko.wikipedia.org/wiki/고용률) |
| **실업률** | 경제활동인구 중 실업자가 차지하는 비율. 수식: (실업자 수 / 경제활동인구) × 100 | [위키백과 - 실업률](https://ko.wikipedia.org/wiki/실업률) |
| **소비자물가지수** (CPI) | Consumer Price Index의 약자. 가계가 소비하는 상품과 서비스의 가격 변동을 측정하는 지수. 인플레이션 측정의 대표적 지표이다. | [위키백과 - 소비자물가지수](https://ko.wikipedia.org/wiki/소비자_물가_지수) |
| **건설수주액** | 건설업체가 발주자로부터 수주한 공사 금액의 합계. 건설업 경기 전망을 나타내는 선행지표로 활용된다. | [e-나라지표](https://www.index.go.kr/unity/potal/indicator/IndexInfo.do?clasCd=10&idxCd=4201) |
| **수출액** | 국내에서 생산된 재화가 국경을 넘어 외국으로 판매된 금액의 합계. 대외 무역 수지를 파악하는 핵심 지표이다. | [한국무역협회](https://stat.kita.net/) |
| **수입액** | 외국에서 생산된 재화가 국내로 반입된 금액의 합계. 내수 경기 및 산업 생산에 대한 의존도를 파악할 수 있다. | [한국무역협회](https://stat.kita.net/) |
| **국내인구이동** | 국내 시·도 간 주민등록 주소지 이동을 의미. 순이동자수는 전입자에서 전출자를 뺀 값이다. | [국가통계포털 KOSIS](https://kosis.kr/statHtml/statHtml.do?orgId=101&tblId=DT_1B26001_A01) |
| **전년동기비** (YoY) | Year over Year의 약자. 당해 분기를 전년 같은 분기와 비교한 것. 예: 2025년 2분기 vs 2024년 2분기 | [Investopedia - YoY](https://www.investopedia.com/terms/y/year-over-year.asp) |
| **전분기비** (QoQ) | Quarter over Quarter의 약자. 당해 분기를 직전 분기와 비교한 것. 예: 2025년 2분기 vs 2025년 1분기 | [Investopedia - QoQ](https://www.investopedia.com/terms/q/quarter-over-quarter.asp) |
| **증감률** | 두 시점 간의 값 변화를 백분율로 나타낸 것. 수식: ((당기값 - 기준값) / 기준값) × 100 | [위키백과 - 변화율](https://ko.wikipedia.org/wiki/변화율) |
| **기여도** | 전체 성장률에서 특정 항목이 기여한 정도를 나타내는 값. 수식: ((당기값 - 전기값) / 전체 전기값) × 100 | [한국은행 경제통계시스템](https://ecos.bok.or.kr/) |
| **지수** (Index) | 기준 시점의 값을 100으로 설정하고, 비교 시점의 값을 상대적인 비율로 나타낸 수치. 시계열 비교에 유용하다. | [위키백과 - 경제지수](https://ko.wikipedia.org/wiki/경제_지수) |
| **결측치** | 데이터셋에서 누락되거나 기록되지 않은 값. Missing Value, NaN(Not a Number) 등으로 표현된다. | [위키백과 - 결측값](https://ko.wikipedia.org/wiki/결측값) |
| **이상치** (Outlier) | 데이터의 일반적인 패턴에서 크게 벗어난 값. 데이터 품질 검증 시 주의가 필요하다. | [위키백과 - 이상값](https://ko.wikipedia.org/wiki/이상값) |

---

## 2. 기술/프로그래밍 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **Flask** | Python으로 작성된 마이크로 웹 프레임워크. 간단하고 유연한 구조로 웹 애플리케이션을 빠르게 개발할 수 있다. | [Flask 공식 문서](https://flask.palletsprojects.com/) |
| **마이크로 프레임워크** (Micro Framework) | 최소한의 핵심 기능만 제공하는 가벼운 웹 프레임워크. Flask, Bottle 등이 있다. 필요한 기능은 확장(extension)으로 추가한다. | [위키백과 - 마이크로 프레임워크](https://en.wikipedia.org/wiki/Microframework) |
| **풀 스택 프레임워크** (Full Stack Framework) | 웹 개발에 필요한 모든 기능(ORM, 인증, 관리자 페이지 등)을 내장한 프레임워크. Django, Ruby on Rails 등이 있다. | [위키백과 - 웹 프레임워크](https://ko.wikipedia.org/wiki/웹_프레임워크) |
| **Django** | Python으로 작성된 풀 스택 웹 프레임워크. ORM, 관리자 페이지, 인증 시스템 등이 내장되어 있어 대규모 애플리케이션에 적합하다. | [Django 공식 문서](https://www.djangoproject.com/) |
| **FastAPI** | Python으로 작성된 현대적인 웹 프레임워크. 높은 성능과 자동 API 문서화 기능으로 API 개발에 특화되어 있다. | [FastAPI 공식 문서](https://fastapi.tiangolo.com/) |
| **WSGI** | Web Server Gateway Interface의 약자. Python 웹 애플리케이션과 웹 서버 간의 표준 인터페이스. Flask, Django는 WSGI를 따르므로 다양한 서버에서 실행 가능하다. | [위키백과 - WSGI](https://ko.wikipedia.org/wiki/웹_서버_게이트웨이_인터페이스) |
| **유연성** (Flexibility) | 시스템이나 코드가 다양한 방식으로 사용되거나 확장될 수 있는 정도. Flask는 필요한 기능만 선택적으로 추가할 수 있어 유연성이 높다. | - |
| **확장성** (Extensibility) | 시스템에 새로운 기능을 쉽게 추가할 수 있는 정도. Flask의 extension 시스템은 다양한 기능을 플러그인 형태로 추가할 수 있게 한다. | - |
| **가벼움** (Lightweight) | 시스템이 작고 간단하여 빠르게 실행되고 리소스를 적게 사용하는 특성. Flask는 핵심 기능만 포함하여 가볍다. | - |
| **학습 곡선** (Learning Curve) | 새로운 기술이나 도구를 익히는 데 걸리는 시간과 난이도. Flask는 간단한 구조로 학습 곡선이 낮다. | [위키백과 - 학습 곡선](https://en.wikipedia.org/wiki/Learning_curve) |
| **커뮤니티** (Community) | 특정 기술이나 프로젝트를 사용하는 개발자들의 집단. 활발한 커뮤니티는 문서, 예제, 도움을 제공한다. | - |
| **문서화** (Documentation) | 코드나 API의 사용법을 설명한 문서. Flask 확장, API 엔드포인트 등의 문서화는 개발 속도와 코드 이해도를 향상시킨다. | - |
| **의존성 관리** (Dependency Management) | 프로젝트가 사용하는 외부 라이브러리와 패키지를 관리하는 것. requirements.txt나 Pipfile로 관리한다. | [pip 문서](https://pip.pypa.io/) |
| **라우팅** (Routing) | URL 경로를 특정 함수나 핸들러에 매핑하는 과정. Flask에서는 @app.route() 데코레이터로 간단하게 정의할 수 있다. | [Flask 라우팅](https://flask.palletsprojects.com/en/latest/quickstart/#routing) |
| **뷰 함수** (View Function) | 웹 요청을 처리하고 응답을 반환하는 함수. Flask에서는 각 라우트에 연결된 함수를 뷰 함수라고 한다. | [Flask 뷰 함수](https://flask.palletsprojects.com/en/latest/quickstart/#http-methods) |
| **데코레이터 패턴** (Decorator Pattern) | 함수나 메서드의 동작을 수정하지 않고 기능을 추가하는 디자인 패턴. Flask의 @app.route()가 대표적인 예이다. | [위키백과 - 데코레이터 패턴](https://ko.wikipedia.org/wiki/데코레이터_패턴) |
| **컨벤션 오버 구성** (Convention over Configuration) | 설정보다는 관례(convention)를 따르면 자동으로 작동하는 방식. Django는 이 방식, Flask는 명시적 구성(explicit configuration)을 선호한다. | [위키백과 - CoC](https://en.wikipedia.org/wiki/Convention_over_configuration) |
| **명시적 구성** (Explicit Configuration) | 모든 설정을 명시적으로 작성하는 방식. Flask는 이 방식을 채택하여 코드가 명확하고 이해하기 쉽다. | - |
| **개발 서버** (Development Server) | 개발 중에 사용하는 간단한 웹 서버. Flask는 내장 개발 서버를 제공하여 빠르게 테스트할 수 있다. | [Flask 개발 서버](https://flask.palletsprojects.com/en/latest/quickstart/#a-minimal-application) |
| **프로덕션 서버** (Production Server) | 실제 사용자를 대상으로 서비스를 제공하는 서버. Gunicorn, uWSGI 등이 Flask 애플리케이션을 실행한다. | [Flask 배포](https://flask.palletsprojects.com/en/latest/deploying/) |
| **모듈화** (Modularity) | 시스템을 독립적인 모듈로 나누는 설계 방식. Flask의 Blueprint를 사용하면 애플리케이션을 모듈화할 수 있다. | [위키백과 - 모듈화](https://ko.wikipedia.org/wiki/모듈화) |
| **프로토타이핑** (Prototyping) | 빠르게 동작하는 초기 버전을 만드는 과정. Flask는 간단한 구조로 빠른 프로토타이핑이 가능하다. | [위키백과 - 프로토타입](https://ko.wikipedia.org/wiki/프로토타입) |
| **템플릿 엔진** (Template Engine) | 정적 HTML에 동적 데이터를 삽입하여 렌더링하는 도구. Flask는 기본적으로 Jinja2 템플릿 엔진을 사용한다. | [위키백과 - 템플릿 엔진](https://en.wikipedia.org/wiki/Template_processor) |
| **Jinja2** | Python용 템플릿 엔진. HTML에 Python 변수와 로직을 삽입하여 동적 웹 페이지를 생성할 수 있다. Flask의 기본 템플릿 엔진이다. | [Jinja2 공식 문서](https://jinja.palletsprojects.com/) |
| **플러그인** (Plugin) | 기존 시스템에 기능을 추가하는 확장 모듈. Flask extension은 플러그인 방식으로 기능을 추가한다. | [위키백과 - 플러그인](https://ko.wikipedia.org/wiki/플러그인) |
| **확장 라이브러리** (Extension Library) | 프레임워크의 기능을 확장하는 외부 라이브러리. Flask-SQLAlchemy, Flask-Login 등이 Flask 확장 라이브러리이다. | [Flask 확장](https://flask.palletsprojects.com/en/latest/extensions/) |
| **보일러플레이트 코드** (Boilerplate Code) | 반복적으로 작성해야 하는 상용구 코드. Flask는 최소한의 보일러플레이트로 시작할 수 있다. | [위키백과 - 보일러플레이트](https://en.wikipedia.org/wiki/Boilerplate_code) |
| **즉시 사용 가능** (Out of the Box) | 추가 설정이나 설치 없이 바로 사용할 수 있는 상태. Flask는 최소한의 코드로 바로 실행할 수 있다. | - |
| **pandas** | Python 데이터 분석 라이브러리. DataFrame이라는 자료구조를 사용하여 표 형태의 데이터를 효율적으로 처리한다. | [pandas 공식 문서](https://pandas.pydata.org/) |
| **openpyxl** | Python에서 Excel 파일(.xlsx)을 읽고 쓰는 라이브러리. 수식, 서식, 차트 등을 처리할 수 있다. | [openpyxl 공식 문서](https://openpyxl.readthedocs.io/) |
| **numpy** | Python의 과학 계산용 라이브러리. 다차원 배열과 행렬 연산, 수학 함수를 제공한다. | [NumPy 공식 문서](https://numpy.org/) |
| **Blueprint** | Flask에서 애플리케이션을 모듈화하는 패턴. 라우트와 뷰 함수를 그룹화하여 코드를 체계적으로 관리할 수 있다. | [Flask Blueprint 문서](https://flask.palletsprojects.com/en/latest/blueprints/) |
| **API** | Application Programming Interface의 약자. 서로 다른 소프트웨어 간 통신을 위한 인터페이스 규약이다. | [위키백과 - API](https://ko.wikipedia.org/wiki/API) |
| **REST API** | Representational State Transfer API. HTTP 프로토콜을 사용하여 자원을 CRUD(생성, 조회, 수정, 삭제)하는 웹 서비스 설계 방식이다. | [위키백과 - REST](https://ko.wikipedia.org/wiki/REST) |
| **JSON** | JavaScript Object Notation의 약자. 키-값 쌍으로 구성된 경량 데이터 교환 포맷. 사람과 기계 모두 읽기 쉽다. | [JSON 공식 사이트](https://www.json.org/json-ko.html) |
| **DataFrame** | pandas의 2차원 테이블 형태 자료구조. 행과 열로 구성되며, 엑셀의 시트와 유사한 개념이다. | [pandas DataFrame 문서](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.html) |
| **KOSIS** | Korean Statistical Information Service의 약자. 국가통계포털로, 국가데이터처에서 제공하는 각종 통계 데이터를 열람할 수 있다. | [KOSIS 포털](https://kosis.kr/) |
| **KOSIS Open API** | KOSIS에서 제공하는 공개 API. 프로그래밍 방식으로 통계 데이터를 자동 수집할 수 있다. | [KOSIS Open API 가이드](https://kosis.kr/openapi/index.do) |
| **CLI** | Command Line Interface의 약자. 텍스트 기반의 명령어 입력 방식으로 프로그램을 실행하는 인터페이스이다. | [위키백과 - CLI](https://ko.wikipedia.org/wiki/명령_줄_인터페이스) |
| **pip** | Python 패키지 관리자. PyPI(Python Package Index)에서 패키지를 설치, 업데이트, 삭제할 수 있다. | [pip 공식 문서](https://pip.pypa.io/) |
| **requirements.txt** | Python 프로젝트의 의존성 패키지 목록을 기록한 파일. `pip install -r requirements.txt`로 일괄 설치할 수 있다. | [pip requirements 문서](https://pip.pypa.io/en/stable/reference/requirements-file-format/) |
| **HTML** | HyperText Markup Language의 약자. 웹 페이지의 구조와 내용을 정의하는 마크업 언어이다. | [MDN HTML 문서](https://developer.mozilla.org/ko/docs/Web/HTML) |
| **CSS** | Cascading Style Sheets의 약자. HTML 요소의 스타일(색상, 레이아웃, 폰트 등)을 정의하는 언어이다. | [MDN CSS 문서](https://developer.mozilla.org/ko/docs/Web/CSS) |
| **JavaScript** | 웹 브라우저에서 실행되는 스크립트 언어. 동적인 웹 페이지 기능(버튼 클릭, 데이터 요청 등)을 구현한다. | [MDN JavaScript 문서](https://developer.mozilla.org/ko/docs/Web/JavaScript) |
| **iframe** | Inline Frame의 약자. HTML 문서 안에 다른 HTML 문서를 삽입하는 요소. 미리보기 기능에서 활용된다. | [MDN iframe 문서](https://developer.mozilla.org/ko/docs/Web/HTML/Element/iframe) |
| **fetch API** | JavaScript에서 HTTP 요청을 보내는 내장 API. 서버와 비동기 통신을 할 때 사용한다. | [MDN Fetch API 문서](https://developer.mozilla.org/ko/docs/Web/API/Fetch_API) |
| **드래그 앤 드롭** | 마우스로 파일이나 요소를 끌어서 특정 위치에 놓는 인터랙션 방식. 파일 업로드에 자주 사용된다. | [MDN Drag and Drop](https://developer.mozilla.org/ko/docs/Web/API/HTML_Drag_and_Drop_API) |
| **sanitization** (입력값 정제) | 사용자 입력에서 악의적인 코드나 특수 문자를 제거하거나 이스케이프 처리하는 과정. 보안 취약점을 방지한다. | [OWASP - Input Validation](https://owasp.org/www-community/Improper_Input_Validation) |
| **middleware** (미들웨어) | 애플리케이션과 시스템 사이에서 요청과 응답을 가로채서 처리하는 소프트웨어 계층. Flask에서는 요청 전/후 처리, 인증, 로깅 등에 사용된다. | [Flask 미들웨어](https://flask.palletsprojects.com/en/latest/patterns/middleware/) |
| **async/await** | 비동기 프로그래밍을 위한 키워드. async는 비동기 함수를 선언하고, await는 비동기 작업이 완료될 때까지 대기한다. | [MDN async/await](https://developer.mozilla.org/ko/docs/Web/JavaScript/Reference/Statements/async_function) |
| **decorator** (데코레이터) | 함수나 클래스를 수정하지 않고 기능을 추가하는 Python 기능. @ 기호로 사용하며, Flask의 @app.route() 같은 라우트 정의에 활용된다. | [Python 데코레이터](https://docs.python.org/ko/3/glossary.html#term-decorator) |
| **dependency injection** (의존성 주입) | 객체가 필요로 하는 의존성을 외부에서 주입하는 디자인 패턴. 코드의 결합도를 낮추고 테스트 가능성을 높인다. | [위키백과 - 의존성 주입](https://ko.wikipedia.org/wiki/의존성_주입) |

---

## 3. 데이터/파일 관련 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **기초자료 수집표** | KOSIS 등에서 수집한 원본 통계 데이터를 담은 엑셀 파일. 11개 시트(광공업생산, 서비스업생산 등)로 구성된다. | 프로젝트 내부 문서 |
| **분석표** | 기초자료를 가공하여 집계 및 분석 수식이 적용된 엑셀 파일. 42개 시트로 구성되며, 보고서 생성의 입력 자료이다. | 프로젝트 내부 문서 |
| **템플릿** (문서 템플릿) | 반복적으로 사용되는 문서의 틀. 이 프로젝트에서는 Jinja2 HTML 템플릿으로 보고서 형식을 정의한다. (코드 템플릿과 구분됨) | [위키백과 - 템플릿](https://ko.wikipedia.org/wiki/템플릿) |
| **스키마** (Schema, 데이터 스키마) | 데이터의 구조와 형식을 정의한 명세서. JSON 스키마 파일은 템플릿에서 사용할 변수 구조를 정의한다. | [JSON Schema](https://json-schema.org/) |
| **집계 시트** | 분석표에서 기초자료의 값을 정리한 시트. 시트명에 '집계'가 포함된다. (예: A(광공업생산)집계) | 프로젝트 내부 문서 |
| **분석 시트** | 분석표에서 집계 시트의 데이터를 참조하여 증감률 등을 계산하는 시트. 시트명에 '분석'이 포함된다. | 프로젝트 내부 문서 |
| **엑셀 수식** | 셀에서 계산을 수행하는 표현식. =으로 시작하며, 셀 참조와 함수를 사용할 수 있다. (예: =SUM(A1:A10)) | [Microsoft Excel 수식](https://support.microsoft.com/ko-kr/office/수식-개요) |
| **셀 병합** | 여러 셀을 하나로 합치는 기능. Merged Cell은 데이터 처리 시 특별한 처리가 필요하다. | [Excel 셀 병합](https://support.microsoft.com/ko-kr/office/셀-병합-및-병합-취소) |
| **스키마 검증** (Schema Validation) | 데이터가 JSON Schema나 XML Schema 등으로 정의된 구조와 형식에 맞는지 확인하는 검증 방법. | [JSON Schema 검증](https://json-schema.org/understanding-json-schema/) |
| **타입 검사** (Type Checking) | 변수나 데이터의 자료형이 예상한 타입과 일치하는지 확인하는 과정. Python에서는 isinstance() 등으로 수행한다. | [Python 타입 힌트](https://docs.python.org/ko/3/library/typing.html) |
| **제약 조건** (Constraint) | 데이터에 적용되는 규칙이나 제한사항. 예: 최대값, 최소값, 필수 여부, 형식 패턴 등. | [위키백과 - 데이터베이스 제약조건](https://ko.wikipedia.org/wiki/데이터베이스_제약조건) |
| **형식 검증** (Format Validation) | 데이터가 특정 형식(예: 이메일 주소, 전화번호, 날짜 형식)에 맞는지 확인하는 검증. 정규표현식으로 구현할 수 있다. | [MDN 정규표현식](https://developer.mozilla.org/ko/docs/Web/JavaScript/Guide/Regular_Expressions) |
| **범위 검증** (Range Validation) | 데이터 값이 허용된 최소값과 최대값 사이에 있는지 확인하는 검증. 숫자, 날짜, 문자열 길이 등에 적용한다. | - |
| **필수 필드** (Required Field) | 반드시 입력되어야 하는 데이터 필드. 빈 값(null, 빈 문자열)을 허용하지 않는다. | [MDN required 속성](https://developer.mozilla.org/ko/docs/Web/HTML/Attributes/required) |
| **선택 필드** (Optional Field) | 입력하지 않아도 되는 데이터 필드. 값이 없어도 유효한 것으로 간주한다. | - |
| **데이터 정규화** (Data Normalization) | 데이터를 일관된 형식으로 변환하는 과정. 예: 공백 제거, 대소문자 통일, 날짜 형식 통일 등. | [위키백과 - 정규화](https://ko.wikipedia.org/wiki/정규화) |
| **데이터 무결성** (Data Integrity) | 데이터가 정확하고 일관되며 신뢰할 수 있는 상태를 유지하는 것. 중복 제거, 참조 무결성 등을 포함한다. | [위키백과 - 데이터 무결성](https://ko.wikipedia.org/wiki/데이터_무결성) |
| **데이터 품질** (Data Quality) | 데이터가 목적에 적합한 정도를 나타내는 지표. 정확성, 완전성, 일관성, 적시성 등을 평가한다. | [위키백과 - 데이터 품질](https://en.wikipedia.org/wiki/Data_quality) |
| **이상치 탐지** (Outlier Detection) | 데이터셋에서 다른 데이터와 크게 다른 값을 찾아내는 과정. 통계적 방법이나 머신러닝으로 수행할 수 있다. | [위키백과 - 이상값](https://ko.wikipedia.org/wiki/이상값) |
| **데이터 정제** (Data Cleansing) | 데이터셋에서 오류, 중복, 불일치를 제거하거나 수정하는 과정. 데이터 전처리의 핵심 단계이다. | [위키백과 - 데이터 정제](https://en.wikipedia.org/wiki/Data_cleansing) |
| **데이터 프로파일링** (Data Profiling) | 데이터셋의 구조, 통계적 특성, 품질을 분석하여 요약 정보를 생성하는 과정. 데이터 검증 전에 수행된다. | [위키백과 - 데이터 프로파일링](https://en.wikipedia.org/wiki/Data_profiling) |
| **ETL** | Extract(추출), Transform(변환), Load(적재)의 약자. 여러 소스에서 데이터를 추출하고, 검증 및 변환한 후 목적지에 적재하는 프로세스. | [위키백과 - ETL](https://ko.wikipedia.org/wiki/ETL) |
| **체크섬** (Checksum) | 데이터의 무결성을 검증하기 위해 계산된 값. 데이터가 전송 또는 저장 중 변경되었는지 확인할 수 있다. | [위키백과 - 체크섬](https://ko.wikipedia.org/wiki/체크섬) |
| **데이터 검증 규칙** (Validation Rule) | 데이터가 유효한지 판단하기 위한 명시적인 규칙 집합. 비즈니스 로직에 따라 정의된다. | - |
| **유효성 검사** (Verification) | 데이터가 요구사항이나 명세를 만족하는지 확인하는 과정. 검증(validation)과 유사하지만 더 공식적인 의미를 가진다. | [위키백과 - 검증](https://ko.wikipedia.org/wiki/검증) |
| **데이터 타입** (Data Type) | 데이터의 종류를 나타내는 분류. Python에는 int, float, str, bool, list, dict 등이 있다. | [Python 데이터 타입](https://docs.python.org/ko/3/library/stdtypes.html) |
| **형변환** (Type Casting) | 한 데이터 타입을 다른 데이터 타입으로 변환하는 과정. 예: 문자열 "123"을 정수 123으로 변환. | [Python 형변환](https://docs.python.org/ko/3/library/functions.html#int) |
| **null 체크** | 데이터가 null(없음) 값인지 확인하는 검증. 데이터베이스나 프로그래밍에서 None, NULL, NaN 등을 체크한다. | [Python None 체크](https://docs.python.org/ko/3/library/stdtypes.html#the-null-object) |
| **중복 검사** (Duplicate Check) | 동일한 데이터가 이미 존재하는지 확인하는 검증. 고유성 제약 조건이나 기본키 검증에 사용된다. | - |
| **크로스 검증** (Cross Validation) | 여러 데이터 소스나 필드를 비교하여 일관성을 확인하는 검증 방법. 데이터 간 관계를 검증할 때 사용한다. | - |

---

## 4. 보고서 관련 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **지역경제동향** | 17개 시·도별 경제 현황을 분석한 국가데이터처의 분기별 보고서. 생산, 소비, 물가, 고용 등을 포함한다. | [국가데이터처 보고서](https://kostat.go.kr/) |
| **인포그래픽** | 정보(Information)와 그래픽(Graphic)의 합성어. 복잡한 데이터를 시각적으로 표현한 이미지이다. | [위키백과 - 인포그래픽](https://ko.wikipedia.org/wiki/인포그래픽) |
| **일러두기** | 보고서 본문 앞에 위치하여 통계의 작성 목적, 용어 정의, 주의사항 등을 안내하는 페이지이다. | - |
| **부문별 보고서** | 경제 부문(생산, 소비, 고용 등)별로 전국 17개 시·도의 현황을 분석한 보고서 페이지들이다. | 프로젝트 내부 문서 |
| **시도별 보고서** | 특정 시·도의 모든 경제 지표(생산, 소비, 고용 등)를 종합 분석한 보고서 페이지들이다. | 프로젝트 내부 문서 |
| **통계표** | 숫자 데이터를 행과 열로 정리한 표. 보고서 부록에 세부 수치를 기록한다. | - |

---

## 5. 국가데이터처 관련 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **국가데이터처** | 대한민국의 데이터 및 통계 정책을 담당하는 중앙행정기관. 영문명: Ministry of Data and Statistics. (구 통계청) | [국가데이터처 홈페이지](https://kostat.go.kr/) |
| **경제활동인구** | 15세 이상 인구 중 취업자와 실업자를 합한 인구. 즉, 일할 능력과 의사가 있는 인구이다. | [국가통계포털 KOSIS](https://kosis.kr/) |
| **생산가능인구** | 경제활동을 할 수 있는 연령대(15세~64세)의 인구. 노동 공급의 잠재적 규모를 나타낸다. | [국가통계포털 KOSIS](https://kosis.kr/) |
| **취업자** | 조사 대상 주간에 수입을 목적으로 1시간 이상 일한 사람 또는 무급 가족 종사자이다. | [국가통계포털 KOSIS](https://kosis.kr/) |
| **실업자** | 조사 대상 주간에 일하지 않았으나, 일할 능력과 의사가 있고 적극적으로 구직활동을 한 사람이다. | [국가통계포털 KOSIS](https://kosis.kr/) |

---

## 6. 웹 개발 아키텍처 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **프론트엔드** | 사용자가 직접 보고 상호작용하는 웹의 클라이언트 측. HTML, CSS, JavaScript로 구현된다. | [위키백과 - 프론트엔드](https://ko.wikipedia.org/wiki/프론트엔드와_백엔드) |
| **백엔드** | 서버 측에서 데이터 처리, 비즈니스 로직, 데이터베이스 연동을 담당하는 부분이다. | [위키백과 - 백엔드](https://ko.wikipedia.org/wiki/프론트엔드와_백엔드) |
| **라우트** (Route) | URL 경로와 해당 경로를 처리하는 함수를 연결하는 것. Flask에서는 @app.route() 데코레이터로 정의한다. | [Flask 라우팅 문서](https://flask.palletsprojects.com/en/latest/quickstart/#routing) |
| **엔드포인트** | API에서 특정 리소스에 접근하기 위한 URL 경로. (예: /api/upload, /api/generate-preview) | - |
| **세션** (Session) | 사용자별 상태 정보를 서버에 저장하는 메커니즘. 로그인 상태, 업로드 파일 경로 등을 유지한다. | [위키백과 - 세션](https://ko.wikipedia.org/wiki/세션_(컴퓨터_과학)) |
| **비동기** (Async) | 작업 완료를 기다리지 않고 다음 작업을 수행하는 방식. 웹에서는 페이지 새로고침 없이 데이터를 주고받을 때 사용한다. | [MDN 비동기 JavaScript](https://developer.mozilla.org/ko/docs/Learn/JavaScript/Asynchronous) |
| **모달** (Modal) | 기존 페이지 위에 오버레이되어 표시되는 대화 상자. 결측치 입력, 확인 메시지 등에 사용된다. | [위키백과 - 모달 윈도](https://ko.wikipedia.org/wiki/모달_윈도) |
| **대시보드** (Dashboard) | 여러 정보와 기능을 한 화면에 종합적으로 표시하는 관리자 페이지. 차트, 그래프, 통계, 주요 기능 버튼 등을 포함한다. | [위키백과 - 대시보드](https://en.wikipedia.org/wiki/Dashboard_(web)) |
| **모듈** (Module) | 독립적인 기능을 수행하는 코드 단위. 재사용 가능하고 다른 모듈과 결합도가 낮아 유지보수가 쉽다. | [위키백과 - 모듈](https://ko.wikipedia.org/wiki/모듈) |
| **컴포넌트** (Component) | 재사용 가능한 독립적인 UI 요소나 기능 단위. React, Vue 등 프론트엔드 프레임워크에서 핵심 개념이다. | [MDN 컴포넌트](https://developer.mozilla.org/ko/docs/Glossary/Component) |
| **클라이언트** (Client) | 서버로부터 서비스를 요청하는 프로그램이나 장치. 웹 브라우저가 클라이언트의 대표적인 예이다. | [위키백과 - 클라이언트](https://ko.wikipedia.org/wiki/클라이언트_(컴퓨팅)) |
| **서버** (Server) | 클라이언트의 요청을 처리하고 응답을 제공하는 컴퓨터나 프로그램. 웹 서버, 데이터베이스 서버 등이 있다. | [위키백과 - 서버](https://ko.wikipedia.org/wiki/서버) |
| **요청** (Request) | 클라이언트가 서버에 보내는 데이터나 작업 요청. HTTP 요청에는 GET, POST, PUT, DELETE 등이 있다. | [MDN HTTP 요청](https://developer.mozilla.org/ko/docs/Web/HTTP/Methods) |
| **응답** (Response) | 서버가 클라이언트의 요청에 대해 반환하는 데이터나 결과. HTTP 응답에는 상태 코드, 헤더, 본문이 포함된다. | [MDN HTTP 응답](https://developer.mozilla.org/ko/docs/Web/HTTP/Status) |
| **HTTP** | HyperText Transfer Protocol의 약자. 웹에서 클라이언트와 서버 간 통신을 위한 프로토콜이다. | [위키백과 - HTTP](https://ko.wikipedia.org/wiki/HTTP) |
| **GET 요청** | 서버에서 데이터를 조회(읽기)하는 HTTP 메서드. URL 파라미터로 데이터를 전달하며, 브라우저 주소창에 입력하는 것이 GET 요청이다. | [MDN GET](https://developer.mozilla.org/ko/docs/Web/HTTP/Methods/GET) |
| **POST 요청** | 서버에 데이터를 생성(쓰기)하거나 제출하는 HTTP 메서드. 요청 본문에 데이터를 포함하여 전송한다. | [MDN POST](https://developer.mozilla.org/ko/docs/Web/HTTP/Methods/POST) |
| **상태 코드** (Status Code) | HTTP 응답의 상태를 나타내는 3자리 숫자. 200(성공), 404(찾을 수 없음), 500(서버 오류) 등이 있다. | [MDN HTTP 상태 코드](https://developer.mozilla.org/ko/docs/Web/HTTP/Status) |
| **UI** | User Interface의 약자. 사용자가 시스템과 상호작용하는 인터페이스. 버튼, 입력창, 메뉴 등의 화면 요소를 포함한다. | [위키백과 - 사용자 인터페이스](https://ko.wikipedia.org/wiki/사용자_인터페이스) |
| **UX** | User Experience의 약자. 사용자가 제품이나 서비스를 사용하면서 느끼는 전체적인 경험. 사용성, 접근성, 만족도를 포함한다. | [위키백과 - 사용자 경험](https://ko.wikipedia.org/wiki/사용자_경험) |
| **DOM** | Document Object Model의 약자. HTML 문서를 객체로 표현한 것. JavaScript로 DOM을 조작하여 화면을 동적으로 변경할 수 있다. | [MDN DOM](https://developer.mozilla.org/ko/docs/Web/API/Document_Object_Model) |
| **이벤트** (Event) | 사용자의 행동(클릭, 입력, 스크롤 등)이나 시스템에서 발생하는 사건. JavaScript에서 이벤트 핸들러로 처리한다. | [MDN 이벤트](https://developer.mozilla.org/ko/docs/Web/Events) |
| **이벤트 핸들러** (Event Handler) | 특정 이벤트가 발생했을 때 실행되는 함수. onclick, addEventListener() 등으로 등록한다. | [MDN 이벤트 핸들러](https://developer.mozilla.org/ko/docs/Web/Events/Event_handlers) |
| **렌더링** (Rendering) | 데이터나 템플릿을 HTML로 변환하여 화면에 표시하는 과정. 서버 사이드 렌더링(SSR)과 클라이언트 사이드 렌더링(CSR)이 있다. | [위키백과 - 렌더링](https://ko.wikipedia.org/wiki/렌더링) |
| **클라이언트 사이드 렌더링** (CSR) | 브라우저에서 JavaScript를 실행하여 HTML을 동적으로 생성하는 방식. React, Vue 등 SPA에서 사용한다. | [위키백과 - CSR](https://en.wikipedia.org/wiki/Single-page_application) |
| **서버 사이드 렌더링** (SSR) | 서버에서 HTML을 미리 생성하여 클라이언트에 전송하는 방식. Flask의 Jinja2 템플릿 렌더링이 SSR에 해당한다. | [위키백과 - SSR](https://en.wikipedia.org/wiki/Server-side_scripting) |
| **SPA** | Single Page Application의 약자. 단일 HTML 페이지에서 JavaScript로 동적으로 콘텐츠를 변경하는 웹 애플리케이션. | [위키백과 - SPA](https://ko.wikipedia.org/wiki/싱글_페이지_애플리케이션) |
| **API 호출** | 클라이언트가 서버의 API 엔드포인트로 HTTP 요청을 보내는 것. fetch API, axios 등으로 수행한다. | - |
| **비동기 통신** | 요청을 보낸 후 응답을 기다리지 않고 다른 작업을 수행할 수 있는 통신 방식. AJAX, fetch API 등이 사용된다. | [MDN 비동기 JavaScript](https://developer.mozilla.org/ko/docs/Learn/JavaScript/Asynchronous) |
| **AJAX** | Asynchronous JavaScript and XML의 약자. 페이지 새로고침 없이 서버와 비동기적으로 데이터를 주고받는 기술. | [MDN AJAX](https://developer.mozilla.org/ko/docs/Web/Guide/AJAX) |
| **상태 관리** (State Management) | 애플리케이션의 데이터 상태를 관리하는 것. 전역 상태, 로컬 상태 등을 일관되게 유지한다. | [위키백과 - 상태 관리](https://en.wikipedia.org/wiki/State_management) |
| **프로퍼티** (Property) | 객체의 속성이나 특징. JavaScript에서 객체의 변수나 함수를 프로퍼티라고 한다. | [MDN 프로퍼티](https://developer.mozilla.org/ko/docs/Glossary/Property) |
| **메서드** (Method) | 객체에 속한 함수. 객체의 동작을 정의한다. JavaScript에서 객체의 함수를 메서드라고 한다. | [MDN 메서드](https://developer.mozilla.org/ko/docs/Glossary/Method) |
| **함수** (Function) | 특정 작업을 수행하는 코드 블록. 입력(매개변수)을 받아 처리하고 결과(반환값)를 반환할 수 있다. | [MDN 함수](https://developer.mozilla.org/ko/docs/Web/JavaScript/Guide/Functions) |
| **변수** (Variable) | 데이터를 저장하는 메모리 공간의 이름. 값이 변경될 수 있는 저장소이다. | [MDN 변수](https://developer.mozilla.org/ko/docs/Web/JavaScript/Guide/Grammar_and_types#변수) |
| **함수형 프로그래밍** | 함수를 중심으로 프로그램을 구성하는 프로그래밍 패러다임. 순수 함수, 불변성을 강조한다. | [위키백과 - 함수형 프로그래밍](https://ko.wikipedia.org/wiki/함수형_프로그래밍) |
| **객체 지향 프로그래밍** | 객체와 클래스를 중심으로 프로그램을 구성하는 프로그래밍 패러다임. 상속, 캡슐화, 다형성을 활용한다. | [위키백과 - 객체 지향 프로그래밍](https://ko.wikipedia.org/wiki/객체_지향_프로그래밍) |
| **컨트롤러** (Controller) | 사용자의 입력을 받아 모델과 뷰 사이를 조정하는 역할. MVC 패턴의 C에 해당한다. Flask의 뷰 함수가 컨트롤러 역할을 한다. | [위키백과 - 컨트롤러](https://ko.wikipedia.org/wiki/컨트롤러) |
| **뷰** (View) | 사용자에게 보이는 화면이나 출력. MVC 패턴의 V에 해당한다. HTML 템플릿이 뷰에 해당한다. | [위키백과 - 뷰](https://ko.wikipedia.org/wiki/뷰) |
| **모델** (Model) | 데이터와 비즈니스 로직을 처리하는 부분. MVC 패턴의 M에 해당한다. 데이터베이스 연동, 데이터 검증 등이 포함된다. | [위키백과 - 모델](https://ko.wikipedia.org/wiki/모델) |
| **MVC 패턴** | Model-View-Controller 패턴. 애플리케이션을 데이터(모델), 화면(뷰), 제어(컨트롤러)로 분리하는 설계 패턴. | [위키백과 - MVC](https://ko.wikipedia.org/wiki/모델-뷰-컨트롤러) |
| **템플릿** (코드 템플릿) | 반복 사용되는 코드나 구조의 틀. Jinja2 템플릿은 HTML에 동적 데이터를 삽입하는 코드 템플릿이다. (문서 템플릿과 구분됨) | [위키백과 - 템플릿](https://ko.wikipedia.org/wiki/템플릿) |
| **스타일시트** | HTML 요소의 외관(색상, 레이아웃, 폰트 등)을 정의하는 CSS 파일이나 코드. | [MDN CSS](https://developer.mozilla.org/ko/docs/Web/CSS) |
| **레이아웃** (Layout) | 웹 페이지의 요소들을 배치하는 구조. 헤더, 사이드바, 본문, 푸터 등의 위치를 정의한다. | [MDN 레이아웃](https://developer.mozilla.org/ko/docs/Learn/CSS/CSS_layout) |
| **반응형 디자인** | 화면 크기에 따라 레이아웃이 자동으로 조정되는 웹 디자인. 모바일, 태블릿, 데스크톱에서 모두 최적화된다. | [MDN 반응형 디자인](https://developer.mozilla.org/ko/docs/Learn/CSS/CSS_layout/Responsive_Design) |
| **사이드바** (Sidebar) | 메인 콘텐츠 옆에 위치하는 보조 메뉴나 정보 영역. 대시보드에서 네비게이션 메뉴로 자주 사용된다. | - |
| **헤더** (Header) | 웹 페이지 상단에 위치하는 영역. 로고, 메뉴, 검색창 등이 포함된다. | [MDN header 요소](https://developer.mozilla.org/ko/docs/Web/HTML/Element/header) |
| **푸터** (Footer) | 웹 페이지 하단에 위치하는 영역. 저작권 정보, 연락처, 링크 등이 포함된다. | [MDN footer 요소](https://developer.mozilla.org/ko/docs/Web/HTML/Element/footer) |
| **네비게이션** (Navigation) | 웹사이트의 여러 페이지나 섹션으로 이동할 수 있게 하는 메뉴. 보통 헤더나 사이드바에 위치한다. | [MDN nav 요소](https://developer.mozilla.org/ko/docs/Web/HTML/Element/nav) |
| **콜백 함수** (Callback Function) | 특정 이벤트나 작업 완료 후 호출되는 함수. 비동기 처리에서 자주 사용된다. | [MDN 콜백](https://developer.mozilla.org/ko/docs/Glossary/Callback_function) |
| **프로미스** (Promise) | 비동기 작업의 완료 또는 실패를 나타내는 객체. .then(), .catch() 메서드로 처리한다. | [MDN Promise](https://developer.mozilla.org/ko/docs/Web/JavaScript/Reference/Global_Objects/Promise) |

---

## 7. 버전 관리 및 개발 도구 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **Git** | 분산 버전 관리 시스템. 코드의 변경 이력을 추적하고, 여러 개발자가 협업할 수 있게 한다. | [Git 공식 사이트](https://git-scm.com/) |
| **저장소** (Repository, Repo) | Git으로 관리되는 프로젝트 폴더. .git 디렉토리에 버전 관리 정보가 저장된다. 로컬 저장소와 원격 저장소로 나뉜다. | [Git 저장소 문서](https://git-scm.com/book/ko/v2/Git의-기초-Git-저장소-만들기) |
| **스테이징** (Staging) | 변경된 파일을 커밋할 준비가 되었다고 표시하는 과정. git add 명령으로 파일을 스테이징 영역(Index)에 추가한다. | [Git 스테이징 문서](https://git-scm.com/book/ko/v2/Git의-기초-변경사항-저장소에-저장하기) |
| **add** (추가) | 변경된 파일을 스테이징 영역에 추가하는 Git 명령. `git add <파일명>` 또는 `git add .`로 모든 변경사항을 추가할 수 있다. 커밋 전 반드시 거쳐야 하는 단계이다. | [Git add 문서](https://git-scm.com/docs/git-add) |
| **커밋** (Commit) | 스테이징된 변경 사항을 저장소에 영구적으로 기록하는 Git 명령. 각 커밋은 고유한 해시값(커밋 ID)을 가지며, 메시지와 함께 저장된다. | [Git 커밋 문서](https://git-scm.com/docs/git-commit) |
| **스테이징 영역** (Staging Area, Index) | 커밋되기 전에 파일이 임시로 저장되는 영역. git add로 추가한 파일들이 모이는 곳이다. 실제 커밋 전에 변경사항을 검토하고 선택적으로 커밋할 수 있다. | [Git Index 문서](https://git-scm.com/book/ko/v2/Git의-기초-변경사항-저장소에-저장하기) |
| **작업 디렉토리** (Working Directory) | 실제 파일들이 있는 작업 중인 디렉토리. 파일을 수정하면 작업 디렉토리에 변경사항이 반영되며, git add로 스테이징해야 커밋할 수 있다. | [Git 작업 디렉토리](https://git-scm.com/book/ko/v2/Git의-기초-변경사항-저장소에-저장하기) |
| **버전** (Version) | 특정 시점의 코드 상태를 나타내는 스냅샷. Git에서는 커밋이 버전에 해당한다. 각 버전(커밋)은 고유한 해시값으로 식별된다. | [Git 버전 관리](https://git-scm.com/book/ko/v2/시작하기-Git이란) |
| **해시** (Hash) | 커밋을 고유하게 식별하는 암호화된 문자열. SHA-1 알고리즘으로 생성되며, 커밋 ID라고도 부른다. 예: `a1b2c3d4e5f6...` 전체 해시는 40자리이며, 앞 7자리로도 식별 가능하다. | [Git 해시 문서](https://git-scm.com/book/ko/v2/Git의-내부-Git-객체) |
| **커밋 ID** (Commit ID) | 커밋을 식별하는 고유한 해시값. 커밋 메시지, 작성자, 날짜, 부모 커밋 등의 정보를 기반으로 생성된다. | [Git 커밋 ID](https://git-scm.com/book/ko/v2/Git의-기초-커밋-히스토리-조회하기) |
| **HEAD** | 현재 작업 중인 브랜치의 최신 커밋을 가리키는 포인터. HEAD를 이동하면 다른 커밋으로 체크아웃할 수 있다. | [Git HEAD 문서](https://git-scm.com/book/ko/v2/Git의-기초-커밋-히스토리-조회하기) |
| **브랜치** (Branch) | 독립적인 개발 흐름을 만드는 기능. 메인 코드에 영향을 주지 않고 새 기능을 개발하거나 버그를 수정할 수 있다. 기본 브랜치는 보통 main 또는 master이다. | [Git 브랜치 문서](https://git-scm.com/book/ko/v2/Git-브랜치-브랜치란-무엇인가) |
| **main 브랜치** | 기본 브랜치. 프로젝트의 안정적인 버전이 저장되는 메인 개발 라인이다. (이전에는 master 브랜치였음) | [Git main 브랜치](https://git-scm.com/docs/git-branch) |
| **체크아웃** (Checkout) | 특정 브랜치나 커밋으로 전환하는 Git 명령. `git checkout <브랜치명>` 또는 `git switch <브랜치명>`으로 사용한다. | [Git checkout 문서](https://git-scm.com/docs/git-checkout) |
| **switch** | 브랜치를 전환하는 Git 명령. checkout보다 명확한 의미를 가진다. `git switch <브랜치명>`으로 사용한다. | [Git switch 문서](https://git-scm.com/docs/git-switch) |
| **병합** (Merge) | 두 브랜치의 변경 사항을 하나로 합치는 과정. `git merge <브랜치명>`으로 실행하며, 같은 부분이 다르게 수정되면 충돌(conflict)이 발생할 수 있다. | [Git 병합 문서](https://git-scm.com/docs/git-merge) |
| **충돌** (Conflict) | 병합 시 같은 파일의 같은 부분이 서로 다르게 수정되어 Git이 자동으로 병합할 수 없는 상황. 수동으로 해결해야 한다. | [Git 충돌 해결](https://git-scm.com/book/ko/v2/Git-브랜치-브랜치의-분기) |
| **충돌 해결** (Conflict Resolution) | 병합 충돌을 수동으로 해결하는 과정. 충돌된 부분을 확인하고 어떤 버전을 사용할지 또는 두 버전을 어떻게 통합할지 결정한다. | [Git 충돌 해결 문서](https://git-scm.com/docs/git-merge#_how_conflicts_are_presented) |
| **원격 저장소** (Remote Repository) | 인터넷이나 네트워크에 있는 저장소. GitHub, GitLab 등에 호스팅되며, 여러 개발자가 협업할 수 있게 한다. | [Git 원격 저장소](https://git-scm.com/book/ko/v2/Git-기초-원격-저장소) |
| **로컬 저장소** (Local Repository) | 개발자의 컴퓨터에 있는 저장소. 원격 저장소와 독립적으로 작업할 수 있다. | [Git 로컬 저장소](https://git-scm.com/book/ko/v2/Git-기초-변경사항-저장소에-저장하기) |
| **origin** | 기본 원격 저장소의 별칭(alias). `git clone`으로 저장소를 복제하면 자동으로 origin이 설정된다. | [Git origin 문서](https://git-scm.com/book/ko/v2/Git-기초-원격-저장소) |
| **clone** (복제) | 원격 저장소를 로컬로 복사하는 Git 명령. `git clone <저장소-URL>`로 실행하면 전체 프로젝트와 히스토리가 복제된다. | [Git clone 문서](https://git-scm.com/docs/git-clone) |
| **fetch** (가져오기) | 원격 저장소의 변경사항을 가져오되 로컬 브랜치에 병합하지 않는 Git 명령. 원격의 상태만 확인할 때 사용한다. | [Git fetch 문서](https://git-scm.com/docs/git-fetch) |
| **pull** (당기기) | 원격 저장소의 변경사항을 가져와서 현재 브랜치에 자동으로 병합하는 Git 명령. `git fetch + git merge`와 같다. | [Git pull 문서](https://git-scm.com/docs/git-pull) |
| **push** (밀기) | 로컬 커밋을 원격 저장소에 업로드하는 Git 명령. `git push <원격명> <브랜치명>`으로 실행한다. | [Git push 문서](https://git-scm.com/docs/git-pull) |
| **싱크** (Sync) | 로컬과 원격 저장소를 동기화하는 작업. 일반적으로 `git pull`로 원격 변경사항을 가져오고 `git push`로 로컬 변경사항을 업로드하는 과정을 의미한다. | - |
| **status** (상태) | 작업 디렉토리와 스테이징 영역의 상태를 확인하는 Git 명령. 변경된 파일, 스테이징된 파일, 추적되지 않은 파일을 보여준다. | [Git status 문서](https://git-scm.com/docs/git-status) |
| **log** (로그) | 커밋 히스토리를 조회하는 Git 명령. 커밋 ID, 작성자, 날짜, 메시지 등을 확인할 수 있다. | [Git log 문서](https://git-scm.com/docs/git-log) |
| **diff** (차이) | 파일의 변경사항을 비교하여 보여주는 Git 명령. 커밋 전 변경 내용을 확인하거나 두 커밋 간 차이를 볼 때 사용한다. | [Git diff 문서](https://git-scm.com/docs/git-diff) |
| **reset** (리셋) | 커밋을 취소하거나 이전 상태로 되돌리는 Git 명령. `--soft`, `--mixed`, `--hard` 옵션으로 다양한 수준의 되돌림이 가능하다. | [Git reset 문서](https://git-scm.com/docs/git-reset) |
| **revert** (되돌리기) | 특정 커밋의 변경사항을 취소하는 새로운 커밋을 만드는 Git 명령. 히스토리를 유지하면서 변경을 되돌릴 때 사용한다. | [Git revert 문서](https://git-scm.com/docs/git-revert) |
| **stash** (임시 저장) | 작업 중인 변경사항을 임시로 저장하는 Git 명령. 커밋하지 않고도 브랜치를 전환하거나 다른 작업을 할 수 있다. | [Git stash 문서](https://git-scm.com/docs/git-stash) |
| **태그** (Tag) | 특정 커밋에 이름을 붙여 표시하는 기능. 주로 릴리스 버전(v1.0.0 등)을 표시할 때 사용한다. | [Git 태그 문서](https://git-scm.com/book/ko/v2/Git-기초-태그) |
| **cherry-pick** | 특정 커밋만 선택하여 현재 브랜치에 적용하는 Git 명령. 다른 브랜치의 특정 변경사항만 가져올 때 유용하다. | [Git cherry-pick 문서](https://git-scm.com/docs/git-cherry-pick) |
| **rebase** (리베이스) | 브랜치의 커밋을 다른 브랜치 위에 재배치하는 Git 명령. 히스토리를 선형으로 만들 수 있지만 충돌이 발생할 수 있다. | [Git rebase 문서](https://git-scm.com/docs/git-rebase) |
| **포크** (Fork) | 원격 저장소(GitHub 등)를 자신의 계정으로 복사하는 것. 원본 저장소를 수정할 권한이 없을 때 독립적으로 작업할 수 있다. | [GitHub Fork 문서](https://docs.github.com/ko/pull-requests/collaborating-with-pull-requests/working-with-forks/about-forks) |
| **풀 리퀘스트** (Pull Request, PR) | 자신의 변경사항을 원본 저장소에 반영해달라고 요청하는 것. 코드 리뷰를 통해 검토 후 병합된다. GitHub에서는 PR, GitLab에서는 Merge Request(MR)라고 한다. | [GitHub PR 문서](https://docs.github.com/ko/pull-requests/collaborating-with-pull-requests/about-pull-requests) |
| **머지 리퀘스트** (Merge Request, MR) | GitLab에서 사용하는 용어. Pull Request와 동일한 개념이다. | [GitLab MR 문서](https://docs.gitlab.com/ee/user/project/merge_requests/) |
| **이슈** (Issue) | 버그 리포트, 기능 제안, 작업 항목 등을 추적하는 기능. GitHub, GitLab 등에서 제공한다. | [GitHub Issues 문서](https://docs.github.com/ko/issues) |
| **브랜치 전략** (Branching Strategy) | 브랜치를 어떻게 생성하고 병합할지 결정하는 전략. Git Flow, GitHub Flow, GitLab Flow 등이 있다. | [위키백과 - Git Flow](https://en.wikipedia.org/wiki/Git-flow) |
| **Git Flow** | main, develop, feature, release, hotfix 브랜치를 사용하는 브랜치 전략. 대규모 프로젝트에 적합하다. | [Git Flow 문서](https://nvie.com/posts/a-successful-git-branching-model/) |
| **GitHub Flow** | main 브랜치와 feature 브랜치만 사용하는 간단한 브랜치 전략. 빠른 배포가 필요한 프로젝트에 적합하다. | [GitHub Flow 문서](https://docs.github.com/ko/get-started/quickstart/github-flow) |
| **분산 버전 관리** (Distributed Version Control) | 각 개발자가 전체 저장소 히스토리를 가지고 있는 버전 관리 방식. Git, Mercurial 등이 있다. 중앙 서버 없이도 작업 가능하다. | [위키백과 - 분산 버전 관리](https://ko.wikipedia.org/wiki/분산_버전_관리) |
| **버전 관리** (Version Control) | 코드 변경 이력을 추적하고 관리하는 것. Git이 대표적인 버전 관리 시스템이다. | [위키백과 - 버전 관리](https://ko.wikipedia.org/wiki/버전_관리) |
| **CI/CD** | Continuous Integration / Continuous Deployment의 약자. 코드 변경 시 자동으로 빌드, 테스트, 배포하는 프로세스이다. | [위키백과 - CI/CD](https://ko.wikipedia.org/wiki/CI/CD) |
| **Docker** | 애플리케이션을 컨테이너로 패키징하여 어디서든 동일하게 실행할 수 있게 하는 플랫폼이다. | [Docker 공식 사이트](https://www.docker.com/) |

---

## 8. 데이터베이스 및 아키텍처 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **데이터베이스** (Database) | 구조화된 데이터를 체계적으로 저장하고 관리하는 시스템. 데이터의 중앙 집중식 저장, 검색, 업데이트를 제공한다. | [위키백과 - 데이터베이스](https://ko.wikipedia.org/wiki/데이터베이스) |
| **RDBMS** | Relational Database Management System의 약자. 관계형 데이터베이스 관리 시스템. 테이블 간 관계를 정의하여 데이터를 관리한다. | [위키백과 - 관계형 데이터베이스](https://ko.wikipedia.org/wiki/관계형_데이터베이스) |
| **SQL** | Structured Query Language의 약자. 관계형 데이터베이스에서 데이터를 조회, 삽입, 수정, 삭제하기 위한 표준 언어이다. | [위키백과 - SQL](https://ko.wikipedia.org/wiki/SQL) |
| **NoSQL** | 관계형 데이터베이스가 아닌 데이터 저장 방식을 통칭. 문서형(MongoDB), 키-값형(Redis), 컬럼형(Cassandra) 등이 있다. | [위키백과 - NoSQL](https://ko.wikipedia.org/wiki/NoSQL) |
| **파일 기반 저장** (File-based Storage) | 데이터베이스 대신 파일 시스템(Excel, CSV, JSON 등)에 데이터를 저장하는 방식. 간단한 프로젝트나 일회성 작업에 적합하다. | - |
| **ORM** | Object-Relational Mapping의 약자. 객체 지향 프로그래밍과 관계형 데이터베이스 간의 변환을 자동화하는 기술. SQLAlchemy, Django ORM 등이 있다. | [위키백과 - ORM](https://ko.wikipedia.org/wiki/객체_관계_매핑) |
| **ACID** | Atomicity(원자성), Consistency(일관성), Isolation(격리성), Durability(지속성)의 약자. 데이터베이스 트랜잭션의 신뢰성을 보장하는 특성들이다. | [위키백과 - ACID](https://ko.wikipedia.org/wiki/ACID) |
| **트랜잭션** (Transaction) | 데이터베이스에서 하나의 논리적 작업 단위. 모두 성공하거나 모두 실패해야 하는 여러 작업을 묶는다. | [위키백과 - 트랜잭션](https://ko.wikipedia.org/wiki/트랜잭션) |
| **정규화** (Normalization) | 데이터베이스 설계 시 중복을 최소화하고 데이터 무결성을 향상시키기 위해 테이블을 분리하는 과정. | [위키백과 - 정규화](https://ko.wikipedia.org/wiki/데이터베이스_정규화) |
| **스키마** (Schema, 데이터베이스 스키마) | 데이터베이스의 구조 정의. 테이블, 컬럼, 데이터 타입, 제약 조건 등을 명시한다. (데이터 스키마와 구분됨) | [위키백과 - 데이터베이스 스키마](https://ko.wikipedia.org/wiki/데이터베이스_스키마) |
| **인덱스** (Index) | 데이터베이스에서 검색 속도를 향상시키기 위해 생성하는 자료 구조. 책의 목차와 유사한 개념이다. | [위키백과 - 인덱스](https://ko.wikipedia.org/wiki/데이터베이스_인덱스) |
| **쿼리** (Query) | 데이터베이스에서 데이터를 조회하거나 조작하기 위한 요청. SQL 문으로 작성된다. | [위키백과 - 쿼리](https://ko.wikipedia.org/wiki/쿼리) |
| **CRUD** | Create(생성), Read(읽기), Update(수정), Delete(삭제)의 약자. 데이터 조작의 기본 4가지 작업을 의미한다. | [위키백과 - CRUD](https://ko.wikipedia.org/wiki/CRUD) |
| **마이그레이션** (Migration) | 데이터베이스 스키마의 변경 이력을 관리하고 버전을 업데이트하는 과정. 테이블 생성, 수정, 삭제를 체계적으로 관리한다. | [위키백과 - 데이터베이스 마이그레이션](https://en.wikipedia.org/wiki/Schema_migration) |
| **연결 풀** (Connection Pool) | 데이터베이스 연결을 미리 생성해 두고 재사용하는 기법. 연결 생성 비용을 줄이고 성능을 향상시킨다. | [위키백과 - 연결 풀](https://en.wikipedia.org/wiki/Connection_pool) |
| **캐싱** (Caching) | 자주 사용되는 데이터를 빠른 저장소(메모리 등)에 임시 저장하여 반복 조회 시 성능을 향상시키는 기법. | [위키백과 - 캐시](https://ko.wikipedia.org/wiki/캐시) |
| **데이터 일관성** (Data Consistency) | 데이터베이스의 데이터가 항상 유효한 상태를 유지하는 것. 트랜잭션과 제약 조건으로 보장된다. | [위키백과 - 일관성](https://ko.wikipedia.org/wiki/일관성) |
| **동시성 제어** (Concurrency Control) | 여러 사용자가 동시에 데이터베이스에 접근할 때 데이터 무결성을 보장하는 메커니즘. 락(Lock) 등이 사용된다. | [위키백과 - 동시성 제어](https://en.wikipedia.org/wiki/Concurrency_control) |
| **프로토타입** (Prototype) | 최종 제품을 만들기 전에 핵심 기능만 구현한 초기 버전. 빠른 검증과 피드백 수집을 목적으로 한다. | [위키백과 - 프로토타입](https://ko.wikipedia.org/wiki/프로토타입) |
| **MVP** | Minimum Viable Product의 약자. 최소 기능 제품. 핵심 기능만으로 사용자에게 가치를 제공하는 최소한의 제품이다. | [위키백과 - MVP](https://ko.wikipedia.org/wiki/최소_기능_제품) |
| **확장성** (Scalability) | 시스템이 증가하는 부하나 데이터량에 대응할 수 있는 능력. 수평 확장(서버 추가)과 수직 확장(서버 성능 향상)이 있다. | [위키백과 - 확장성](https://ko.wikipedia.org/wiki/확장성) |
| **성능 최적화** (Performance Optimization) | 시스템의 처리 속도와 효율을 개선하는 과정. 쿼리 최적화, 인덱싱, 캐싱 등이 포함된다. | - |
| **코드 복잡도** (Code Complexity) | 코드의 이해하기 어려운 정도나 유지보수 난이도. 복잡도가 높으면 버그 발생 가능성이 증가한다. | [위키백과 - 복잡도](https://ko.wikipedia.org/wiki/복잡도) |
| **YAGNI** | "You Aren't Gonna Need It"의 약자. 필요하지 않은 기능을 미리 구현하지 말라는 프로그래밍 원칙. | [위키백과 - YAGNI](https://en.wikipedia.org/wiki/You_aren%27t_gonna_need_it) |
| **KISS** | "Keep It Simple, Stupid"의 약자. 단순하게 유지하라는 설계 원칙. 불필요한 복잡성을 피해야 한다는 의미이다. | [위키백과 - KISS](https://en.wikipedia.org/wiki/KISS_principle) |
| **비용-효과 분석** (Cost-Benefit Analysis) | 기능 추가나 기술 도입에 따른 비용과 얻는 이익을 비교 평가하는 분석. 개발 시간, 유지보수 비용 등을 고려한다. | [위키백과 - 비용편익분석](https://ko.wikipedia.org/wiki/비용편익분석) |
| **오버 엔지니어링** (Over-engineering) | 현재 요구사항에 비해 과도하게 복잡한 시스템을 만드는 것. 불필요한 추상화나 미래의 가능성만을 위한 설계를 포함한다. | [위키백과 - 오버 엔지니어링](https://en.wikipedia.org/wiki/Overengineering) |
| **레거시 시스템** (Legacy System) | 오래된 기술로 구축되어 있지만 여전히 사용 중인 시스템. 현대적 기술로 교체하기 어려운 경우가 많다. | [위키백과 - 레거시 시스템](https://ko.wikipedia.org/wiki/레거시_시스템) |
| **기술 부채** (Technical Debt) | 빠른 개발을 위해 장기적으로 더 많은 비용이 드는 선택을 하는 것. 나중에 리팩토링하거나 재설계해야 할 수 있다. | [위키백과 - 기술 부채](https://ko.wikipedia.org/wiki/기술_부채) |
| **아키텍처 결정** (Architecture Decision) | 시스템의 구조와 기술 스택을 선택하는 설계 결정. 프로젝트의 성공에 큰 영향을 미친다. | [위키백과 - 소프트웨어 아키텍처](https://ko.wikipedia.org/wiki/소프트웨어_아키텍처) |
| **단일 책임 원칙** (SRP) | Single Responsibility Principle의 약자. 하나의 클래스나 모듈은 하나의 책임만 가져야 한다는 객체 지향 설계 원칙. | [위키백과 - 단일 책임 원칙](https://ko.wikipedia.org/wiki/단일_책임_원칙) |
| **결합도** (Coupling) | 모듈 간 의존성의 정도. 결합도가 높으면 한 모듈의 변경이 다른 모듈에 영향을 많이 준다. | [위키백과 - 결합도](https://ko.wikipedia.org/wiki/결합도) |
| **응집도** (Cohesion) | 모듈 내부의 요소들이 관련되어 있는 정도. 높은 응집도는 유지보수성을 향상시킨다. | [위키백과 - 응집도](https://ko.wikipedia.org/wiki/응집도) |

---

## 9. 코드 품질 및 리팩토링 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **리팩토링** (Refactoring) | 코드의 외부 동작은 변경하지 않고 내부 구조를 개선하는 과정. 가독성, 유지보수성, 확장성을 향상시키기 위해 수행한다. | [위키백과 - 리팩토링](https://ko.wikipedia.org/wiki/리팩터링) |
| **코드 품질** (Code Quality) | 코드의 좋고 나쁨을 판단하는 기준. 가독성, 유지보수성, 성능, 안정성 등을 포함한다. | [위키백과 - 소프트웨어 품질](https://ko.wikipedia.org/wiki/소프트웨어_품질) |
| **유지보수성** (Maintainability) | 코드를 이해하고 수정하기 쉬운 정도. 명확한 네이밍, 적절한 주석, 모듈화가 유지보수성을 높인다. | [위키백과 - 유지보수성](https://en.wikipedia.org/wiki/Software_maintainability) |
| **가독성** (Readability) | 코드를 읽고 이해하기 쉬운 정도. 일관된 스타일, 명확한 변수명, 적절한 들여쓰기가 가독성을 향상시킨다. | - |
| **재사용성** (Reusability) | 코드를 다른 곳에서도 다시 사용할 수 있는 정도. 모듈화, 일반화를 통해 재사용성을 높일 수 있다. | - |
| **DRY 원칙** | "Don't Repeat Yourself"의 약자. 같은 코드를 반복하지 말라는 원칙. 함수나 클래스로 추출하여 중복을 제거한다. | [위키백과 - DRY](https://en.wikipedia.org/wiki/Don%27t_repeat_yourself) |
| **코드 리뷰** (Code Review) | 다른 개발자가 작성한 코드를 검토하여 버그, 개선점, 스타일을 확인하는 과정. 코드 품질 향상과 지식 공유에 도움이 된다. | [위키백과 - 코드 리뷰](https://ko.wikipedia.org/wiki/코드_리뷰) |
| **테스트 커버리지** (Test Coverage) | 코드 중 얼마나 많은 부분이 테스트로 검증되었는지를 나타내는 지표. 높은 커버리지는 버그 발견 가능성을 높인다. | [위키백과 - 코드 커버리지](https://ko.wikipedia.org/wiki/코드_커버리지) |
| **레거시 코드** (Legacy Code) | 오래되어 이해하기 어렵거나 테스트가 없는 코드. 리팩토링이 필요한 코드를 가리킨다. | [위키백과 - 레거시 시스템](https://ko.wikipedia.org/wiki/레거시_시스템) |
| **코드 냄새** (Code Smell) | 코드에서 잠재적인 문제를 암시하는 패턴이나 구조. 긴 함수, 중복 코드, 과도한 매개변수 등이 코드 냄새에 해당한다. | [위키백과 - 코드 냄새](https://en.wikipedia.org/wiki/Code_smell) |
| **정적 분석** (Static Analysis) | 코드를 실행하지 않고 소스 코드 자체를 분석하여 버그나 문제점을 찾아내는 기법. 컴파일 시점에 수행된다. | [위키백과 - 정적 분석](https://ko.wikipedia.org/wiki/정적_프로그램_분석) |
| **린터** (Linter) | 코드의 문법 오류, 스타일 위반, 잠재적 버그를 자동으로 찾아주는 도구. Python에서는 pylint, flake8 등이 있다. | [위키백과 - 린터](https://en.wikipedia.org/wiki/Lint_(software)) |
| **포매터** (Formatter) | 코드 스타일을 자동으로 일관되게 맞춰주는 도구. Python에서는 black, autopep8 등이 있다. | [Black 포매터](https://black.readthedocs.io/) |
| **코드 스타일** (Code Style) | 코드 작성 규칙이나 관례. PEP 8은 Python의 공식 코드 스타일 가이드이다. | [PEP 8](https://www.python.org/dev/peps/pep-0008/) |
| **네이밍 컨벤션** (Naming Convention) | 변수, 함수, 클래스 등의 이름을 짓는 규칙. 일관된 네이밍은 코드 가독성을 높인다. | [PEP 8 - 네이밍](https://www.python.org/dev/peps/pep-0008/#naming-conventions) |
| **주석** (Comment) | 코드에 대한 설명을 작성한 텍스트. 코드의 의도나 복잡한 로직을 설명할 때 사용한다. | [PEP 8 - 주석](https://www.python.org/dev/peps/pep-0008/#comments) |
| **문서화** (Documentation) | 코드의 사용법, API, 설계 의도를 설명한 문서. 함수 docstring, README, API 문서 등이 있다. | [위키백과 - 소프트웨어 문서화](https://en.wikipedia.org/wiki/Software_documentation) |
| **Docstring** | Python 함수나 클래스의 첫 번째 문장에 작성하는 문서 문자열. 함수의 기능, 매개변수, 반환값을 설명한다. | [PEP 257 - Docstring](https://www.python.org/dev/peps/pep-0257/) |
| **단위 테스트** (Unit Test) | 개별 함수나 메서드의 동작을 검증하는 테스트. 작은 단위로 나누어 테스트하여 버그를 빠르게 발견할 수 있다. | [위키백과 - 단위 테스트](https://ko.wikipedia.org/wiki/단위_테스트) |
| **통합 테스트** (Integration Test) | 여러 모듈이나 컴포넌트를 함께 작동시켜 검증하는 테스트. 모듈 간 상호작용을 확인한다. | [위키백과 - 통합 테스트](https://ko.wikipedia.org/wiki/통합_테스트) |
| **TDD** | Test-Driven Development의 약자. 테스트 주도 개발. 테스트를 먼저 작성하고, 그 테스트를 통과하는 코드를 작성하는 개발 방법론. | [위키백과 - TDD](https://ko.wikipedia.org/wiki/테스트_주도_개발) |
| **버그** (Bug) | 프로그램의 오류나 결함. 의도하지 않은 동작이나 프로그램 중단을 일으킨다. | [위키백과 - 버그](https://ko.wikipedia.org/wiki/버그) |
| **디버깅** (Debugging) | 프로그램의 버그(오류)를 찾아내고 수정하는 과정. 디버거 도구를 사용하여 코드를 단계별로 실행하고 변수 값을 확인한다. | [위키백과 - 디버깅](https://ko.wikipedia.org/wiki/디버깅) |
| **추상화** (Abstraction) | 복잡한 것을 단순화하여 표현하는 것. 공통된 특징만 남기고 세부 사항을 숨긴다. | [위키백과 - 추상화](https://ko.wikipedia.org/wiki/추상화) |
| **캡슐화** (Encapsulation) | 데이터와 메서드를 하나의 단위로 묶고, 내부 구현을 숨기는 것. 객체 지향 프로그래밍의 핵심 개념이다. | [위키백과 - 캡슐화](https://ko.wikipedia.org/wiki/캡슐화) |
| **인터페이스** (Interface) | 클래스나 모듈이 제공하는 기능을 정의한 계약. 구현 세부 사항을 숨기고 사용 방법만 공개한다. | [위키백과 - 인터페이스](https://ko.wikipedia.org/wiki/인터페이스_(컴퓨팅)) |
| **의존성 주입** (Dependency Injection) | 객체가 필요로 하는 의존성을 외부에서 주입하는 디자인 패턴. 결합도를 낮추고 테스트 가능성을 높인다. | [위키백과 - 의존성 주입](https://ko.wikipedia.org/wiki/의존성_주입) |
| **인터페이스 분리 원칙** (ISP) | Interface Segregation Principle의 약자. 클라이언트가 사용하지 않는 인터페이스에 의존하지 않아야 한다는 원칙. | [위키백과 - SOLID 원칙](https://ko.wikipedia.org/wiki/SOLID) |
| **개방-폐쇄 원칙** (OCP) | Open-Closed Principle의 약자. 확장에는 열려 있고 수정에는 닫혀 있어야 한다는 원칙. | [위키백과 - SOLID 원칙](https://ko.wikipedia.org/wiki/SOLID) |
| **리스코프 치환 원칙** (LSP) | Liskov Substitution Principle의 약자. 하위 타입은 상위 타입으로 대체 가능해야 한다는 원칙. | [위키백과 - SOLID 원칙](https://ko.wikipedia.org/wiki/SOLID) |
| **의존성 역전 원칙** (DIP) | Dependency Inversion Principle의 약자. 고수준 모듈은 저수준 모듈에 의존하지 않아야 한다는 원칙. | [위키백과 - SOLID 원칙](https://ko.wikipedia.org/wiki/SOLID) |
| **SOLID 원칙** | 객체 지향 설계의 5가지 원칙. SRP, OCP, LSP, ISP, DIP를 합쳐서 SOLID라고 한다. | [위키백과 - SOLID](https://ko.wikipedia.org/wiki/SOLID) |
| **클린 코드** (Clean Code) | 읽기 쉽고 이해하기 쉬운 코드. 로버트 C. 마틴의 "클린 코드" 책에서 제시한 개념이다. | [위키백과 - Clean Code](https://en.wikipedia.org/wiki/Clean_code) |
| **자기 설명적 코드** (Self-documenting Code) | 주석 없이도 코드 자체만으로 의도가 명확히 드러나는 코드. 명확한 변수명, 함수명을 사용한다. | - |
| **복잡도 분석** (Complexity Analysis) | 알고리즘의 시간 복잡도와 공간 복잡도를 분석하는 것. 빅오 표기법(O notation)을 사용한다. | [위키백과 - 시간 복잡도](https://ko.wikipedia.org/wiki/시간_복잡도) |
| **성능 프로파일링** (Performance Profiling) | 프로그램의 성능 병목 지점을 찾아내는 과정. 어느 부분이 느린지 분석하여 최적화한다. | [위키백과 - 프로파일링](https://ko.wikipedia.org/wiki/성능_프로파일링) |

---

## 10. 배포 및 패키징 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **패키징** (Packaging) | Python 프로그램을 배포 가능한 형태로 만드는 과정. 소스 코드, 의존성, 리소스를 하나의 패키지로 묶는다. PyInstaller, cx_Freeze 등을 사용하여 실행 파일(.exe)로 만들 수 있다. | [Python 패키징 가이드](https://packaging.python.org/) |
| **배포** (Deployment) | 개발된 소프트웨어를 사용자가 설치하고 실행할 수 있도록 제공하는 과정. 실행 파일, 설치 패키지, 웹 배포 등이 포함된다. | [위키백과 - 소프트웨어 배포](https://en.wikipedia.org/wiki/Software_deployment) |
| **빌드** (Build) | 소스 코드를 실행 가능한 형태로 변환하는 과정. 컴파일, 패키징, 리소스 포함 등을 수행한다. | [위키백과 - 빌드](https://en.wikipedia.org/wiki/Software_build) |
| **실행 파일** (Executable) | 직접 실행할 수 있는 파일. Windows에서는 .exe, macOS/Linux에서는 실행 권한이 있는 파일이다. Python은 인터프리터 언어이지만 PyInstaller 등으로 실행 파일을 만들 수 있다. | [위키백과 - 실행 파일](https://ko.wikipedia.org/wiki/실행_파일) |
| **독립 실행 파일** (Standalone Executable) | 외부 의존성 없이 단독으로 실행 가능한 파일. Python 인터프리터와 필요한 라이브러리를 모두 포함한다. | - |
| **PyInstaller** | Python 애플리케이션을 Windows, macOS, Linux용 실행 파일로 변환하는 도구. 가장 널리 사용되는 Python 패키징 도구 중 하나이다. | [PyInstaller 공식 문서](https://pyinstaller.org/) |
| **cx_Freeze** | Python 애플리케이션을 실행 파일로 변환하는 크로스 플랫폼 도구. PyInstaller의 대안으로 사용된다. | [cx_Freeze 공식 문서](https://cx-freeze.readthedocs.io/) |
| **py2exe** | Windows 전용으로 Python 애플리케이션을 .exe 파일로 변환하는 도구. 현재는 PyInstaller를 더 많이 사용한다. | [py2exe 공식 문서](https://www.py2exe.org/) |
| **Nuitka** | Python을 C++로 컴파일한 후 실행 파일로 만드는 도구. 성능 향상과 실행 파일 생성을 동시에 제공한다. | [Nuitka 공식 문서](https://nuitka.net/) |
| **컴파일** (Compile) | 고급 언어로 작성된 소스 코드를 저급 언어(기계어, 바이트코드)로 변환하는 과정. Python은 바이트코드로 컴파일되지만 실행 파일로 만들 때 완전 컴파일도 가능하다. | [위키백과 - 컴파일](https://ko.wikipedia.org/wiki/컴파일) |
| **인터프리터** (Interpreter) | 소스 코드를 한 줄씩 읽어 실행하는 프로그램. Python은 기본적으로 인터프리터 언어이지만, 실행 파일로 만들면 인터프리터를 포함한 형태가 된다. | [위키백과 - 인터프리터](https://ko.wikipedia.org/wiki/인터프리터) |
| **바이트코드** (Bytecode) | 컴파일러나 인터프리터가 생성하는 중간 코드. Python의 .pyc 파일이 바이트코드이다. | [위키백과 - 바이트코드](https://ko.wikipedia.org/wiki/바이트코드) |
| **의존성 포함** (Bundling) | 실행 파일에 필요한 라이브러리와 리소스를 함께 포함시키는 것. 외부 설치 없이 실행 가능하게 만든다. | - |
| **리소스** (Resource) | 프로그램 실행에 필요한 추가 파일. 이미지, 데이터 파일, 설정 파일 등이 포함된다. | - |
| **설치 패키지** (Installation Package) | 프로그램을 설치하기 위한 패키지 파일. Windows의 .msi, macOS의 .dmg, Linux의 .deb/.rpm 등이 있다. | [위키백과 - 소프트웨어 패키지](https://ko.wikipedia.org/wiki/소프트웨어_패키지) |
| **휜 애플리케이션** (Frozen Application) | Python 인터프리터와 필요한 모듈을 포함하여 하나의 실행 파일로 만든 애플리케이션. "고정된" 애플리케이션이라는 의미이다. | - |
| **바이너리** (Binary) | 컴퓨터가 직접 실행할 수 있는 기계어 코드. 실행 파일은 바이너리 형태이다. | [위키백과 - 바이너리](https://ko.wikipedia.org/wiki/바이너리) |
| **크로스 플랫폼** (Cross-platform) | 여러 운영체제(Windows, macOS, Linux)에서 실행 가능한 프로그램. Python은 크로스 플랫폼 언어이다. | [위키백과 - 크로스 플랫폼](https://ko.wikipedia.org/wiki/크로스_플랫폼) |
| **스펙 파일** (Spec File) | PyInstaller에서 빌드 설정을 정의하는 파일. 포함할 파일, 아이콘, 데이터 파일 등을 지정한다. | [PyInstaller Spec 파일](https://pyinstaller.org/en/stable/spec-files.html) |
| **단일 파일 모드** (One-file Mode) | 모든 의존성을 하나의 실행 파일로 묶는 PyInstaller 모드. 파일 하나만 배포하면 된다. | [PyInstaller 문서](https://pyinstaller.org/en/stable/usage.html) |
| **단일 디렉토리 모드** (One-dir Mode) | 실행 파일과 필요한 라이브러리를 별도 폴더로 배포하는 PyInstaller 모드. 시작이 빠르지만 여러 파일이 필요하다. | [PyInstaller 문서](https://pyinstaller.org/en/stable/usage.html) |
| **UPX 압축** | UPX(Ultimate Packer for eXecutables)를 사용하여 실행 파일 크기를 줄이는 압축 기법. | [UPX 공식 사이트](https://upx.github.io/) |
| **코드 서명** (Code Signing) | 실행 파일에 디지털 서명을 추가하여 출처를 증명하는 것. Windows에서 "알 수 없는 게시자" 경고를 제거할 수 있다. | [위키백과 - 코드 서명](https://en.wikipedia.org/wiki/Code_signing) |
| **배포판** (Distribution, Python 패키지 배포판) | 특정 형태로 패키징된 소프트웨어. Python 패키지의 경우 wheel, sdist 등이 있다. (소프트웨어 배포 Deployment와 구분됨) | [Python 배포 가이드](https://packaging.python.org/guides/distributing-packages-using-setuptools/) |
| **wheel** | Python 패키지의 표준 배포 형식. .whl 확장자를 가지며 빠르게 설치할 수 있다. | [위키백과 - wheel](https://en.wikipedia.org/wiki/Python_wheel) |
| **setuptools** | Python 패키지를 빌드하고 배포하는 도구. setup.py 파일을 사용하여 패키지 메타데이터를 정의한다. | [setuptools 공식 문서](https://setuptools.readthedocs.io/) |
| **setup.py** | Python 패키지의 빌드 및 배포 설정을 정의하는 파일. 패키지명, 버전, 의존성 등을 명시한다. | [Python 패키징 가이드](https://packaging.python.org/guides/writing-pyproject-toml/) |
| **의존성** (Dependency) | 프로그램 실행에 필요한 외부 라이브러리나 패키지. requirements.txt에 나열된다. | [위키백과 - 의존성](https://en.wikipedia.org/wiki/Dependency_(computer_science)) |
| **가상 환경** (Virtual Environment) | 프로젝트별로 독립적인 Python 패키지 환경을 만드는 것. 의존성 충돌을 방지한다. | [Python venv 문서](https://docs.python.org/ko/3/library/venv.html) |
| **런타임** (Runtime) | 프로그램이 실행되는 환경. Python 런타임, .NET 런타임 등이 있다. | [위키백과 - 런타임](https://ko.wikipedia.org/wiki/런타임) |
| **포터블 애플리케이션** (Portable Application) | 설치 없이 어디서나 실행 가능한 애플리케이션. USB 등으로 옮겨 실행할 수 있다. | [위키백과 - 포터블 소프트웨어](https://ko.wikipedia.org/wiki/포터블_소프트웨어) |
| **자동 업데이트** (Auto-update) | 프로그램이 자동으로 업데이트를 확인하고 설치하는 기능. 실행 파일로 배포할 때 구현이 복잡할 수 있다. | - |
| **배포 전략** (Deployment Strategy) | 소프트웨어를 사용자에게 제공하는 방법. 웹 배포, 설치 패키지, 실행 파일 등이 있다. | - |

---

## 11. 오류 처리 및 장애 대응 용어

| 용어 | 정의 | 비고 (참고 링크) |
|------|------|------------------|
| **fallback** (대체 수단) | 기본적으로 사용하는 방법이나 시스템이 실패하거나 사용할 수 없을 때 대체로 사용하는 방법. 원래 의도했던 방식이 실행되지 않을 때 안전장치로 다른 방식으로 진행하는 것이다. 예: API 호출 실패 시 기본값 반환, 라이브러리 미설치 시 대체 코드 실행, 네트워크 오류 시 캐시된 데이터 사용 | [MDN Fallback](https://developer.mozilla.org/ko/docs/Glossary/Fallback) |
| **failover** (장애 조치) | 시스템이나 서비스가 장애를 일으킬 때 자동으로 다른 백업 시스템으로 전환하는 메커니즘. 고가용성을 보장하기 위한 기술이다. | [위키백과 - 장애 조치](https://ko.wikipedia.org/wiki/장애_조치) |
| **고가용성** (High Availability) | 시스템이 장시간 동안 중단 없이 동작하는 능력. 다중화, 자동 복구 등을 통해 달성한다. | [위키백과 - 고가용성](https://ko.wikipedia.org/wiki/고가용성) |
| **장애 허용** (Fault Tolerance) | 시스템의 일부 구성 요소가 실패해도 전체 시스템이 계속 동작할 수 있는 능력. 장애를 견디는 설계이다. | [위키백과 - 장애 허용](https://ko.wikipedia.org/wiki/장애_허용) |
| **우아한 성능 저하** (Graceful Degradation) | 시스템 일부가 실패하거나 사용할 수 없을 때, 전체 기능 대신 기본 기능이라도 제공하는 전략. 사용자 경험을 최소한으로 유지한다. | [위키백과 - Graceful Degradation](https://en.wikipedia.org/wiki/Graceful_degradation) |
| **점진적 향상** (Progressive Enhancement) | 기본 기능부터 시작하여 향상된 기능을 점진적으로 추가하는 설계 전략. 기본 기능은 항상 작동하도록 보장한다. | [위키백과 - 점진적 향상](https://en.wikipedia.org/wiki/Progressive_enhancement) |
| **회로 차단기 패턴** (Circuit Breaker Pattern) | 반복적으로 실패하는 서비스 호출을 차단하여 전체 시스템을 보호하는 패턴. 일정 시간 후 재시도한다. | [위키백과 - Circuit Breaker](https://en.wikipedia.org/wiki/Circuit_breaker_design_pattern) |
| **재시도 패턴** (Retry Pattern) | 실패한 작업을 자동으로 다시 시도하는 패턴. 일시적인 오류에서 유용하다. 지수 백오프(exponential backoff)와 함께 사용된다. | [위키백과 - Retry Pattern](https://en.wikipedia.org/wiki/Retry_pattern) |
| **지수 백오프** (Exponential Backoff) | 재시도 간격을 점진적으로 늘려가는 전략. 예: 1초, 2초, 4초, 8초... 서버 부하를 줄이면서 재시도한다. | [위키백과 - Exponential Backoff](https://en.wikipedia.org/wiki/Exponential_backoff) |
| **타임아웃** (Timeout) | 작업 수행에 허용되는 최대 시간. 이 시간을 초과하면 작업을 중단하고 오류로 처리한다. 무한 대기를 방지한다. | [MDN Timeout](https://developer.mozilla.org/ko/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) |
| **헬스 체크** (Health Check) | 시스템이나 서비스가 정상적으로 동작하는지 주기적으로 확인하는 메커니즘. /health 엔드포인트 등을 사용한다. | [위키백과 - Health Check](https://en.wikipedia.org/wiki/Health_check_(computing)) |
| **하트비트** (Heartbeat) | 시스템이 살아있음을 주기적으로 알리는 신호. 장애를 빠르게 감지할 수 있다. | [위키백과 - Heartbeat](https://en.wikipedia.org/wiki/Heartbeat_(computing)) |
| **모니터링** (Monitoring) | 시스템의 상태, 성능, 오류를 지속적으로 관찰하고 기록하는 것. 문제를 조기에 발견할 수 있다. | [위키백과 - 모니터링](https://en.wikipedia.org/wiki/System_monitoring) |
| **알림** (Alert) | 시스템에서 중요한 이벤트나 오류가 발생했을 때 개발자나 운영자에게 알리는 메커니즘. 이메일, SMS, Slack 등으로 전송된다. | - |
| **로깅** (Logging) | 프로그램 실행 중 발생하는 이벤트를 기록하는 것. 디버깅, 모니터링, 감사에 활용된다. | [Python logging](https://docs.python.org/ko/3/library/logging.html) |
| **예외 처리** (Exception Handling) | 프로그램 실행 중 발생할 수 있는 오류 상황을 처리하는 코드. try-except 구문으로 구현한다. | [Python 예외 처리](https://docs.python.org/ko/3/tutorial/errors.html) |
| **오류 처리** (Error Handling) | 오류가 발생했을 때 적절히 처리하는 과정. 사용자에게 친절한 메시지 표시, 로깅, 복구 시도 등이 포함된다. | [위키백과 - 예외 처리](https://ko.wikipedia.org/wiki/예외_처리) |
| **방어적 프로그래밍** (Defensive Programming) | 예상치 못한 입력이나 상황에 대비하여 안전하게 코드를 작성하는 기법. 검증, 예외 처리, 기본값 사용 등이 포함된다. | [위키백과 - 방어적 프로그래밍](https://en.wikipedia.org/wiki/Defensive_programming) |
| **안전 모드** (Safe Mode) | 오류 발생 시 최소한의 기능만 동작하는 제한 모드. 시스템이 완전히 중단되지 않도록 한다. | - |
| **롤백** (Rollback) | 시스템 변경 사항을 이전 상태로 되돌리는 것. 배포 실패나 오류 발생 시 이전 안정 버전으로 복구한다. | [위키백과 - 롤백](https://en.wikipedia.org/wiki/Rollback_(data_management)) |
| **롤포워드** (Rollforward) | 롤백 대신 문제를 수정하여 앞으로 진행하는 방식. 빠른 수정이 가능할 때 사용한다. | - |
| **다중화** (Redundancy) | 시스템 구성 요소를 중복하여 배치하는 것. 하나가 실패해도 다른 것으로 대체할 수 있다. | [위키백과 - 중복성](https://ko.wikipedia.org/wiki/중복성) |
| **로드 밸런싱** (Load Balancing) | 여러 서버에 요청을 분산시키는 것. 서버 과부하를 방지하고 가용성을 향상시킨다. | [위키백과 - 로드 밸런싱](https://ko.wikipedia.org/wiki/로드_밸런싱) |
| **백업** (Backup) | 데이터나 시스템의 복사본을 저장하는 것. 장애 발생 시 복구에 사용한다. | [위키백과 - 백업](https://ko.wikipedia.org/wiki/백업) |
| **재해 복구** (Disaster Recovery) | 큰 장애나 재해 발생 후 시스템을 복구하는 계획과 프로세스. 비즈니스 연속성을 보장한다. | [위키백과 - 재해 복구](https://ko.wikipedia.org/wiki/재해_복구) |
| **복구 시간 목표** (RTO) | Recovery Time Objective의 약자. 장애 발생 후 시스템을 복구해야 하는 최대 허용 시간. | [위키백과 - RTO](https://en.wikipedia.org/wiki/Recovery_time_objective) |
| **복구 시점 목표** (RPO) | Recovery Point Objective의 약자. 허용 가능한 데이터 손실의 최대 범위. 마지막 백업 이후 얼마나 데이터를 잃어도 되는지를 나타낸다. | [위키백과 - RPO](https://en.wikipedia.org/wiki/Recovery_point_objective) |
| **입력 검증** (Input Validation) | 사용자 입력을 검증하여 악의적인 코드나 잘못된 데이터를 차단하는 것. 보안과 안정성에 중요하다. 데이터 검증의 일종이지만 보안에 중점을 둔다. | [OWASP - Input Validation](https://owasp.org/www-community/Improper_Input_Validation) |
| **사전 조건 검사** (Precondition Check) | 함수나 메서드 실행 전에 필요한 조건이 충족되었는지 확인하는 것. 방어적 프로그래밍의 일종이다. | - |
| **사후 조건 검사** (Postcondition Check) | 함수나 메서드 실행 후 결과가 올바른지 확인하는 것. 계약에 의한 설계(Design by Contract)에서 사용된다. | - |
| **가드 클로즈** (Guard Clause) | 함수 초기에 조건을 확인하고 조기에 반환하는 패턴. 중첩된 if문을 줄여 가독성을 향상시킨다. | - |
| **기본값** (Default Value) | 값이 제공되지 않았을 때 사용하는 기본 설정값. None 대신 의미 있는 기본값을 사용한다. | - |
| **널 체크** (Null Check) | 값이 null(None)인지 확인하는 것. NullPointerException이나 AttributeError를 방지한다. | - |
| **옵셔널** (Optional) | 값이 있을 수도 있고 없을 수도 있음을 나타내는 타입. Python의 Optional[Type], Rust의 Option 등이 있다. | - |

---

## 부록: 약어 정리

| 약어 | 풀이 | 의미 |
|------|------|------|
| **ACID** | Atomicity, Consistency, Isolation, Durability | 데이터베이스 트랜잭션 특성 |
| **API** | Application Programming Interface | 소프트웨어 간 통신 인터페이스 |
| **CLI** | Command Line Interface | 명령줄 인터페이스 |
| **CPI** | Consumer Price Index | 소비자물가지수 |
| **CRUD** | Create, Read, Update, Delete | 데이터 조작 기본 작업 |
| **CSS** | Cascading Style Sheets | 스타일시트 언어 |
| **DIP** | Dependency Inversion Principle | 의존성 역전 원칙 |
| **DRY** | Don't Repeat Yourself | 코드 중복 방지 원칙 |
| **ETL** | Extract, Transform, Load | 데이터 추출·변환·적재 |
| **AJAX** | Asynchronous JavaScript and XML | 비동기 웹 통신 기술 |
| **CSR** | Client-Side Rendering | 클라이언트 사이드 렌더링 |
| **DOM** | Document Object Model | 문서 객체 모델 |
| **GDP** | Gross Domestic Product | 국내총생산 |
| **GRDP** | Gross Regional Domestic Product | 지역내총생산 |
| **HTML** | HyperText Markup Language | 웹 마크업 언어 |
| **HTTP** | HyperText Transfer Protocol | 웹 통신 프로토콜 |
| **IIP** | Index of Industrial Production | 광공업생산지수 |
| **ISP** | Interface Segregation Principle | 인터페이스 분리 원칙 |
| **JSON** | JavaScript Object Notation | 데이터 교환 포맷 |
| **KOSIS** | Korean Statistical Information Service | 국가통계포털 |
| **LSP** | Liskov Substitution Principle | 리스코프 치환 원칙 |
| **MVC** | Model-View-Controller | 모델-뷰-컨트롤러 패턴 |
| **MVP** | Minimum Viable Product | 최소 기능 제품 |
| **NaN** | Not a Number | 숫자가 아님 (결측치) |
| **NoSQL** | Not Only SQL | 비관계형 데이터베이스 |
| **OCP** | Open-Closed Principle | 개방-폐쇄 원칙 |
| **ORM** | Object-Relational Mapping | 객체-관계 매핑 |
| **QoQ** | Quarter over Quarter | 전분기비 |
| **RDBMS** | Relational Database Management System | 관계형 데이터베이스 관리 시스템 |
| **REST** | Representational State Transfer | 웹 서비스 설계 아키텍처 |
| **RPO** | Recovery Point Objective | 복구 시점 목표 |
| **RTO** | Recovery Time Objective | 복구 시간 목표 |
| **SOLID** | Single Responsibility, Open-Closed, Liskov Substitution, Interface Segregation, Dependency Inversion | 객체 지향 설계 5원칙 |
| **SPA** | Single Page Application | 싱글 페이지 애플리케이션 |
| **SQL** | Structured Query Language | 구조화 질의어 |
| **SRP** | Single Responsibility Principle | 단일 책임 원칙 |
| **TDD** | Test-Driven Development | 테스트 주도 개발 |
| **SSR** | Server-Side Rendering | 서버 사이드 렌더링 |
| **UI** | User Interface | 사용자 인터페이스 |
| **URL** | Uniform Resource Locator | 웹 주소 |
| **UX** | User Experience | 사용자 경험 |
| **WSGI** | Web Server Gateway Interface | 웹 서버 게이트웨이 인터페이스 |
| **YAGNI** | You Aren't Gonna Need It | 필요 없는 기능 구현 금지 원칙 |
| **YoY** | Year over Year | 전년동기비 |

---

> **최종 수정일**: 2026년 1월 1일  
> **작성자**: 캡스톤 프로젝트 팀

