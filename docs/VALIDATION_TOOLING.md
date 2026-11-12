# 재현 가능한 저장소 검증

저장소 루트에서 다음 명령 하나로 현재 소스의 검증을 실행합니다. Python 3.10 이상과 표준 라이브러리만 필요합니다.

```sh
python -B tools/verify.py
```

검증은 순서대로 배포 BAS 재생성 비교, 세 가지 VBA 조건 분기의 정적 구조 검사, 현재 배포 증거의 해시·버전 일치, 검증 도구의 결함 주입 검사, 메일 백업 모델·소스 계약 검사를 수행합니다. 하나라도 실패하면 0이 아닌 종료 코드로 끝납니다. 메일 파일이나 사용자 Outlook 환경에 접근하지 않습니다.

## 각 단계의 근거

| 명령 | 확인 범위 |
| --- | --- |
| `python tools/build.py --check` | 조립기에 명시된 소스 조각에서 조립한 ASCII/CRLF BAS가 배포 파일과 바이트 단위로 같은지 확인 |
| `python tools/validate.py` | VBA7 32비트·64비트 및 Legacy 32비트 조건 분기를 각각 확인 |
| `python tools/release_evidence.py` | 현재 `docs/validation.json`의 BAS 해시·바이트·버전·행 수·진입점과 알려진 검증 결과의 대상 해시가 현재 배포 파일과 일치하는지 확인 |
| `python -B -m unittest discover -s development -p test_validation_tools.py` | 일부러 잘못된 공개 매크로·참조·선언·조건 분기·인코딩 등을 넣었을 때 검증이 거부하는지 확인 |
| `python tools/test.py` | 각 백업 모델의 실제 실행 건수·종료 코드·정상 종료 문구 확인; 건너뛴 검사와 expected failure를 통과로 세지 않음 |

정적 검사는 모듈 이름, `Option Explicit`, 공개 진입점 하나, CRLF·ASCII, 물리적 줄 길이와 연속 줄 연결 수, 프로시저 시작·종료, 중복 모듈 심볼, 선언 위치, 프로젝트 이름 규칙의 미정의 심볼, 강제 `End`를 확인합니다. 각 조건 분기에서 가장 큰 프로시저의 **소스 크기**도 보고합니다. 소스 크기는 VBA 컴파일 후의 코드 크기와 같지 않으므로 Office의 프로시저 크기 제한을 만족했다는 증거로 사용하지 않습니다.

조건 전처리기는 현재 소스가 쓰는 `#If`·`#ElseIf`·`#Else`·`#End If`, `VBA7`·`Win64`·`Win32`·`Mac`, `True`·`False`·`0`·`1`, 괄호와 `And`·`Or`·`Not`만 지원합니다. 새 조건 문법이나 알 수 없는 기호는 추측해서 통과시키지 않고 실패합니다. 조건부 선언의 모든 지원 분기를 별도로 검사합니다.

## 기계 판독 결과

현재 배포 BAS의 SHA-256에 묶인 결과를 JSON으로 받을 수 있습니다.

```sh
python tools/validate.py --json
python tools/test.py --json
```

이 명령은 결과를 표준 출력으로 보냅니다. 별도의 보고서를 보존하려면 실행한 커밋과 함께 저장합니다. 이전 버전의 기록을 새 버전의 실행 결과로 사용하지 않습니다. 현재 배포 보고서는 `docs/validation.json`에 두고 이전 결과는 별도 역사 기록으로 보존합니다. 증거 확인은 알려진 BAS 검사 객체 및 명시적인 `artifact_sha256` 참조를 확인하며, 다른 원본 파일의 해시를 BAS 해시와 같다고 요구하지 않습니다.

## GitHub Actions

워크플로는 push·pull request·수동 실행에서 Ubuntu와 Windows, Python 3.10과 3.14를 조합해 같은 `verify.py`를 실행하도록 구성했습니다. 이 조합은 지원하는 최소 Python과 더 최신 Python, 서로 다른 파일시스템·개행 동작을 확인하기 위한 것입니다. 실제 통과 여부는 해당 커밋의 Actions 결과를 확인합니다.

토큰 권한은 `contents: read`이며, checkout의 자격 증명 지속 저장을 끕니다. 공식 저장소에서 확인한 릴리스 커밋으로 고정했습니다.

- [actions/checkout v7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1): `3d3c42e5aac5ba805825da76410c181273ba90b1`
- [actions/setup-python v7.1.0](https://github.com/actions/setup-python/releases/tag/v7.1.0): `6d4fd22b20465ccc0f7354476588e8df7c9ca0e0`
- [GitHub 공식 보안 지침](https://docs.github.com/en/actions/reference/security/secure-use): 권한 최소화와 고정 참조의 근거

## 확인하지 않는 항목

이 도구는 VBA 컴파일러나 완전한 VBA 구문 분석기가 아닙니다. 모든 지역 변수·형식·인수의 정확성, COM·Outlook MAPI 동작, 파일 잠금·전원 손실·레지스트리 동작, 실제 기존 MSG 호환성과 10만 건 처리 시간을 인증하지 않습니다. GitHub Windows runner에서 Python 검사가 통과해도 Outlook 자체가 실행된 것은 아닙니다.

이전 작성 단계의 ANTLR 결과는 이전 보고서의 근거로 남습니다. 여기의 정적 도구는 그 파서를 실행하지 않으며, 제삼자 ANTLR 런타임이나 생성된 파서 코드를 포함하지 않습니다. 실제 대상 Windows PC의 클래식 Outlook에서 BAS 가져오기·VBA 컴파일·임시 환경 검사·대표 메일 검증을 완료해야 운영 검증이 닫힙니다.
