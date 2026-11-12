# v3.5.2 검증 결과

**메일 모델·소스 계약 610/610개와 검증기 결함 주입 47/47개, 총 657개가 로컬에서 통과했습니다.** 검사 제외나 건너뛰기를 통과에 포함하지 않았습니다. 이는 실제 Outlook 실행 결과와 구분합니다.

## 검증 대상

| 항목 | 값 |
| --- | --- |
| 배포 파일 | `OutlookMailBackup.bas` |
| 버전 | 3.5.2 |
| 바이트 | 749,746 |
| 물리적 줄 수 | 14,765 |
| 인코딩·줄바꿈 | ASCII / CRLF |
| SHA-256 | `7e8d6ee1db2d0067244fd384d42f3dce1d2f50275a664d713f5789540b1127d8` |
| 공개 실행 매크로 | `MB_01_RunBackup` 하나 |
| 로컬 검증 환경 | Linux / Python 3.12.14 |

같은 소스에서 BAS를 재생성해 바이트 단위 일치를 확인했습니다. [검증 JSON](validation.json)은 이 BAS와 결과를 연결합니다. `tools/release_evidence.py`는 버전·해시·바이트·행 수와 알려진 BAS 검사 증거가 현재 파일과 맞는지 확인합니다. 이전 [3.4.0 결과](history/validation-v3.4.0.json), [3.5.0 결과](history/validation-v3.5.0.json), [3.5.1 결과](history/validation-v3.5.1.json)는 별도 기록에 보존했습니다.

## 메일 모델·소스 계약

| 검사 모음 | 건수 | 범위 | 결과 |
| --- | ---: | --- | --- |
| `v34_settings_models.py` | 43 | 설정·위치 재발견 | PASS |
| `v34_revision_models.py` | 35 | 내용 변경 리비전 | PASS |
| `v34_metrics_models.py` | 27 | 성능 표본·예측 | PASS |
| `v34_fidelity_models.py` | 28 | 비교 범위·제한 | PASS |
| `v34_recovery_merge_models.py` | 44 | 복구·잔존 기록 병합 | PASS |
| `v34_recovery_proof_models.py` | 41 | 과거 증거 보존 | PASS |
| `v34_policy_upgrade_models.py` | 25 | 이전 정책·형식 읽기 | PASS |
| `v341_path_models.py` | 16 | 연결 경로·쓰기 경계 | PASS |
| `v341_index_alias_models.py` | 41 | 인덱스 충돌·대량 별칭 | PASS |
| `v35_layout_models.py` | 32 | 새 폴더명·삭제 후 재사용 | PASS |
| `v35_relative_models.py` | 48 | 상대 경로·혼합 저널·네이티브 연결 | PASS |
| `v351_recovery_continuity_models.py` | 39 | 과거 prefix·snapshot 고정·merge checkpoint | PASS |
| `v351_layout_binding_models.py` | 57 | 명시 구조 선택·증거 생존·충돌·재시작 | PASS |
| `v351_review_guidance_models.py` | 46 | 원인 분류·보고서 게시·메인 통합·해소 안내 | PASS |
| `v351_selection_models.py` | 36 | 개별 REVIEW 의무·증거 hash·cache 선별·오류 격리 | PASS |
| `test_hresult_contract.py` | 21 | CVErr 범위·32/64비트 SCODE·접근 오류 구분 | PASS |
| `v352_startup_models.py` | 31 | 실제 재시도 판정식·취소·정리·경로 대체·진단·공개 호출 흐름 | PASS |
| **합계** | **610** | Python 모델 및 실제 소스 계약 | **PASS** |

이전 상태 모델과 새 구현의 후보 선택 결과를 교차 비교했고, 충돌 전후의 저널·캐시 불변성, 부분 복구와 재개, 경로 경계와 혼합 형식을 검증했습니다. 소스의 검사 순서와 호출 계약도 확인합니다. Python 모델은 COM·MAPI 호출의 실행을 대신하지 않습니다.

## 검증 도구 자체 검사

47개 결함 주입 검사에서 잘못된 공개 매크로, 미정의 프로젝트 심볼, 중복 선언, 잘못된 조건 분기, BOM·줄바꿈, 변조된 BAS, 다른 버전의 보고서, skipped·expected-failure 출력을 거부했습니다. `verify.py`는 실제 검사 수와 정상 종료를 함께 확인합니다.

## 현재 BAS 구조 검사

| 조건 분기 | 프로시저 | 공개 진입점 | 검출된 구조 오류 |
| --- | ---: | ---: | ---: |
| VBA7 / Windows 32비트 | 537 | 1 | 0 |
| VBA7 / Windows 64비트 | 537 | 1 | 0 |
| Legacy VBA / Windows 32비트 | 537 | 1 | 0 |

ASCII/CRLF, 선언·프로시저 배치, 프로젝트 이름 규칙의 심볼 연결, 줄 길이·연속 줄 연결 제한 등을 검사했습니다. **Office VBA 컴파일이나 모든 형식·인수·COM 호출을 확인한 것은 아닙니다.** 자세한 도구 범위는 [검증 도구 설명](VALIDATION_TOOLING.md)에 있습니다.

## 추가 문법 검사

별도 ANTLR 파서로 새 프로시저 16개, 변경 프로시저 11개와 전체 선언부를 검사했습니다. VBA7 32비트·64비트 및 Legacy 32비트의 각 28단위에서 문법 오류 0건입니다. 변경하지 않은 전체 프로시저를 다시 파싱한 결과는 아닙니다. 이전 v3.5.1의 전체 522단위 검사 결과는 해당 버전의 [역사 기록](history/validation-v3.5.1.json)에 보존했습니다.

이 검사는 Office 컴파일·형식 검사·실행을 대신하지 않습니다. 외부 파서와 런타임은 저장소 기본 검증 명령에 포함되지 않습니다. 정확한 대상 해시·범위·분기별 결과는 [검증 JSON](validation.json)에 기록했습니다.

## 성능 근거

같은 제목 100,000개의 독립 Python 모델에서 서로 다른 ID 또는 ID 없는 서로 다른 발송 시각은 신규 후보만 분류하도록 검증했습니다. 동일한 물리 파일의 재가져오기 별칭 100,000개는 헤더 읽기 1회, 후보 방문 100,000회로 모델링했습니다. 이는 이전의 반복 전체 탐색을 줄이는 알고리즘 근거이며 **Outlook 처리 시간 실측이 아닙니다.** 실제 100,000메일 실행은 아직 NOT_RUN입니다.

추가된 보조 선별 캐시는 100,000개 합성 기록에서 이전 성공 범위에만 해당하는 이미 검증된 메일의 반복 상세 비교 감소를 검증합니다. pending·unknown·changed·policy·audit 선택은 유지합니다. 전수 메타데이터 순회와 snapshot prefix 해시 비용은 별도로 남습니다.

## 원격 CI

워크플로는 Ubuntu/Windows × Python 3.10/3.14의 4개 조합에서 같은 `python -B tools/verify.py`를 실행합니다. 원격 결과는 [GitHub Actions](https://github.com/greatgc-flow/OutlookMailBackup/actions)에서 해당 커밋에 연결된 실행으로 확인합니다. 구성 파일이 있다는 사실만으로 CI 통과라고 표시하지 않습니다. Windows의 Python 검사와 Outlook 실행은 서로 다른 검증입니다.

## 대상 PC 실행 상태

v3.5.1 배포 후 사용자가 `fail; failed checks=1, remain NOT_RUN`을 보고했습니다(`USER_REPORTED`). 실제 가져온 모듈 버전·실패 검사명·상세 보고서는 확보하지 못했습니다. 이 보고와 별도로 음수 HRESULT를 CVErr에 전달하는 코드 결함을 확인하여 수정했습니다. 해당 결함이 사용자 실행의 직접 원인이었는지는 미확인입니다.

다음 항목은 **수정한 v3.5.2에 대해 NOT_RUN**입니다.

- 대상 PC의 BAS 가져오기와 `Debug > Compile VBAProject`.
- 실제 Outlook/MAPI·기존 MSG·특수 첨부파일과의 호환성.
- 첫 실행에 포함된 상대 경로·이동·메타데이터 복구·합성 MSG 네이티브 검사.
- 실제 Windows junction, 파일 가드, 레지스트리와 파일시스템 장애.
- 실제 10만 건의 전체 소요 시간과 실제 전원 손실·하드웨어 내구성.

모든 임의 MAPI 속성의 동일성은 인증 범위 밖입니다. 이 항목들은 [실제 Outlook 검증 절차](OUTLOOK_VALIDATION.md)에 따라 결과를 남겨야 합니다.

## 평가의 해석

설계와 자동 검증의 완결성 평가는 실행 증거와 별개입니다. 자동 검사 통과율은 100%지만, 이를 운영 품질 100점 또는 모든 과거 출력과의 100% 상호 쓰기 호환성으로 바꾸지 않습니다. 상대 경로를 사용해도 소유권·이동·복구의 절대 경로 바인딩은 남습니다. 새 운영 기록을 이전 작성기로 다시 쓰는 동작은 지원하지 않습니다.

[v3.5.2 시작 오류 수정·자동 복구](STARTUP_RECOVERY_20261112.md) · [v3.5.1 당시 수정과 점수](REMEDIATION_20261111.md) · [시나리오별 계약](SCENARIOS.md) · [폴더와 경로](PATHS.md) · [설치와 사용](../README.md)
