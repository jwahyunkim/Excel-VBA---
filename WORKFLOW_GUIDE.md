# Git 버전 개발 및 릴리즈 사용 설명서

이 문서는 `version-workflow.ps1`과 npm 명령을 이용한 버전 개발, 작업 브랜치, 릴리즈 방법을 설명합니다.

## 가장 간단한 사용법

명령을 외우기 어렵다면 다음 명령 하나만 실행하고 메뉴 번호를 선택합니다.

```powershell
npm run workflow
```

직접 실행할 때 사용하는 명령은 다음과 같습니다.

| 작업 | 명령 |
|---|---|
| 새 버전 개발 시작 | `npm run version:start` |
| 자식 작업 브랜치 생성 | `npm run branch:create -- feature/브랜치명` |
| 현재 자식 브랜치 병합 | `npm run branch:merge` |
| 개발 완료 및 자동 릴리즈 | `npm run version:release` |
| 현재 버전 개발 취소 | `npm run version:cancel` |
| 상태 확인 | `npm run workflow:status` |

## 브랜치 구조

한 버전의 작업 이력은 다음 단계로 보존됩니다.

```text
main
└─ release/v3.3.0       릴리즈 준비
   └─ develop/v3.3.0    버전 통합 개발
      └─ feature/...    개별 기능 또는 수정
```

각 브랜치의 역할은 다음과 같습니다.

- `main`: 배포 완료 버전과 `vA.B.C` 태그가 남는 기준 브랜치
- `develop/vA.B.C`: 해당 버전의 여러 작업을 통합하는 개발 브랜치
- `release/vA.B.C`: 보안 배포본 생성·검증과 최종 배포를 준비하는 브랜치
- `feature/*`, `fix/*`: 실제 기능 개발과 오류 수정을 진행하는 자식 브랜치

`release` 브랜치는 개발 시작 시 미리 만들지 않습니다. `npm run version:release`를 실행할 때 자동 생성한 뒤 `develop → release → main` 순서로 일반 merge합니다.

## 버전 번호 선택

버전은 `vA.B.C` 형식을 사용합니다.

```text
v3.2.12
 │ │ └─ C: patch
 │ └─── B: minor
 └───── A: major
```

현재 버전이 `v3.2.12`일 때 결과는 다음과 같습니다.

| 선택 | 용도 | 새 버전 | 초기화 규칙 |
|---|---|---|---|
| `patch` | 작은 오류 수정 | `v3.2.13` | 없음 |
| `minor` | 호환되는 기능 추가 | `v3.3.0` | C를 `0`으로 초기화 |
| `major` | 큰 변경 또는 호환성 변경 | `v4.0.0` | B와 C를 `0`으로 초기화 |

다음 명령을 실행하면 최신 Git 태그를 확인하고 세 가지 결과를 보여줍니다.

```powershell
npm run version:start
```

```text
현재 버전: v3.2.12
[1] patch (c 증가) : v3.2.13
[2] minor (b 증가) : v3.3.0  - c는 0으로 초기화
[3] major (a 증가) : v4.0.0  - b와 c는 0으로 초기화
올릴 버전 단위 선택:
```

단위를 명령에 바로 넣어도 됩니다.

```powershell
npm run version:start -- patch
npm run version:start -- minor
npm run version:start -- major
```

버전 숫자는 직접 수정하지 않습니다. 최신 원격 태그를 기준으로 자동 계산됩니다.

## 전체 작업 순서

### 1. 새 버전 개발 시작

```powershell
npm run version:start
```

예를 들어 현재 태그가 `v3.2.12`이고 `minor`를 선택하면 최신 `main`에서 다음 브랜치를 생성하고 이동합니다.

```text
develop/v3.3.0
```

브랜치를 만든 직후 개발 원본 Excel 파일명과 `config.json`의 `excel_file`도 `v3.3.0`으로 함께 변경되고, `vba-files`의 보안 VBA 소스가 새 개발본 내부에 자동 동기화됩니다. 이 변경은 자동 스테이징하지 않으므로 개발 작업과 함께 원하는 롤백 단위로 커밋합니다.

처음 워크플로 파일을 추가한 상태라면 이 브랜치에서 함께 커밋합니다.

```powershell
git add package.json version-workflow.ps1 WORKFLOW_GUIDE.md
git commit -m "Git 버전 워크플로 추가"
git push
```

### 2. 자식 작업 브랜치 생성

현재 `develop/v3.3.0`에서 작업 브랜치를 만듭니다.

```powershell
npm run branch:create -- feature/report-export
```

현재 브랜치가 부모로 자동 지정되고 기록됩니다.

브랜치 이름을 생략하면 실행 중에 물어봅니다.

```powershell
npm run branch:create
```

```text
새 자식 브랜치 이름: feature/report-export
```

### 3. 개발 작업

개발 중간 커밋은 필요할 때 직접 만들 수 있습니다. 다만 최종 완료 시점의 스테이징과 커밋은 프로젝트 워크플로가 자동 처리하므로 필수는 아닙니다.

### 4. 자식 브랜치를 develop로 병합

작업 브랜치에서 실행합니다.

```powershell
npm run branch:merge
```

스크립트가 다음 작업을 처리합니다.

1. 현재 작업 브랜치의 모든 변경을 자동 스테이징·커밋
2. 현재 작업 브랜치 푸시
3. 기록된 부모 `develop/vA.B.C`를 대상으로 PR 생성
4. 일반 merge 방식으로 병합
5. 작업 브랜치 삭제
6. 부모 develop 브랜치로 이동

다른 PC에서 받은 브랜치처럼 부모 기록이 없으면 병합 대상 브랜치 이름을 물어봅니다.

필요한 경우 부모를 직접 지정할 수도 있습니다.

```powershell
npm run branch:merge -- -SourceBranch feature/report-export -TargetBranch develop/v3.3.0
```

### 5. 개발 완료 및 자동 릴리즈

모든 작업 브랜치를 병합한 뒤 `develop/vA.B.C`에서 실행합니다.

```powershell
npm run version:release
```

스크립트가 다음 순서로 처리합니다.

1. `develop/vA.B.C`의 모든 남은 변경을 자동 스테이징·커밋하고 푸시
2. `main`에서 `release/vA.B.C` 생성
3. `develop/vA.B.C → release/vA.B.C` PR 생성 및 일반 merge
4. 보안 배포본 생성 및 임시 복사본 자동 검증
5. 개발 원본, config, 배포본을 릴리즈 준비 커밋으로 반영
6. `release/vA.B.C → main` PR 생성 및 일반 merge
7. 임시 develop/release 브랜치 삭제
8. `main`에 `vA.B.C` 태그 생성 및 푸시

이 자동 스테이징·커밋은 Codex가 임의로 실행하는 것이 아니라 사용자가 `npm run version:release` 또는 메뉴의 완료 항목을 선택했을 때 프로젝트 스크립트가 수행합니다.

배포 보안 설정을 실제로 수정하면 작업 성공 직후 이 릴리즈 절차를 자동 호출합니다. 별도의 VS Code 커밋이나 `npm run version:release` 실행은 필요하지 않습니다. 배포본 생성은 현재 버전의 파일만 만들며 버전을 올리거나 자동 릴리즈하지 않습니다. 개발본 동기화, 일괄 실행과 읽기 전용 항목도 자동 릴리즈하지 않습니다.

### 6. 새 버전 개발 취소

릴리즈 전 `develop/vA.B.C`에서 다음 명령을 실행합니다.

```powershell
npm run version:cancel
```

화면에 표시된 개발 브랜치 이름을 다시 입력하면 다음 순서로 취소합니다.

1. 미커밋 변경과 추적되지 않은 파일을 Git stash에 백업
2. 부모 브랜치에 없는 개발 커밋을 `cancel-backup/vA.B.C-날짜시간` 로컬 태그로 백업
3. 원래 부모 브랜치(기본값 `main`)로 복귀
4. 원격 및 로컬 `develop/vA.B.C` 브랜치 삭제

자식 브랜치나 `release/vA.B.C`, 동일 버전 태그가 남아 있으면 안전을 위해 취소하지 않습니다. 완료 화면에 백업 태그와 stash 복구 명령이 표시됩니다. 백업 태그는 원격으로 자동 푸시하지 않습니다.

예를 들어 기존 파일명이 다음과 같다면:

```text
업무 간트 v4.0.0.xlsm
```

`v4.0.1` 개발을 시작하면 자동으로 다음처럼 변경됩니다.

```text
업무 간트 v4.0.1.xlsm
```

## 브랜치 이름 규칙

영문 소문자와 하이픈 사용을 권장합니다.

| 작업 종류 | 형식 | 예시 |
|---|---|---|
| 신규 기능 | `feature/설명` | `feature/report-export` |
| 오류 수정 | `fix/설명` | `fix/date-calculation` |
| 코드 정리 | `refactor/설명` | `refactor/gantt-module` |
| 문서·설정 | `chore/설명` | `chore/update-guide` |

공백은 사용할 수 없습니다.

```text
feature/report export  사용 불가
feature/report-export  권장
```

## 다른 기준 브랜치 사용

기본 기준 브랜치는 `main`입니다. 다른 기준 브랜치에서 시작해야 한다면 다음처럼 지정할 수 있습니다.

```powershell
npm run version:start -- minor -BaseBranch maintenance
```

일반적인 릴리즈에서는 지정할 필요가 없습니다.

## 실행 전 확인

브랜치 생성 전에는 기존 작업을 정리해야 합니다. 자식 병합과 릴리즈를 선택하면 해당 브랜치의 변경은 프로젝트 워크플로가 자동 스테이징·커밋합니다.

```powershell
npm run workflow:status
```

다음 표시가 있으면 먼저 커밋하거나 변경을 정리합니다.

```text
 M 수정된파일
?? 새파일
```

현재 브랜치도 확인합니다.

```powershell
git branch --show-current
```

`version:release`는 반드시 릴리즈할 `develop/vA.B.C`에서 실행해야 합니다. `feature/*` 브랜치에서 바로 실행하지 않습니다.

## 필요한 프로그램

- Git
- Node.js와 npm
- Windows PowerShell
- GitHub CLI(`gh`)

GitHub CLI는 PR을 생성하고 병합할 권한이 있는 계정으로 로그인되어 있어야 합니다.

```powershell
gh auth status
```

## 자주 발생하는 오류

### 커밋되지 않은 변경 사항이 있습니다

변경 내용을 커밋한 후 다시 실행합니다.

```powershell
git add .
git commit -m "작업 내용"
git push
```

### 부모 브랜치를 찾을 수 없습니다

브랜치 이름을 확인합니다.

```powershell
git branch --all
```

### GitHub CLI를 찾을 수 없습니다

GitHub CLI 설치와 로그인 상태를 확인합니다.

```powershell
gh auth status
```

### 태그가 이미 존재합니다

동일한 버전이 이미 릴리즈된 상태입니다. 최신 태그를 확인하고 새 버전 개발을 다시 시작합니다.

```powershell
git tag --list --sort=-version:refname
```

## 자주 쓰는 네 명령

```powershell
# 버전 단위 선택 후 develop/vA.B.C 생성
npm run version:start

# 자식 작업 브랜치 생성
npm run branch:create -- feature/report-export

# 자식 브랜치를 develop로 병합
npm run branch:merge

# develop → release → main 릴리즈
npm run version:release
```
