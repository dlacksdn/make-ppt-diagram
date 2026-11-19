# make-ppt-diagram

논문 방법론 다이어그램을 **RILS LAB 세미나 템플릿** 위에 그리는 Claude 스킬입니다. 두 모드가 있고, 요청한 모드만 실행합니다.

| 모드 | 호출 예 | 입력 → 출력 |
|---|---|---|
| 🅐 그리기 | "방법론 그려줘", "도식화해줘" | 논문 + 공식 코드 → phase 분할 합의 → **PNG** |
| 🅑 PPT | "PPT로 만들어줘" | 🅐의 그림 또는 손그림 → **편집 가능한 .pptx** |

두 모드 모두 같은 kit(`pptx_kit.py`)로 그리므로, 🅐에서 그린 그림은 손실 없이 .pptx가 됩니다.
스타일 수치(pt, 선 두께, 간격, 색 역할)는 `pptx_kit.py`에 상수로 들어 있고, 절차와 규칙은 `SKILL.md`에 있습니다.

## 설치

```bash
git clone <this-repo> make-ppt-diagram
cd make-ppt-diagram
pip install -r requirements.txt
```

렌더 검수(PPT → PNG)에는 [LibreOffice](https://www.libreoffice.org/)가 필요합니다. `render()`가 Windows·macOS·Linux 설치 경로를 자동으로 찾습니다.

### Claude Code에서 스킬로 쓰기
폴더째 스킬 디렉터리에 두면 됩니다. 폴더 이름은 `SKILL.md`의 `name`과 같게 맞춥니다.

```
~/.claude/skills/make-ppt-diagram/      # 개인 스킬
<project>/.claude/skills/make-ppt-diagram/  # 프로젝트 스킬
```

### 동작 확인

```bash
python demo_diwa_reward.py /tmp/demo.pptx
python -c "import pptx_kit; print(pptx_kit.render('/tmp/demo.pptx'))"
```

## 구성

| 파일 | 내용 |
|---|---|
| `SKILL.md` | 스킬 설명서: 작업 순서, 고정 규칙, 의미 어휘, 레이아웃, 체크리스트 |
| `pptx_kit.py` | 다이어그램 키트: `open_deck`, `box`, `arrow`, `panel`, `legend`, `render`, 레퍼런스용 `list_refs`·`add_refs`, 기존 파일용 `copy_diagram`·`safe_save`·`fingerprint` 등 |
| `base.pptx` | 템플릿 6페이지만 남긴 베이스 |
| `refs/` | 레퍼런스: `kit_*.pptx`(kit 규칙 준수, 수치까지 참고) · `legacy_seminars.pptx`(구성만 참고). 발표 1개 = 파일 1개 |
| `assets/` | 🔥❄️👁️ 아이콘(SVG/PNG), 글자 폭 측정용 폰트 |

## 라이선스 참고
- `assets/fonts/`: Carlito, Caladea — SIL Open Font License 1.1 (각 `LICENSE-*.txt`)
- `base.pptx`, `refs/`: 연구실 템플릿·로고와 세미나 자료가 들어 있습니다.
