# make-ppt-diagram

논문 방법론 그림(Claude가 그린 시각화 또는 태블릿 손그림)을 **RILS LAB 세미나 템플릿** 위의 **편집 가능한 PowerPoint 다이어그램**으로 옮기는 Claude 스킬입니다.
스타일 수치(pt, 선 두께, 간격, 색 역할)는 `pptx_kit.py`에 상수로 들어 있습니다. 규칙과 작업 순서는 `SKILL.md`에 있습니다.

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
| `pptx_kit.py` | 다이어그램 키트: `open_deck`, `box`, `arrow`, `panel`, `legend`, `render`, `list_refs` 등 |
| `base.pptx` | 템플릿 6페이지만 남긴 베이스 |
| `refs.pptx` | 레퍼런스 그림 13장 (노트에 출처·레이아웃·평가) |
| `assets/` | 🔥❄️👁️ 아이콘(SVG/PNG), 글자 폭 측정용 폰트 |

## 라이선스 참고
- `assets/fonts/`: Carlito, Caladea — SIL Open Font License 1.1 (각 `LICENSE-*.txt`)
- `base.pptx`, `refs.pptx`: 연구실 템플릿·로고와 세미나 자료가 들어 있습니다. 공개 저장소로 올리기 전에 확인이 필요합니다.
