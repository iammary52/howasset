"""
Claude Code 루틴 설치 스크립트

사용법:
  python setup_routines.py

1. secrets.json 에 자격증명 입력 (secrets.template.json 참고)
2. python setup_routines.py 실행
3. Claude Code 재시작 -> 사이드바 Routines 에서 확인
"""

import json
import shutil
from pathlib import Path

ROUTINES_DIR = Path(__file__).parent / "routines"
TARGET_DIR   = Path.home() / ".claude" / "scheduled-tasks"
SECRETS_PATH = Path(__file__).parent / "secrets.json"
TEMPLATE_PATH = Path(__file__).parent / "secrets.template.json"


def load_secrets() -> dict:
    if SECRETS_PATH.exists():
        with open(SECRETS_PATH, encoding="utf-8") as f:
            return json.load(f)

    print("secrets.json 이 없습니다. 자격증명을 입력해 주세요.")
    print("(입력 내용은 secrets.json 에 저장되며 GitHub 에는 올라가지 않습니다)\n")
    secrets = {
        "GITHUB_PAT":        input("GitHub PAT (ghp_...): ").strip(),
        "GMAIL_ADDRESS":     input("Gmail 주소:           ").strip(),
        "GMAIL_APP_PASSWORD": input("Gmail 앱 비밀번호:    ").strip(),
    }
    with open(SECRETS_PATH, "w", encoding="utf-8") as f:
        json.dump(secrets, f, ensure_ascii=False, indent=2)
    print(f"\nsecrets.json 저장 완료: {SECRETS_PATH}\n")
    return secrets


def inject_secrets(text: str, secrets: dict) -> str:
    for key, value in secrets.items():
        text = text.replace(f"{{{{{key}}}}}", value)
    return text


def install():
    secrets = load_secrets()

    if not ROUTINES_DIR.exists():
        print("오류: routines/ 폴더가 없습니다.")
        return

    installed = []
    for skill_dir in sorted(ROUTINES_DIR.iterdir()):
        if not skill_dir.is_dir():
            continue
        skill_md = skill_dir / "SKILL.md"
        if not skill_md.exists():
            continue

        content = skill_md.read_text(encoding="utf-8")
        content = inject_secrets(content, secrets)

        target = TARGET_DIR / skill_dir.name
        target.mkdir(parents=True, exist_ok=True)
        (target / "SKILL.md").write_text(content, encoding="utf-8")
        installed.append(skill_dir.name)
        print(f"  [OK] {skill_dir.name}")

    if installed:
        print(f"\n{len(installed)}개 루틴 설치 완료!")
        print("Claude Code 를 재시작하면 사이드바 Routines 에서 확인할 수 있습니다.")
    else:
        print("설치할 루틴이 없습니다.")


if __name__ == "__main__":
    print(f"루틴 설치 위치: {TARGET_DIR}\n")
    install()
