import re
from pathlib import Path

from deepagents.backends import LocalShellBackend


def get_safe_path(user_input_path: str, base_dir: str) -> Path:
    base = Path(base_dir).resolve()
    target = (base / user_input_path).resolve()
    if not target.is_relative_to(base):
        raise ValueError("Unauthorized path traversal attempt detected.")
    return target


def get_safe_pdf(user_input_path: str, base_dir: str = "inputs", max_mb: int = 50) -> Path:
    if Path(user_input_path).is_absolute():
        raise ValueError("Absolute paths are not allowed.")
    target = get_safe_path(user_input_path, base_dir)
    if target.suffix.lower() != ".pdf" or not target.is_file():
        raise ValueError("Invalid or missing PDF file.")
    if target.stat().st_size > max_mb * 1024 * 1024:
        raise ValueError(f"PDF is larger than {max_mb} MB.")
    return target


def get_safe_pptx(user_input_path: str, base_dir: str = "outputs") -> Path:
    if Path(user_input_path).is_absolute():
        raise ValueError("Absolute paths are not allowed.")
    target = get_safe_path(user_input_path, base_dir)
    if target.suffix.lower() != ".pptx":
        raise ValueError("Output must be a .pptx file.")
    return target

PDF = r'"?[\w\-./]+\.pdf"?'
PPTX = r'"?[\w\-./]+\.pptx"?'
PP = r'"?skills/pdf-pptx-skill/scripts/process_pdf\.py"?'
CK = r'"?skills/pdf-pptx-skill/scripts/check_pptx\.py"?'
ALLOWED = re.compile(
    rf'^python {PP} {PDF} {PPTX}'
    r'( --extract-only| --summarize)?$'
    rf'|^python {CK} {PPTX}$'
)
FORBIDDEN_CHARS = set(";&|`$<>\n\r\\")

def reject(reason, cmd):
    return f"Command not allowed [{reason}]: {cmd!r}"

class RestrictedShellBackend(LocalShellBackend):
    def execute(self, command: str, *args, **kwargs):
        cmd = command.strip()
        if ".." in cmd:
            raise PermissionError("Path traversal not allowed")
        if not ALLOWED.match(cmd) or any(c in cmd for c in FORBIDDEN_CHARS):
            raise PermissionError(f"Command not allowed: {cmd[:80]}")
        return super().execute(command, *args, **kwargs)