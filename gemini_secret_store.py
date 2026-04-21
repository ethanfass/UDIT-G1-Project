import base64
import ctypes
import json
import os
from ctypes import wintypes
from pathlib import Path
from typing import Optional, Tuple


APP_NAME = "UDIT-G1-Project"
SECRET_DIR = Path(__file__).resolve().parent / ".local_secrets"
SECRET_PATH = SECRET_DIR / "gemini_api_key.json"
DESCRIPTION = f"{APP_NAME} Gemini API Key"


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


crypt32 = ctypes.windll.crypt32
kernel32 = ctypes.windll.kernel32


def _blob_from_bytes(data: bytes) -> DATA_BLOB:
    if not data:
        return DATA_BLOB(0, None)
    buffer = ctypes.create_string_buffer(data)
    return DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))


def _bytes_from_blob(blob: DATA_BLOB) -> bytes:
    if not blob.cbData:
        return b""
    pointer = ctypes.cast(blob.pbData, ctypes.POINTER(ctypes.c_char))
    return ctypes.string_at(pointer, blob.cbData)


def _protect_data(plaintext: str) -> str:
    input_blob = _blob_from_bytes(plaintext.encode("utf-8"))
    output_blob = DATA_BLOB()
    if not crypt32.CryptProtectData(
        ctypes.byref(input_blob),
        DESCRIPTION,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise OSError("Windows DPAPI encryption failed.")

    try:
        encrypted = _bytes_from_blob(output_blob)
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)

    return base64.b64encode(encrypted).decode("ascii")


def _unprotect_data(ciphertext_b64: str) -> str:
    encrypted = base64.b64decode(ciphertext_b64.encode("ascii"))
    input_blob = _blob_from_bytes(encrypted)
    output_blob = DATA_BLOB()
    if not crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,
        None,
        None,
        None,
        0,
        ctypes.byref(output_blob),
    ):
        raise OSError("Windows DPAPI decryption failed.")

    try:
        decrypted = _bytes_from_blob(output_blob)
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)

    return decrypted.decode("utf-8")


def save_local_api_key(api_key: str) -> Path:
    api_key = api_key.strip()
    if not api_key:
        raise ValueError("API key cannot be empty.")

    SECRET_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "app": APP_NAME,
        "description": DESCRIPTION,
        "encrypted_api_key": _protect_data(api_key),
    }
    SECRET_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return SECRET_PATH


def load_local_api_key() -> Optional[str]:
    if not SECRET_PATH.exists():
        return None

    payload = json.loads(SECRET_PATH.read_text(encoding="utf-8"))
    encrypted = payload.get("encrypted_api_key", "")
    if not encrypted:
        return None
    return _unprotect_data(encrypted)


def clear_local_api_key() -> bool:
    if SECRET_PATH.exists():
        SECRET_PATH.unlink()
        return True
    return False


def local_api_key_exists() -> bool:
    return SECRET_PATH.exists()


def resolve_api_key() -> Optional[str]:
    env_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if env_key:
        return env_key.strip()
    return load_local_api_key()


def resolve_api_key_with_source() -> Tuple[Optional[str], str]:
    env_key = os.getenv("GEMINI_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip(), "env:GEMINI_API_KEY"

    google_env_key = os.getenv("GOOGLE_API_KEY")
    if google_env_key and google_env_key.strip():
        return google_env_key.strip(), "env:GOOGLE_API_KEY"

    local_key = load_local_api_key()
    if local_key:
        return local_key, f"local:{SECRET_PATH}"

    return None, "none"


def mask_api_key(api_key: Optional[str]) -> str:
    if not api_key:
        return "<missing>"
    api_key = api_key.strip()
    if len(api_key) <= 8:
        return "*" * len(api_key)
    return f"{api_key[:4]}...{api_key[-4:]}"
