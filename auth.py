import hashlib
import secrets


def hash_password(password: str, salt: bytes = None):
    if salt is None:
        salt = secrets.token_bytes(16)

    pwd_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return salt, pwd_hash


def verify_password(password: str, salt: bytes, pwd_hash: bytes):
    check = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return secrets.compare_digest(check, pwd_hash)