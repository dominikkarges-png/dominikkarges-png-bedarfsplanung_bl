"""
Erzeugt einen Passwort-Hash fuer die Secrets der Absatzplanung.

AUSFUEHREN (VS Code):  Datei oeffnen, "Run Python File", Name, Rolle und Passwort eingeben.
Ausgabe in die Secrets kopieren (Streamlit Cloud: App, Settings, Secrets).
Das Passwort selbst wird nirgends gespeichert.
"""
import getpass
import hashlib
import secrets

RUNDEN = 600_000

name = input("Name (wie bei der Anmeldung, z. B. D. Karges): ").strip()
rolle = input("Rolle [mb / koordination]: ").strip().lower() or "mb"
while True:
    pw = getpass.getpass("Passwort (mind. 12 Zeichen): ")
    if len(pw) < 12:
        print("Zu kurz, bitte mindestens 12 Zeichen.")
        continue
    if pw != getpass.getpass("Passwort wiederholen: "):
        print("Stimmt nicht ueberein.")
        continue
    break
salt = secrets.token_bytes(16)
h = hashlib.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt, RUNDEN).hex()
print("\nIn die Secrets kopieren:\n")
print(f'[nutzer."{name}"]')
print(f'passwort_hash = "pbkdf2_sha256${RUNDEN}${salt.hex()}${h}"')
print(f'rolle = "{"koordination" if rolle.startswith("k") else "mb"}"\n')
