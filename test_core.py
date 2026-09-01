import os
from pathlib import Path

from core.profile import load_profile

base = Path(__file__).resolve().parent

profile_path = Path(
    os.environ.get(
        "BHEAD_M365_PROFILE",
        str(base / "profiles" / "example_profile.json")
    )
)

profile = load_profile(profile_path)

print("Perfil carregado:", profile.name)
print("Destino:", profile.target)
print("Arquivo de perfil:", profile.path)
