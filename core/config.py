from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass
class BuildOptions:
    dossier_path: Path
    output_path: Path
    mode: str
    verbose: bool = False

    @property
    def full_cv(self) -> bool:
        return self.mode == "full"

    @property
    def short_cv(self) -> bool:
        return self.mode == "short"
