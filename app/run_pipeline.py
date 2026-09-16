"""Orquesta la pipeline completa Bronze -> Silver -> Gold."""
from __future__ import annotations

import build_bronze_manifest
import build_gold
import extract_plan_mejoramiento


def main() -> None:
    print("== 1/3 Bronze -> Silver: catálogo completo ==")
    build_bronze_manifest.main()
    print()
    print("== 2/3 Bronze -> Silver: plan de mejoramiento ==")
    extract_plan_mejoramiento.main()
    print()
    print("== 3/3 Silver -> Gold ==")
    build_gold.main()


if __name__ == "__main__":
    main()
