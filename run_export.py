import os
import subprocess
from pathlib import Path


def export_plant_json(
    project_xml: str | os.PathLike,
    out_json: str | os.PathLike,
    plugin_dll: str | os.PathLike,
    *,
    acad_exe: str | os.PathLike = r"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe",
    workdir: str | os.PathLike = r"C:\PlantAutomationRun",
) -> int:
    """Run AutoCAD Plant 3D with the EXPORTPLANTJSON command and return exit code."""

    acad = Path(acad_exe)
    project_xml = Path(project_xml)
    out_json = Path(out_json)
    plugin_dll = Path(plugin_dll)
    workdir = Path(workdir)

    if not acad.exists():
        raise FileNotFoundError(str(acad))
    if not project_xml.exists():
        raise FileNotFoundError(str(project_xml))
    if not plugin_dll.exists():
        raise FileNotFoundError(str(plugin_dll))

    workdir.mkdir(parents=True, exist_ok=True)
    out_json.parent.mkdir(parents=True, exist_ok=True)

    scr_path = workdir / "run.scr"

    # AutoCAD script: NETLOAD plugin, run command, exit.
    scr_text = "\n".join(
        [
            "_.NETLOAD",
            f"\"{plugin_dll}\"",
            "_.EXPORTPLANTJSON",
            "_.QUIT",
            "Y",
        ]
    ) + "\n"
    scr_path.write_text(scr_text, encoding="utf-8")

    env = os.environ.copy()
    env["PLANT_PROJECT_XML"] = str(project_xml)
    env["PLANT_JSON_OUT"] = str(out_json)

    cmd = [
        str(acad),
        "/product",
        "PLNT3D",
        "/b",
        str(scr_path),
    ]

    completed = subprocess.run(cmd, env=env, cwd=str(workdir))
    return completed.returncode


def default_plugin_dll() -> Path:
    repo_root = Path(__file__).resolve().parent
    return (
        repo_root
        / "addin"
        / "bin"
        / "Release"
        / "net8.0-windows"
        / "PlantJsonExporter.dll"
    )


def main() -> int:
    project_xml = os.environ.get(
        "PLANT_PROJECT_XML",
        r"C:\Users\aleja\Downloads\NIRAS P3D Demo Project\P3D-FBUK Example Project\Project.xml",
    )
    out_json = os.environ.get("PLANT_JSON_OUT", r"C:\PlantAutomationRun\plant_export.json")
    plugin_dll = os.environ.get("PLANT_PLUGIN_DLL", str(default_plugin_dll()))
    if Path(out_json).is_dir():
        out_json = str(Path(out_json) / "plant_export.json")
    return export_plant_json(project_xml, out_json, plugin_dll)


if __name__ == "__main__":
    raise SystemExit(main())
