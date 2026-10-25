# PyInstaller spec: one folder with two executables sharing the same libraries
#   ksefiz.exe          console helper (CLI, excel-run for the workbook buttons)
#   KSeF IZ Setup.exe   windowed setup assistant
# Build:  pyinstaller packaging/windows/ksefiz.spec --noconfirm --distpath dist --workpath build
from PyInstaller.utils.hooks import collect_submodules, copy_metadata

here = SPECPATH  # noqa: F821 - provided by PyInstaller
hidden = collect_submodules("keyring.backends") + ["win32ctypes.core"]
datas = copy_metadata("keyring")  # keyring finds its backends through package metadata

common = dict(pathex=[], hiddenimports=hidden, datas=datas, excludes=["pytest", "mypy"])
cli = Analysis([f"{here}/ksefiz_cli.py"], **common)  # noqa: F821
gui = Analysis([f"{here}/ksefiz_setup.py"], **common)  # noqa: F821

cli_exe = EXE(PYZ(cli.pure), cli.scripts, [], exclude_binaries=True,  # noqa: F821
              name="ksefiz", console=True, upx=False)
gui_exe = EXE(PYZ(gui.pure), gui.scripts, [], exclude_binaries=True,  # noqa: F821
              name="KSeF IZ Setup", console=False, upx=False)

COLLECT(cli_exe, cli.binaries, cli.datas,  # noqa: F821
        gui_exe, gui.binaries, gui.datas,
        name="ksefiz", upx=False)
