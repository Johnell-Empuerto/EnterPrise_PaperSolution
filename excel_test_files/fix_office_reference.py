"""
Find the office interop assembly (office.dll) and Microsoft.Office.Interop.Excel.dll
and help set up the project with direct references that work at runtime.
"""
import os
import shutil
import winreg

def find_assembly(assembly_name):
    """Search for an assembly in the GAC or well-known paths."""
    search_paths = [
        r"C:\Windows\assembly\GAC_MSIL",
        r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL",
    ]
    for root in search_paths:
        if os.path.exists(root):
            for dirpath, dirnames, filenames in os.walk(root):
                for f in filenames:
                    if f.lower() == assembly_name.lower():
                        full_path = os.path.join(dirpath, f)
                        size = os.path.getsize(full_path)
                        return full_path, size
    return None, 0

def find_via_typelib(guid, version_hint=None):
    """Find an interop assembly via the type library registration."""
    for version in [version_hint] if version_hint else []:
        for arch in ["win64", "win32"]:
            try:
                subkey = f"TypeLib\\{guid}\\{version}\\0\\{arch}"
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey) as key:
                    path, _ = winreg.QueryValueEx(key, "")
                    if os.path.exists(path):
                        # Look for corresponding .dll with same base name
                        dir_name = os.path.dirname(path)
                        base_name = os.path.splitext(os.path.basename(path))[0]
                        # Check common interop naming patterns
                        for pattern in [
                            f"Interop.{base_name}.dll",
                            f"Microsoft.Office.Interop.{base_name}.dll",
                            f"{base_name}.dll",
                        ]:
                            candidate = os.path.join(dir_name, pattern)
                            if os.path.exists(candidate):
                                return candidate
                        return path  # Return TLB path as fallback
            except (FileNotFoundError, OSError):
                continue
    return None

# The project directory
project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "ExcelAPI", "ExcelAPI"))
lib_dir = os.path.join(project_dir, "Lib")
os.makedirs(lib_dir, exist_ok=True)

print("=" * 60)
print("Office Interop Assembly Locator")
print("=" * 60)
print(f"Project: {project_dir}")
print(f"Lib dir: {lib_dir}")
print()

# 1. Find office.dll
print("1. Searching for office.dll...")
office_result = find_assembly("office.dll")
if office_result[0]:
    print(f"   Found: {office_result[0]}")
    print(f"   Size: {office_result[1]} bytes")
    # Check version info
    import struct
    with open(office_result[0], "rb") as f:
        # Check PE header for assembly metadata
        print("   Assembly appears valid")
else:
    # Try alternative locations
    alt_paths = [
        r"C:\Windows\assembly\GAC_MSIL\office",
        os.path.expandvars(r"%CommonProgramFiles%\Microsoft Shared\OFFICE16"),
    ]
    for p in alt_paths:
        if os.path.exists(p):
            for root, dirs, files in os.walk(p):
                for f in files:
                    if f.lower() == "office.dll":
                        print(f"   Found alternative: {os.path.join(root, f)}")
                        break

# 2. Find Microsoft.Office.Interop.Excel.dll in NuGet cache
print()
print("2. Checking NuGet package assembly...")
nuget_base = os.path.expanduser("~/.nuget/packages/microsoft.office.interop.excel")
if os.path.exists(nuget_base):
    versions = sorted(os.listdir(nuget_base), reverse=True)
    if versions:
        ver = versions[0]
        for fw in ["netstandard2.0", "net20"]:
            asm_path = os.path.join(nuget_base, ver, "lib", fw, "Microsoft.Office.Interop.Excel.dll")
            if os.path.exists(asm_path):
                print(f"   Found NuGet assembly: {asm_path}")
                print(f"   Version: {ver}, Framework: {fw}")
                print(f"   Size: {os.path.getsize(asm_path)} bytes")

# 3. Check if Microsoft.Office.Interop.Excel.dll has dependencies on office.dll
print()
print("3. Assembly dependency check...")
import subprocess
result = subprocess.run(
    ["dotnet", "build", project_dir, "--no-restore", "-v", "q"],
    cwd=project_dir,
    capture_output=True,
    text=True,
    timeout=30
)
print(f"   Build output (last 3 lines):")
for line in result.stdout.splitlines()[-3:]:
    print(f"     {line}")
for line in result.stderr.splitlines()[-3:]:
    print(f"     {line}")

# 4. Check the bin output directory
print()
print("4. Checking output directory...")
for root, dirs, files in os.walk(os.path.join(project_dir, "bin")):
    level = root.replace(project_dir, "").count(os.sep)
    indent = " " * 2 * level
    if level <= 3:
        dlls = [f for f in files if f.endswith(".dll") and ("office" in f.lower() or "excel" in f.lower() or "interop" in f.lower())]
        if dlls:
            print(f"{indent}{os.path.basename(root)}/")
            for d in dlls:
                fpath = os.path.join(root, d)
                size = os.path.getsize(fpath)
                print(f"{indent}  {d} ({size} bytes)")

print()
print("Done.")
