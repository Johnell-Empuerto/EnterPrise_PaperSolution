"""
Find the location of the Microsoft Office interop assembly (office.dll)
by querying the Windows Registry for the registered type libraries.
"""
import winreg
import os

def query_registry_value(key, subkey, name=""):
    try:
        with winreg.OpenKey(key, subkey) as k:
            value, _ = winreg.QueryValueEx(k, name)
            return value
    except FileNotFoundError:
        return None
    except Exception as e:
        return f"Error: {e}"

def find_excel_typelib():
    """Find the Excel type library path."""
    # Excel TypeLib GUID
    guid = "{00020813-0000-0000-C000-000000000046}"
    
    # Check various versions
    for version in ["1.9", "1.8", "1.7"]:
        for arch in ["win64", "win32"]:
            subkey = f"TypeLib\\{guid}\\{version}\\0\\{arch}"
            path = query_registry_value(winreg.HKEY_CLASSES_ROOT, subkey)
            if path and os.path.exists(path):
                return path, version, arch
    return None, None, None

def find_office_typelib():
    """Find the Office type library path (office.dll)."""
    # Office TypeLib GUID
    guid = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
    
    # Check various versions
    for version in ["2.8", "2.7", "2.6", "2.5"]:
        for arch in ["win64", "win32"]:
            subkey = f"TypeLib\\{guid}\\{version}\\0\\{arch}"
            path = query_registry_value(winreg.HKEY_CLASSES_ROOT, subkey)
            if path and os.path.exists(path):
                return path, version, arch
    return None, None, None

def find_in_gac(assembly_name):
    """Search for an assembly in the GAC."""
    gac_paths = [
        r"C:\Windows\assembly\GAC_MSIL",
        r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL",
    ]
    for gac_root in gac_paths:
        if os.path.exists(gac_root):
            for root, dirs, files in os.walk(gac_root):
                for f in files:
                    if f.lower() == assembly_name.lower():
                        return os.path.join(root, f)
    return None

print("=" * 60)
print("Office Interop Assembly Diagnostics")
print("=" * 60)

# Check Excel TypeLib
print("\n1. Excel Type Library:")
excel_path, excel_ver, excel_arch = find_excel_typelib()
if excel_path:
    print(f"   Found: {excel_path}")
    print(f"   Version: {excel_ver}, Arch: {excel_arch}")
else:
    print("   NOT FOUND in registry")

# Check Office TypeLib
print("\n2. Office Type Library (office.dll):")
office_path, office_ver, office_arch = find_office_typelib()
if office_path:
    print(f"   Found: {office_path}")
    print(f"   Version: {office_ver}, Arch: {office_arch}")
else:
    print("   NOT FOUND in registry")

# Check Excel installation path
print("\n3. Excel Installation:")
for key_path in [
    r"SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot",
    r"SOFTWARE\Microsoft\Office\15.0\Excel\InstallRoot",
]:
    path = query_registry_value(winreg.HKEY_LOCAL_MACHINE, key_path, "Path")
    if path:
        print(f"   Office path: {path}")
        excel_exe = os.path.join(path, "EXCEL.EXE")
        if os.path.exists(excel_exe):
            print(f"   EXCEL.EXE exists: Yes")
        else:
            print(f"   EXCEL.EXE exists: No")
        break

# Check GAC for office.dll
print("\n4. GAC Search for office.dll:")
gac_path = find_in_gac("office.dll")
if gac_path:
    print(f"   Found: {gac_path}")
else:
    print("   Not found in GAC (this is expected for .NET Core SDK projects)")

# Check the project's output directory
print("\n5. NuGet Package Files:")
nuget_cache = os.path.expanduser("~/.nuget/packages/microsoft.office.interop.excel")
if os.path.exists(nuget_cache):
    versions = os.listdir(nuget_cache)
    print(f"   Cached versions: {versions}")
    for ver in versions:
        lib_dir = os.path.join(nuget_cache, ver, "lib")
        if os.path.exists(lib_dir):
            frameworks = os.listdir(lib_dir)
            print(f"   {ver}: frameworks = {frameworks}")
            for fw in frameworks:
                fw_dir = os.path.join(lib_dir, fw)
                if os.path.isdir(fw_dir):
                    files = os.listdir(fw_dir)
                    for f in files:
                        print(f"      {fw}/{f}")

# Summary
print("\n" + "=" * 60)
print("DIAGNOSTICS SUMMARY")
print("=" * 60)

if not excel_path:
    print("❌ Excel type library NOT registered - Excel may not be properly installed")
elif not office_path:
    print("⚠️  Excel type library found but Office type library NOT registered")
    print("   This is the likely cause of the missing office.dll error.")
    print("   Solution: Register or install Office PIAs, or use a direct COM reference.")
else:
    print("✅ Both Excel and Office type libraries are registered")
