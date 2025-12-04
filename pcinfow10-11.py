import subprocess
import datetime
import os
import json
import urllib.parse

def run_powershell(ps_command):
    try:
        completed = subprocess.run(
            ["powershell", "-NoLogo", "-NoProfile", "-Command", ps_command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore"
        )
        return completed.stdout.strip()
    except Exception as e:
        return f"ERROR: {e}"

def make_search_url(name):
    if not name or name.lower() == "unknown":
        return ""
    return f"https://www.google.com/search?q={urllib.parse.quote_plus(name)}"

# ------------------ DETECTORES ------------------

def get_cpu_info():
    cmd = "Get-CimInstance Win32_Processor | Select-Object Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors | ConvertTo-Json"
    out = run_powershell(cmd)
    try:
        data = json.loads(out)
        if isinstance(data, dict):
            data = [data]
        cpus = []
        for cpu in data:
            cpus.append({
                "name": cpu.get("Name", "Unknown"),
                "manufacturer": cpu.get("Manufacturer", "Unknown"),
                "cores": cpu.get("NumberOfCores", "Unknown"),
                "threads": cpu.get("NumberOfLogicalProcessors", "Unknown"),
                "url": make_search_url(cpu.get("Name", ""))
            })
        return cpus
    except:
        return []

def get_gpu_info():
    cmd = "Get-CimInstance Win32_VideoController | Select-Object Name,AdapterCompatibility,DriverVersion | ConvertTo-Json"
    out = run_powershell(cmd)
    try:
        data = json.loads(out)
        if isinstance(data, dict):
            data = [data]
        gpus = []
        for gpu in data:
            gpus.append({
                "name": gpu.get("Name", "Unknown"),
                "vendor": gpu.get("AdapterCompatibility", "Unknown"),
                "driver": gpu.get("DriverVersion", "Unknown"),
                "url": make_search_url(gpu.get("Name", ""))
            })
        return gpus
    except:
        return []

def get_ram_info():
    cmd = "(Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory"
    out = run_powershell(cmd)
    try:
        ram_gb = round(int(out) / (1024**3))
        return ram_gb
    except:
        return 0

def get_disks():
    cmd = "Get-CimInstance Win32_DiskDrive | Select-Object Model,MediaType,Size | ConvertTo-Json"
    out = run_powershell(cmd)
    try:
        data = json.loads(out)
        if isinstance(data, dict):
            data = [data]
        disks = []
        for d in data:
            size_gb = 0
            try:
                size_gb = round(int(d.get("Size", 0)) / (1024**3))
            except:
                pass

            disks.append({
                "model": d.get("Model", "Unknown"),
                "type": d.get("MediaType", "Unknown"),
                "size_gb": size_gb,
                "url": make_search_url(d.get("Model", ""))
            })
        return disks
    except:
        return []

def get_motherboard():
    cmd = "Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product | ConvertTo-Json"
    out = run_powershell(cmd)
    try:
        data = json.loads(out)
        mb = f"{data.get('Manufacturer', '')} {data.get('Product', '')}".strip()
        return {
            "name": mb,
            "url": make_search_url(mb)
        }
    except:
        return {"name": "Unknown", "url": ""}

def get_fans():
    cmd = "Get-CimInstance Win32_Fan | Select-Object Name | ConvertTo-Json"
    out = run_powershell(cmd)
    try:
        data = json.loads(out)
        if isinstance(data, dict):
            data = [data]
        fans = [f.get("Name", "Unknown") for f in data]
        return fans
    except:
        return []

# ------------------ INFORME ------------------

def main():
    now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        desktop = os.getcwd()

    file_path = os.path.join(desktop, f"PC_INFO_{now}.txt")

    cpu = get_cpu_info()
    gpu = get_gpu_info()
    ram = get_ram_info()
    disks = get_disks()
    mb = get_motherboard()
    fans = get_fans()

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("INFORME COMPLETO DEL EQUIPO (Compatible con Windows 11)\n")
        f.write("========================================================\n\n")

        f.write("=== CPU ===\n")
        for c in cpu:
            f.write(f"Modelo      : {c['name']}\n")
            f.write(f"Fabricante  : {c['manufacturer']}\n")
            f.write(f"Núcleos     : {c['cores']}\n")
            f.write(f"Hilos       : {c['threads']}\n")
            f.write(f"URL         : {c['url']}\n\n")

        f.write("=== GPU ===\n")
        for g in gpu:
            f.write(f"Modelo      : {g['name']}\n")
            f.write(f"Vendor      : {g['vendor']}\n")
            f.write(f"Driver      : {g['driver']}\n")
            f.write(f"URL         : {g['url']}\n\n")

        f.write(f"=== RAM ===\nTotal detectado: {ram} GB\n\n")

        f.write("=== Discos ===\n")
        for d in disks:
            f.write(f"Modelo : {d['model']}\n")
            f.write(f"Tipo   : {d['type']}\n")
            f.write(f"Tamaño : {d['size_gb']} GB\n")
            f.write(f"URL    : {d['url']}\n\n")

        f.write("=== Motherboard ===\n")
        f.write(f"Modelo : {mb['name']}\n")
        f.write(f"URL    : {mb['url']}\n\n")

        f.write("=== Ventiladores / Fans ===\n")
        if fans:
            for fan in fans:
                f.write(f"- {fan}\n")
        else:
            f.write("No reportados por el sistema.\n")

    print(f"Informe generado en: {file_path}")
    input("Presiona ENTER para salir...")

if __name__ == "__main__":
    main()
