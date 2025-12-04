# PCInfoScanner 🔍💻
Una herramienta de análisis de hardware para Windows 10 y Windows 11.  
Detecta **CPU, GPU, RAM, discos, motherboard, ventiladores** y genera un informe completo en tu Escritorio.  
100% compatible con Windows 11.

---

## ✨ Características
- 🔥 Detección completa mediante **PowerShell + CIM** (funciona en Win10/Win11)
- 🧠 Obtiene:
  - CPU (modelo, fabricante, núcleos, hilos)
  - GPU (todas: NVIDIA / AMD / Intel / integradas y dedicadas)
  - RAM total
  - SSD / HDD (modelo, tipo, capacidad)
  - Motherboard (modelo + fabricante)
  - Ventiladores detectados por el sistema
- 📝 Genera un **informe detallado en .txt** en el Escritorio
- 🔗 Incluye **URL de búsqueda** para cada componente
- 🛠️ Funciona como script o compilado a `.exe` con PyInstaller

---
Valualo con una estrellita ⭐😉
---

## 📦 Instalación
Cloná el repositorio:

```bash
git clone https://github.com/1vcbGH/PCInfoScanner
cd PCInfoScanner
