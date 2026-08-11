#!/usr/bin/env bash
# Fase 1 — Diagnóstico NvidiaSpark / DGX Spark para VM Win11 ARM + Inventor
# Ejecutar DENTRO del Terminal (o Cursor) de NvidiaSpark, no en la PC Windows local.
set -u

TS="$(date -Iseconds 2>/dev/null || date)"
OUT_DIR="$(cd "$(dirname "$0")" && pwd)/evidencias"
mkdir -p "$OUT_DIR"
OUT="$OUT_DIR/fase1_diagnostico.txt"

exec > >(tee "$OUT") 2>&1

echo "============================================================"
echo "FASE 1 DIAGNOSTICO — NvidiaSpark Win11/Inventor"
echo "Timestamp: $TS"
echo "Host: $(hostname 2>/dev/null || echo unknown)"
echo "Usuario: $(whoami 2>/dev/null || echo unknown)"
echo "PWD: $(pwd)"
echo "============================================================"
echo

section() { echo; echo "----- $1 -----"; }

section "1) Arquitectura y kernel"
uname -a || true
echo "uname -m: $(uname -m)"
lscpu 2>/dev/null | sed -n '1,40p' || true

section "2) Memoria y disco"
free -h || true
df -hT || true
echo
echo "Espacio libre en \$HOME:"
df -h "$HOME" || true

section "3) Virtualización KVM"
ls -l /dev/kvm 2>&1 || true
if [[ -e /dev/kvm ]]; then
  echo "KVM_DEVICE=OK"
else
  echo "KVM_DEVICE=MISSING"
fi
groups || true
id || true
# Lectura no destructiva de flags CPU
if command -v egrep >/dev/null 2>&1; then
  egrep -c '(vmx|svm)' /proc/cpuinfo 2>/dev/null || true
fi
# En ARM lo relevante es kvm_arm / /dev/kvm
lsmod 2>/dev/null | egrep -i 'kvm|vhost' || true
cat /proc/cpuinfo 2>/dev/null | egrep -i 'Features|CPU implementer|CPU architecture|CPU part' | sort -u | head -n 40 || true

section "4) Privilegio sudo"
if sudo -n true 2>/dev/null; then
  echo "SUDO_NOPASSWD=OK"
elif sudo -v 2>/dev/null; then
  echo "SUDO_INTERACTIVE=OK"
else
  echo "SUDO=NO_O_FALLA"
fi

section "5) Docker / Podman"
command -v docker && docker --version || echo "docker=NO"
command -v podman && podman --version || echo "podman=NO"
docker info 2>/dev/null | sed -n '1,35p' || true
docker compose version 2>/dev/null || docker-compose version 2>/dev/null || echo "compose=NO"

section "6) QEMU / libvirt"
command -v qemu-system-aarch64 && qemu-system-aarch64 --version | head -n 2 || echo "qemu-system-aarch64=NO"
command -v virsh && virsh version 2>/dev/null | head -n 10 || echo "virsh=NO"

section "7) GPU / NVIDIA (contexto host; no implica passthrough a VM)"
command -v nvidia-smi && nvidia-smi || echo "nvidia-smi=NO"
command -v nvidia-container-cli && nvidia-container-cli info 2>/dev/null | head -n 20 || true

section "8) Red básica"
ip -br a 2>/dev/null || ifconfig 2>/dev/null | head -n 40 || true
echo "Puertos escuchando (sample):"
ss -lntup 2>/dev/null | head -n 40 || netstat -lntup 2>/dev/null | head -n 40 || true

section "9) Estimación de capacidad"
AVAIL_GB=$(df -P "$HOME" 2>/dev/null | awk 'NR==2{printf "%d", $4/1024/1024}')
MEM_GB=$(free -g 2>/dev/null | awk '/^Mem:/{print $2}')
echo "HOME_FREE_GB_APPROX=${AVAIL_GB:-unknown}"
echo "MEM_TOTAL_GB_APPROX=${MEM_GB:-unknown}"

GO=1
REASON=()
ARCH=$(uname -m)
if [[ "$ARCH" != "aarch64" && "$ARCH" != "arm64" ]]; then
  # Aun así se documenta; el plan oficial asume Spark ARM
  REASON+=("Arquitectura inesperada: $ARCH (el plan PoC asume aarch64)")
fi
if [[ ! -e /dev/kvm ]]; then
  GO=0
  REASON+=("Falta /dev/kvm — sin KVM no hay ruta A/B viable")
fi
if ! command -v docker >/dev/null 2>&1 && ! command -v podman >/dev/null 2>&1; then
  # No es no-go absoluto si hay sudo para instalar
  REASON+=("Docker/Podman no instalado aún (se intentará instalar en Fase 2 si hay sudo)")
fi
if [[ -n "${AVAIL_GB:-}" && "$AVAIL_GB" -lt 100 ]]; then
  GO=0
  REASON+=("Espacio libre < 100 GB en HOME ($AVAIL_GB GB)")
fi
if [[ -n "${MEM_GB:-}" && "$MEM_GB" -lt 12 ]]; then
  GO=0
  REASON+=("RAM host < 12 GB reportada ($MEM_GB GB)")
fi

section "10) VEREDICTO GO/NO-GO"
if [[ $GO -eq 1 ]]; then
  echo "VERDICTO=GO — Continuar a Fase 2 (dockur/windows-arm)"
else
  echo "VERDICTO=NO-GO — No continuar instalación hasta resolver bloqueos"
fi
echo "Observaciones:"
if [[ ${#REASON[@]} -eq 0 ]]; then
  echo "  - Sin bloqueos críticos detectados por el script"
else
  for r in "${REASON[@]}"; do
    echo "  - $r"
  done
fi

echo
echo "============================================================"
echo "Salida guardada en: $OUT"
echo "============================================================"
