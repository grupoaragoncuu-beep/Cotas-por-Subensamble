# Ruta A — Windows 11 ARM con dockur/windows-arm en NvidiaSpark
# Ejecutar SOLO después de Fase 1 = GO y SOLO dentro del host Spark (Linux aarch64).
#
# Uso:
#   cd Planos/NvidiaSpark_Win11_Inventor
#   bash 02_fase2_levantar_win11_arm.sh
#
# Acceso UI típico: http://localhost:8006  (noVNC)
# Si usas NVIDIA Sync desde tu PC: crea túnel SSH -L 8006:localhost:8006 si hace falta.

set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
EV="$ROOT/evidencias"
DATA="$ROOT/vm_data"
COMPOSE="$ROOT/docker-compose.yml"
mkdir -p "$EV" "$DATA"

TS="$(date -Iseconds 2>/dev/null || date)"
LOG="$EV/fase2_levantamiento_$(date +%Y%m%d_%H%M%S).txt"

exec > >(tee -a "$LOG") 2>&1

echo "============================================================"
echo "FASE 2 — Levantamiento Win11 ARM (dockur/windows-arm)"
echo "Timestamp: $TS"
echo "============================================================"

echo "[1/6] Validaciones previas"
uname -m
test -e /dev/kvm || { echo "ERROR: /dev/kvm no existe. Abortando."; exit 1; }
command -v docker >/dev/null || { echo "ERROR: docker no está instalado."; exit 1; }

# Verificar pertenencia a grupo docker si aplica
if ! docker info >/dev/null 2>&1; then
  echo "ERROR: 'docker info' falló. ¿Servicio docker activo? ¿Usuario en grupo docker?"
  exit 1
fi

echo "[2/6] Escribiendo docker-compose.yml"
cat > "$COMPOSE" <<'YAML'
# Generado por piloto ordenado NvidiaSpark Win11+Inventor
# Método: dockur/windows-arm (Windows 11 ARM64 vía KVM en contenedor)
services:
  windows:
    image: ghcr.io/dockur/windows-arm:latest
    container_name: nvidiaSpark_win11_arm_inventor
    environment:
      VERSION: "11"
      RAM_SIZE: "16G"
      CPU_CORES: "6"
      DISK_SIZE: "120G"
      LANGUAGE: "Spanish"
      USERNAME: "arga"
      PASSWORD: "ArgaSparkPiloto!"
    devices:
      - /dev/kvm
      - /dev/net/tun
    cap_add:
      - NET_ADMIN
    ports:
      - "8006:8006"   # noVNC web
      - "3389:3389"   # RDP
    volumes:
      - ./vm_data:/storage
    restart: "no"
    stop_grace_period: 2m
YAML

echo "[3/6] Pull de imagen (puede tardar)"
docker pull ghcr.io/dockur/windows-arm:latest

echo "[4/6] Arranque compose"
if docker compose version >/dev/null 2>&1; then
  docker compose -f "$COMPOSE" up -d
elif command -v docker-compose >/dev/null 2>&1; then
  docker-compose -f "$COMPOSE" up -d
else
  echo "ERROR: no hay docker compose"
  exit 1
fi

echo "[5/6] Estado del contenedor"
docker ps -a --filter name=nvidiaSpark_win11_arm_inventor
docker logs --tail 80 nvidiaSpark_win11_arm_inventor || true

echo "[6/6] Instrucciones de acceso"
cat <<EOF

------------------------------------------------------------
SIGUIENTE PASO MANUAL
1) Abre en el navegador del Spark (o túnel Sync):
   http://localhost:8006
2) Espera a que termine la instalación automática de Windows 11 ARM.
   Puede tomar 20-90 minutos la primera vez.
3) Usuario/contraseña definidos en docker-compose.yml (USERNAME/PASSWORD).
4) Cuando el escritorio esté listo, anota evidencia en BITACORA.md
   y continúa con 03_fase3_checklist_inventor.md
------------------------------------------------------------

ADVERTENCIA DE EVIDENCIA / DESLINDE:
- No se configura GPU passthrough (no disponible/soportado en Spark).
- Inventor correrá (si instala) bajo emulación x64 en Win11 ARM + gráficos software.
- Este paso CUMPLE el método ordenado (VM Windows en NvidiaSpark).

Log: $LOG
EOF
