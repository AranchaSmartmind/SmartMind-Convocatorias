
set -e

PROJECT_NAME="smartmind"
COMPOSE_FILE="docker-compose.yml"


RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

print_success() {
    echo -e "${GREEN}✓ $1${NC}"
}

print_error() {
    echo -e "${RED}✗ $1${NC}"
}

print_info() {
    echo -e "${YELLOW}ℹ $1${NC}"
}

check_docker() {
    if ! command -v docker &> /dev/null; then
        print_error "Docker no está instalado. Por favor, instala Docker primero."
        exit 1
    fi
    
    if ! command -v docker-compose &> /dev/null; then
        print_error "Docker Compose no está instalado. Por favor, instala Docker Compose primero."
        exit 1
    fi
    
    print_success "Docker y Docker Compose están instalados"
}

build() {
    print_info "Construyendo imagen Docker..."
    docker-compose build --no-cache
    print_success "Imagen construida exitosamente"
}

start() {
    print_info "Iniciando SmartMind..."
    docker-compose up -d
    print_success "SmartMind iniciado en http://localhost:8501"
}

stop() {
    print_info "Deteniendo SmartMind..."
    docker-compose down
    print_success "SmartMind detenido"
}

restart() {
    stop
    start
}

logs() {
    print_info "Mostrando logs (Ctrl+C para salir)..."
    docker-compose logs -f
}

clean() {
    print_info "Limpiando contenedores, imágenes y volúmenes..."
    docker-compose down -v
    docker rmi ${PROJECT_NAME}:latest 2>/dev/null || true
    print_success "Limpieza completada"
}

status() {
    print_info "Estado de los contenedores:"
    docker-compose ps
}

show_help() {
    echo "Uso: ./start.sh [comando]"
    echo ""
    echo "Comandos disponibles:"
    echo "  build     - Construir la imagen Docker"
    echo "  start     - Iniciar la aplicación"
    echo "  stop      - Detener la aplicación"
    echo "  restart   - Reiniciar la aplicación"
    echo "  logs      - Ver logs en tiempo real"
    echo "  status    - Ver estado de contenedores"
    echo "  clean     - Limpiar todo (contenedores, imágenes, volúmenes)"
    echo "  help      - Mostrar esta ayuda"
    echo ""
    echo "Si no se especifica comando, se ejecuta 'start'"
}

check_docker

case "${1:-start}" in
    build)
        build
        ;;
    start)
        start
        ;;
    stop)
        stop
        ;;
    restart)
        restart
        ;;
    logs)
        logs
        ;;
    status)
        status
        ;;
    clean)
        clean
        ;;
    help|--help|-h)
        show_help
        ;;
    *)
        print_error "Comando desconocido: $1"
        show_help
        exit 1
        ;;
esac