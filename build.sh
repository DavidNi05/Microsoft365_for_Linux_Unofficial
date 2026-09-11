#!/usr/bin/env bash
set -e

APP_ID="io.github.DavidNi05.Microsoft365_for_Linux_Unofficial"
REPO_DIR="repo"
BUILD_DIR="build-dir"
BUNDLE_NAME="Microsoft365-for-Linux-Unofficial.flatpak"
MANIFEST="${APP_ID}.json"

if [ "$1" = "--uninstall" ] || [ "$1" = "uninstall" ]; then
    echo "==> 1. Executando limpeza de dados residuais no host..."
    flatpak run "${APP_ID}" --clean-uninstall 2>/dev/null || true
    
    echo "==> 2. Desinstalando o Flatpak e apagando dados do sandbox (~/.var/app)..."
    flatpak uninstall --delete-data -y "${APP_ID}" 2>/dev/null || true
    
    echo "✓ Aplicativo desinstalado e todos os dados foram apagados com sucesso!"
    exit 0
fi

if [ ! -f "${MANIFEST}" ]; then
    echo "Erro: Manifesto ${MANIFEST} não encontrado."
    exit 1
fi

echo "==> 1. Limpando artefatos anteriores..."
rm -rf "${BUILD_DIR}" "${REPO_DIR}" "${BUNDLE_NAME}"

echo "==> 2. Compilando pacote Flatpak via ${MANIFEST} (GNOME Platform 50)..."
flatpak-builder \
    --force-clean \
    --user \
    --install-deps-from=flathub \
    --repo="${REPO_DIR}" \
    "${BUILD_DIR}" \
    "${MANIFEST}"

echo "==> 3. Gerando o bundle autônomo (.flatpak)..."
flatpak build-bundle \
    "${REPO_DIR}" \
    "${BUNDLE_NAME}" \
    "${APP_ID}"

echo "✓ Processo finalizado com sucesso!"
echo "✓ Pacote exportado: ${BUNDLE_NAME}"
