#!/usr/bin/env bash
set -e

APP_ID="com.microsoft365linux.Unofficial"
REPO_DIR="repo"
BUILD_DIR="build-dir"
BUNDLE_NAME="Microsoft365-for-Linux-Unofficial.flatpak"

# Detecta se o manifesto local usa extensão .yml ou .yaml
if [ -f "${APP_ID}.yml" ]; then
  MANIFEST="${APP_ID}.yml"
elif [ -f "${APP_ID}.yaml" ]; then
  MANIFEST="${APP_ID}.yaml"
else
  echo "Erro: Manifesto ${APP_ID}.yml ou ${APP_ID}.yaml não foi encontrado no diretório."
  exit 1
fi

echo "==> 1. Limpando artefatos anteriores..."
rm -rf "${BUILD_DIR}" "${REPO_DIR}" "${BUNDLE_NAME}"

echo "==> 2. Compilando pacote Flatpak via ${MANIFEST}..."
flatpak-builder \
  --force-clean \
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
