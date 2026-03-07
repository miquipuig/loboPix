# lobopix

Extensio de Chrome (Manifest V3) centrada exclusivament en galeria d'imatges.

## Funcionalitats actuals

- `Galeria Web`: carrega imatges detectades a la pestanya activa.
- `Galeria Manual`: permet pujar i eliminar imatges pròpies (persistides a `chrome.storage.local`).
- `Usuari`: secció reservada (placeholder).
- `About`: informació bàsica de l'app i versió.

## Build

```bash
npm run build
```

Sortida a `dist/`.

## Carrega en Chrome

1. Obre `chrome://extensions`.
2. Activa `Mode desenvolupador`.
3. Prem `Carregar descomprimida`.
4. Selecciona la carpeta `dist/`.
