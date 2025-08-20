# exc3l-matrix-editor

Editor da terminale per fogli Excel basato su C++ e libreria xlnt. Permette di visualizzare e modificare una matrice direttamente da un file `.xlsx`.

## Requisiti

- CMake >= 3.21
- Compilatore C++17 (MSVC, clang, MinGW/GCC)
- Libreria `xlnt` (consigliato tramite vcpkg su Windows)

## Struttura del progetto

```
.
├── CMakeLists.txt
├── poseidone/
│   └── zues_x01/
│       ├── CMakeLists.txt
│       └── finaleprova.cpp
└── README.md
```

## Windows (consigliato: vcpkg)

1. Installare [CMake](https://cmake.org/download/) e [Visual Studio Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/), oppure MinGW.
2. Installare vcpkg e la libreria xlnt:
   ```powershell
   git clone https://github.com/microsoft/vcpkg.git
   .\vcpkg\bootstrap-vcpkg.bat
   .\vcpkg\vcpkg install xlnt:x64-windows
   ```
3. Generare e compilare con CMake (MSVC, x64):
   ```powershell
   $env:VCPKG_ROOT = "C:\\path\\to\\vcpkg"
   cmake -S . -B build -G "Visual Studio 17 2022" -A x64 -DCMAKE_TOOLCHAIN_FILE=$env:VCPKG_ROOT\scripts\buildsystems\vcpkg.cmake
   cmake --build build --config Release
   ```

Eseguibile: `build\poseidone\zues_x01\Release\exc3l-matrix-editor.exe`

## Linux

Installare dipendenze e compilare:
```bash
# Esempio con vcpkg
git clone https://github.com/microsoft/vcpkg.git
./vcpkg/bootstrap-vcpkg.sh
./vcpkg/vcpkg install xlnt

export VCPKG_ROOT=$(pwd)/vcpkg
cmake -S . -B build -DCMAKE_TOOLCHAIN_FILE=$VCPKG_ROOT/scripts/buildsystems/vcpkg.cmake
cmake --build build -j
```

In alternativa, installare `xlnt` dal sorgente seguendo le istruzioni del progetto ufficiale e assicurarsi che `find_package(xlnt CONFIG REQUIRED)` trovi il pacchetto.

## Utilizzo

```bash
./build/poseidone/zues_x01/exc3l-matrix-editor
```

Poi inserire il percorso di un file `.xlsx` (ad esempio `poseidone/provao1.xlsx`).

Comandi disponibili in sessione:

```
MOSTRA
AIUTO
SET r c valore
COPIA_DA "file.xlsx" foglio r c righe colonne destR destC
ESPORTA "nuovo_file.xlsx"  (alias: SALVA_COME)
FOGLIO nome|indice
Esci
```

Esempi:

```
SET 3 2 nuovo testo
COPIA_DA "altra_cartella/sorgente.xlsx" "Sheet1" 1 1 10 5 20 3
ESPORTA "output/cartella/risultato.xlsx"
```

## Note di progetto

- Build system unificato con CMake (niente path assoluti).
- `xlnt` viene risolta via `find_package` per essere coerente su Windows e Linux.
- File generati e binari sono esclusi tramite `.gitignore`.


