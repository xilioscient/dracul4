Param(
    [string]$Configuration = "Release",
    [string]$Generator = "Visual Studio 17 2022",
    [string]$Arch = "x64"
)

if (-not $env:VCPKG_ROOT) {
    Write-Host "Set VCPKG_ROOT to your vcpkg path (e.g., C:\\src\\vcpkg)" -ForegroundColor Yellow
}

cmake -S .. -B ../build -G "$Generator" -A $Arch `
    -DCMAKE_TOOLCHAIN_FILE=$env:VCPKG_ROOT\scripts\buildsystems\vcpkg.cmake

cmake --build ../build --config $Configuration


