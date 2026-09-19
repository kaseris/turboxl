$ErrorActionPreference = 'Stop'
# Pin the tool itself as well as the registry baseline in vcpkg.json.
$revision = (Get-Content "$PSScriptRoot/../../vcpkg.json" | ConvertFrom-Json).'builtin-baseline'
$env:VCPKG_ROOT = "$env:RUNNER_TEMP/turboxl-vcpkg"
if (!(Test-Path "$env:VCPKG_ROOT/.git")) {
    git clone --filter=blob:none --no-checkout https://github.com/microsoft/vcpkg.git $env:VCPKG_ROOT
    if ($LASTEXITCODE) { throw 'vcpkg clone failed' }
}
git -C $env:VCPKG_ROOT checkout --detach $revision
if ($LASTEXITCODE) { throw 'vcpkg checkout failed' }
& "$env:VCPKG_ROOT/bootstrap-vcpkg.bat" -disableMetrics
if ($LASTEXITCODE) { throw 'vcpkg bootstrap failed' }
& "$env:VCPKG_ROOT/vcpkg.exe" install --triplet x64-windows-static-md --clean-after-build
if ($LASTEXITCODE) { throw 'vcpkg install failed' }
$root = $env:VCPKG_ROOT.Replace('\', '/')
"VCPKG_ROOT=$root" >> $env:GITHUB_ENV
$installed = "$env:GITHUB_WORKSPACE/vcpkg_installed".Replace('\', '/')
"VCPKG_INSTALLED_DIR=$installed" >> $env:GITHUB_ENV
"VCPKG_BINARY_SOURCES=$env:VCPKG_BINARY_SOURCES" >> $env:GITHUB_ENV
"CMAKE_TOOLCHAIN_FILE=$root/scripts/buildsystems/vcpkg.cmake" >> $env:GITHUB_ENV
& "$env:VCPKG_ROOT/vcpkg.exe" list
