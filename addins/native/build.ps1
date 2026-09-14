param([string]$OutputDirectory = (Join-Path $PSScriptRoot 'bin'))
$ErrorActionPreference = 'Stop'
foreach ($arch in @('x86_64','i686')) {
    & cargo build --locked --release --manifest-path (Join-Path $PSScriptRoot 'Cargo.toml') --target "$arch-pc-windows-msvc"
    if ($LASTEXITCODE -ne 0) { throw "原生插件 $arch 编译失败。请安装对应 Rust MSVC target 与 Windows SDK。" }
    $destination = Join-Path $OutputDirectory $arch
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "target/$arch-pc-windows-msvc/release/thesiscraft_addin.dll") -Destination (Join-Path $destination 'ThesisCraft.Addin.dll') -Force
}
Write-Output "原生插件已生成：$OutputDirectory"
