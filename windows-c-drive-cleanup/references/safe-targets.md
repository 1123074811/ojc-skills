# 安全清理目标清单

## 白名单（可建议清理，用户确认后执行）

### 临时目录

| 路径 | 风险 | 备注 |
|------|------|------|
| `%LOCALAPPDATA%\Temp` | 低 | 占用文件跳过 |
| `C:\Windows\Temp` | 低 | 权限不足跳过 |
| `C:\Temp` / `C:\tmp` | 低 | 自定义临时目录 |
| `%APPDATA%\Temp` | 低 | 若存在 |

### 包管理 / 工具缓存（可重建）

| 路径 | 风险 | 备注 |
|------|------|------|
| `%USERPROFILE%\.cache` | 低 | 通用缓存 |
| `%USERPROFILE%\.npm` | 低 | 旧 npm 缓存 |
| `%LOCALAPPDATA%\npm-cache` | 低 | npm cache |
| `%LOCALAPPDATA%\pip\Cache` | 低 | pip |
| `%LOCALAPPDATA%\pnpm-store` 或 pnpm cache | 中 | 确认未当主 store 误删项目 |
| `%APPDATA%\Code\CachedExtensionVSIXs` | 低 | VS Code 扩展 vsix 缓存 |
| `%APPDATA%\Code\Cache` / `CachedData` / `logs` | 低 | 编辑器缓存日志 |
| `%LOCALAPPDATA%\Microsoft\Windows\INetCache` | 低 | IE/旧网络缓存 |
| `%USERPROFILE%\.gradle\caches` | 中 | 仅清 caches 子目录更稳；全删 `.gradle` 影响 daemon 配置 |

### 应用更新残留

| 模式 | 风险 | 备注 |
|------|------|------|
| `%LOCALAPPDATA%\*-updater` | 低 | installer / pending exe |
| `%LOCALAPPDATA%\*\Updates\*.exe` | 低 | 如 PowerToys Updates |
| `%LOCALAPPDATA%\Temp\*Setup*.exe` | 低 | 临时安装包 |

### 下载目录（需确认）

- 仅建议删除明确安装包：`*Setup*.exe`、`*Installer*.exe`、网盘/IDE 安装包
- 不删：文档、图片、项目压缩包（除非用户点名）

### 其他

| 操作 | 风险 | 备注 |
|------|------|------|
| 清空回收站 | 低 | `Clear-RecycleBin -Force` |
| `C:\Windows\SoftwareDistribution\Download` | 低-中 | Windows Update 下载缓存 |
| Claude 桌面 `vm_bundles` | 中 | 仅用户确认不用桌面 VM/Cowork |

## 灰名单（先说明再清）

| 目标 | 影响 |
|------|------|
| Docker unused data | 删未用镜像/容器/build cache |
| `.m2\repository` 部分 | Maven 依赖重下 |
| Android SDK 旧 platform/system-images | 模拟器/编译缺组件 |
| `C:\$WinREAgent` | 恢复环境代理残留，确认非升级中 |
| `ProgramData\Package Cache` | 安装修复可能需要 |

## 黑名单（默认不动）

- `C:\Windows\System32`、`WinSxS` 手动删文件
- `C:\Windows\Installer` 手动删（会破坏卸载）
- 用户 `Documents` / `Desktop` / 未确认 Downloads
- SSH 密钥、浏览器 Profile 全删
- 活动数据库数据目录、虚拟机正在用的 VHDX（先退出程序）
- Claude Code CLI 安装目录与 `%USERPROFILE%\.claude` 配置（除非用户明确重装 CLI）

## 推荐 PowerShell 片段

### 测目录大小

```powershell
function Get-DirSizeGB([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object Length -Sum).Sum
  if ($null -eq $sum) { $sum = 0 }
  [math]::Round($sum / 1GB, 2)
}
```

### 安全清空目录内容（保留目录）

```powershell
function Clear-DirContents([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return }
  Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
  }
}
```

### 大文件 Top N

```powershell
Get-ChildItem -LiteralPath $roots -Force -Recurse -File -ErrorAction SilentlyContinue |
  Where-Object { $_.Length -ge 100MB } |
  Sort-Object Length -Descending |
  Select-Object -First 80 @{N='GB';E={[math]::Round($_.Length/1GB,2)}}, FullName
```

### 注意：foreach 管道

错误：

```powershell
foreach ($p in $paths) { ... } | Sort-Object SizeGB
```

正确：

```powershell
$results = foreach ($p in $paths) { ... }
$results | Sort-Object SizeGB -Descending
```
