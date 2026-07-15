# 迁移映射与配置

## 标准目录布局（目标盘）

以 `D:` 为例：

```text
D:\
  Docker\                 # Docker Desktop 数据（vhdx）
  Android\                # Android SDK
  DevCache\
    gradle\
    m2\
    jdks\
    miniconda3\
    npm-global\
    npm-cache\
    pub\
    pip\
    develop\
    README-migration.txt
  Temp\
    <USERNAME>\           # 用户 TEMP/TMP
  pagefile.sys            # 可选系统分页文件
```

## 迁移算法（必须遵守）

先输出源、目标、大小、影响和回滚建议；仅在用户明确确认该映射后，使用 `-Execute` 执行以下步骤。不带 `-Execute` 时只输出计划。

1. 若源已是 `Junction`/`SymbolicLink` → 跳过  
2. `New-Item` 创建目标父目录  
3. `robocopy <src> <dst> /E /COPY:DAT /R:1 /W:1 /XJ /MT:8 /NFL /NDL /NP /NJH /NJS`  
4. robocopy 退出码：`0–7` 成功；`>=8` 失败（可检查是否已拷大部分）  
5. 将源重命名为 `<src>.pre-migrate-bak`  
6. `cmd /c mklink /J "<src>" "<dst>"`  
7. 确认 `Get-Item <src>` 的 `LinkType -eq 'Junction'`  
8. 删除 `.pre-migrate-bak`  
9. 写环境变量 / npm config  
10. 验收

### 重命名被锁时

- 不自动删除或清空源目录，不按目录大小推断复制完整。  
- 保留「源 + 目标双份」，报告占用进程线索，让用户关闭相关软件后重新审计和确认。  
- 不要对未知进程滥用 `Stop-Process`。

## 映射表

| Label | Source | Dest | User env / config |
|-------|--------|------|-------------------|
| docker | `%LOCALAPPDATA%\Docker` | `D:\Docker` | Docker Desktop 设置；迁移前退出 Docker |
| android | `%LOCALAPPDATA%\Android` | `D:\Android` | `ANDROID_HOME=D:\Android`，`ANDROID_SDK_ROOT=D:\Android`；PATH 可继续用原 `...\Android\Sdk\...`（经 junction） |
| gradle | `%USERPROFILE%\.gradle` | `D:\DevCache\gradle` | `GRADLE_USER_HOME=D:\DevCache\gradle` |
| m2 | `%USERPROFILE%\.m2` | `D:\DevCache\m2` | 可选 settings.xml；一般靠 junction |
| jdks | `%USERPROFILE%\.jdks` | `D:\DevCache\jdks` | JetBrains 识别原路径即可 |
| miniconda3 | `%USERPROFILE%\miniconda3` | `D:\DevCache\miniconda3` | PATH 保持原路径；conda init 通常仍有效 |
| npm-global | `%APPDATA%\npm` | `D:\DevCache\npm-global` | `npm config set prefix` 仍指向原 AppData\npm；用户 PATH 含该路径 |
| npm-cache | `%LOCALAPPDATA%\npm-cache` | `D:\DevCache\npm-cache` | `npm_config_cache` + `npm config set cache` |
| pub | `%LOCALAPPDATA%\Pub` | `D:\DevCache\pub` | `PUB_CACHE=D:\DevCache\pub\Cache`（若结构是 Pub\Cache） |
| pip | `%LOCALAPPDATA%\pip` | `D:\DevCache\pip` | `PIP_CACHE_DIR=D:\DevCache\pip\Cache` |
| develop | `%USERPROFILE%\develop` | `D:\DevCache\develop` | Flutter 等 PATH 保持 `C:\Users\...\develop\flutter\bin` |
| temp | 用户 TEMP | `D:\Temp\<user>` | `TEMP`/`TMP` 用户级 |

## 环境变量写入模板

```powershell
$drive = 'D:'
$user = $env:USERNAME
$dev = "$drive\DevCache"
$temp = "$drive\Temp\$user"

[Environment]::SetEnvironmentVariable('GRADLE_USER_HOME', "$dev\gradle", 'User')
[Environment]::SetEnvironmentVariable('PUB_CACHE', "$dev\pub\Cache", 'User')
[Environment]::SetEnvironmentVariable('npm_config_cache', "$dev\npm-cache", 'User')
[Environment]::SetEnvironmentVariable('PIP_CACHE_DIR', "$dev\pip\Cache", 'User')
[Environment]::SetEnvironmentVariable('TEMP', $temp, 'User')
[Environment]::SetEnvironmentVariable('TMP', $temp, 'User')
[Environment]::SetEnvironmentVariable('ANDROID_HOME', "$drive\Android", 'User')
[Environment]::SetEnvironmentVariable('ANDROID_SDK_ROOT', "$drive\Android", 'User')

New-Item -ItemType Directory -Force -Path $temp, "$dev\npm-cache", "$dev\pip\Cache" | Out-Null

npm config set cache "$dev\npm-cache" --location=user
npm config set prefix "$env:APPDATA\npm" --location=user
```

## PATH 注意

- **优先靠 junction 保原路径**，少改 PATH，兼容性最好。  
- 若 PATH 已写死 `D:\...` 与 junction 并存，通常无妨。  
- 迁移后检查：

```powershell
[Environment]::GetEnvironmentVariable('Path','User') -split ';' |
  Where-Object { $_ -match 'npm|Android|flutter|miniconda|Python|gradle' }
```

## 回滚

1. 删除联接点（只删联接，别删目标盘数据）：`cmd /c rmdir "C:\path\to\junction"`  
2. 若仍有 `.pre-migrate-bak`：将其改回原名  
3. 若 bak 已删：把 `D:\...` 目标再 robocopy 回 C 原路径  

## 空间计算

- 迁移瞬时占用：目标盘需 ≥ 源大小（复制阶段 C 与 D 都占一份）  
- 联接成功并删 bak 后：C 释放约等于源大小；D 增加约等于源大小  
