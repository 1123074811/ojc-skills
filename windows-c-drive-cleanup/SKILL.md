---
name: windows-c-drive-cleanup
description: Windows C盘空间审计（含逐文件夹用途标注）与清理/迁移建议。默认只扫描并向用户报告可清理、可迁移和勿动项；仅在用户明确同意具体清单后，才执行清理或目录联接迁移与环境变量配置。用于 C 盘不足、清理缓存、迁移到D/E盘、pagefile、Docker 占空间、AppData 太大等场景。
---

# Windows C 盘清理与迁移

在 **Windows（优先 PowerShell 7+）** 上完成：

1. **审计**：找出 C 盘大户与可清理项（尽量精确到文件/子目录），并给每个子文件夹标注**用途**（作用）与处置分类
2. **报告建议**：按「容量总览 → 逐文件夹作用表 → 可清理/可迁移/勿动清单」格式输出，附预计释放空间和影响
3. **等待确认**：用户明确同意具体路径、动作和目标盘后，才执行清理或迁移
4. **验收**：核对剩余空间、联接状态、关键工具仍可用

默认目标盘：`D:`（若用户指定其他盘则替换）。  
本机参考配置见 `references/machine-profile-mechrevo.md`。

## 触发与模式

用户可能说：

- 「C 盘快满了 / 帮我看看能删什么」
- 「安全清理一下」
- 「迁移到 D 盘并配好配置」
- 「Claude/Docker/Gradle 占太多」

按意图选模式（可组合）：

| 模式 | 何时用 | 是否先确认 |
|------|--------|------------|
| `audit` | 只分析，不改动 | 否 |
| `safe-clean` | 清临时/缓存/安装包 | **必须确认具体清单后执行** |
| `migrate` | 迁到其他盘 + junction + env | **必须**确认目标盘与列表 |
| `verify` | 检查联接/环境变量/空间 | 否 |

**硬性规则：**

- 默认只执行只读审计并给出报告和建议，然后停止等待用户决定。
- 未获用户对具体路径和动作的明确同意前，不执行任何清理、删除、迁移、联接或环境变量修改。
- 不把「帮我看看」「给建议」「腾空间」视为执行授权；用户必须明确选择报告中的项目。
- 执行脚本时，只有确认后才可传入 `-Execute`；不带该开关必须保持预览模式。
- 不删/不改：`C:\Windows` 系统核心、注册表乱改、正在使用的数据库数据（除非用户明确要求并已停服务）。
- 不把「能腾空间」等同于「该删」。先分类：**可删 / 可迁 / 勿动**。
- Claude **桌面端** `vm_bundles` 与 **Claude Code CLI** 不是同一套数据；删桌面 VM 通常不影响 CLI，但要先说明。

## 快速开始（推荐顺序）

```text
1) audit   → 报告大户、预计空间、风险与建议
2) 等待用户明确选择清理/迁移项目
3) safe-clean / migrate → 仅执行用户确认的项目
4) verify  → 验收
```

脚本（相对本 skill 根目录）：

```powershell
# 审计
pwsh -NoProfile -File "scripts/scan_c_drive.ps1" -UserName $env:USERNAME

# 清理预览（默认不删除）
pwsh -NoProfile -File "scripts/safe_clean.ps1" -UserName $env:USERNAME

# 用户明确同意报告中的具体路径后，只传入获批路径
pwsh -NoProfile -File "scripts/safe_clean.ps1" -UserName $env:USERNAME -Execute `
  -ApprovedPath "$env:LOCALAPPDATA\Temp","$env:LOCALAPPDATA\npm-cache"

# 迁移到 D:\DevCache（示例：gradle）
pwsh -NoProfile -File "scripts/migrate_to_junction.ps1" `
  -Source "C:\Users\$env:USERNAME\.gradle" `
  -Dest "D:\DevCache\gradle" `
  -Label gradle

# 用户明确同意上述源、目标和影响后执行
pwsh -NoProfile -File "scripts/migrate_to_junction.ps1" `
  -Source "C:\Users\$env:USERNAME\.gradle" `
  -Dest "D:\DevCache\gradle" `
  -Label gradle -Execute

# 批量迁移（按 profile）
pwsh -NoProfile -File "scripts/migrate_to_junction.ps1" -UseDefaultDevCache -TargetDrive D
# 用户确认整张批量映射表后才可追加 -Execute

# 验收
pwsh -NoProfile -File "scripts/verify_junctions.ps1" -TargetDrive D -UserName $env:USERNAME
```

若脚本路径含中文/空格，始终用引号；优先 `pwsh`，否则 `powershell`。

---

## 阶段 A：审计（audit）

### A1. 容量总览

```powershell
Get-PSDrive C,D,E -ErrorAction SilentlyContinue |
  Select-Object Name,
    @{N='UsedGB';E={[math]::Round($_.Used/1GB,2)}},
    @{N='FreeGB';E={[math]::Round($_.Free/1GB,2)}},
    @{N='TotalGB';E={[math]::Round(($_.Used+$_.Free)/1GB,2)}}
```

记录：C 剩余 GB、占比；目标盘是否有足够空闲（迁移需要 **源大小 × 1.2** 以上临时余量，因先复制再切联接）。

### A2. C:\ 顶层与用户画像

优先扫描：

- `C:\Users\<user>`（通常最大）
- `C:\Program Files` / `C:\Program Files (x86)`
- `C:\ProgramData`
- `C:\pagefile.sys` / `C:\hiberfil.sys` / `C:\swapfile.sys`
- 非系统自定义目录（如 `C:\mongodb`）

用户目录再下钻：

- `AppData\Local` / `Roaming` / `LocalLow`
- 开发缓存：`.gradle` `.m2` `.jdks` `.cache` `.npm` `miniconda3` `develop`
- 容器：`AppData\Local\Docker`、`*.vhdx`
- 商店/桌面应用：`AppData\Local\Packages\*`

同时检查 `C:\Users` 下的**其他用户配置文件与异常目录**（逐一测大小，往往藏惊喜）：

- 其他用户：`C:\Users\<其他用户名>`（旧账号/中文账号残留）
- 异常目录：`C:\Users\AppData`（不属于任何 profile，说明有程序把 AppData 写错位置）、`CodexSandboxOffline` 等
- 系统默认：`Default` / `Default User` / `Public` / `WsiAccount`（通常 ≈0，保留）

### A3. 大文件

查找 ≥100MB 文件（用户目录 + ProgramData + 自定义目录），按大小排序 Top 50–100。

### A4. 输出报告格式（给用户）

必须按三部分输出，每张表都带**「作用」列**——用户最看重「这个文件夹是干什么的」（实测反馈）。

```markdown
## C 盘审计
- 总容量 / 已用 / 剩余

## C:\ 顶层各文件夹作用
| 路径 | 大小 | 作用 | 处置 |
|------|------|------|------|
| C:\Windows | ..GB | 操作系统本体 | ⚠️ 勿动 |
| C:\pagefile.sys | ..GB | 虚拟内存页面文件 | 🔄 可迁 D |

## 用户目录各子文件夹作用（嵌套下钻，AppData\Local、Roaming 等逐项标注）
| 路径 | 大小 | 作用 | 处置 |
...

## 可清理清单（低风险，预计释放 X GB）｜可迁移清单｜勿动清单
1. 具体路径 + 大小 + 影响说明
```

- 处置分类只用四种：**🗑️ 可清理 / 🔄 可迁移 / ⚠️ 勿动 / 🧩 卸载决策**。
- 文件夹用途按 `references/folder-purposes.md` 知识库标注；查不到的注明「未知，需人工判断」，**不要编造用途**。
- 大小按 A5 的联接排除法测量；junction 项单独标注（其大小不计入 C 盘）。

### A5. 逐文件夹用途标注（folder-purpose inventory）

对 C:\ 顶层、`C:\Users\<user>` 顶层、`AppData\Local`、`AppData\Roaming` 逐项输出「大小 + 作用 + 处置」。

**联接（junction）大小陷阱（实测）**：

- 对**联接路径本身**执行 `Get-ChildItem -Recurse` 会跟随到目标盘，把 D 盘数据误报为 C 占用（例：`AppData\Local\Docker` 是 junction → D:\Docker，直测得到 16.5 GB，实际 C 盘 0）。
- 度量「真实 C 占用（排除联接）」的可靠方法：`robocopy <src> <dummy> /L /E /XJ /BYTES /NFL /NDL /NJH`，从输出 `Bytes : <total>` 取总数（`/L` 只列出不复制；dummy 目录不会被创建）。
- 或 .NET 迭代式测量：`[System.IO.Directory]::EnumerateFileSystemEntries` + `[System.IO.File]::GetAttributes` 过滤 `ReparsePoint`，**根路径本身也要先做 ReparsePoint 检查**，否则直接测 junction 根仍会穿透到目标盘。
- 报告必须列出所有 junction：`Get-Item <path> -Force | Select LinkType,Target`，并注明其数据在目标盘。

脚本 `scripts/scan_c_drive.ps1` 已含「真实 C 占用（robocopy /XJ）」与「其他用户 profile」两节，自动产出上表所需数字。

---

## 阶段 B：安全清理（safe-clean）

### B1. 默认可清（低风险）

| 类别 | 典型路径 | 说明 |
|------|----------|------|
| 用户临时目录 | `%LOCALAPPDATA%\Temp` | 跳过占用中文件 |
| 系统临时 | `C:\Windows\Temp`, `C:\Temp`, `C:\tmp` | 跳过拒绝访问 |
| 应用更新残留 | `%LOCALAPPDATA%\*-updater`, `PowerToys\Updates` | 安装包/ pending |
| 包管理缓存 | `npm-cache`, `pip\Cache`, `.cache`, `.npm` | 可重建 |
| 编辑器缓存 | `%APPDATA%\Code\CachedExtensionVSIXs`, `Cache`, `CachedData`, `logs` | 可重建 |
| 下载安装包 | `Downloads\*.exe` 安装器（确认后） | 勿删用户文档 |
| 回收站 | `Clear-RecycleBin` | 用户同意后 |
| Windows Update 下载缓存 | `C:\Windows\SoftwareDistribution\Download` | 一般安全，少数需服务停用 |

### B2. 有条件可清（先说明影响）

| 目标 | 影响 |
|------|------|
| `.gradle\caches` | 下次构建重下依赖 |
| Docker 未用镜像/容器 | `docker system prune`（先 `docker system df`） |
| Claude 桌面 `vm_bundles` | 桌面端本地 VM/Cowork 需重下；**通常不影响 Claude Code CLI** |
| WinSxS / Installer | 仅用系统清理工具，不手动乱删 |
| 旧 WindowsApps 版本 | 优先「设置 → 应用」卸载 |

### B3. 禁止默认清理

- `C:\Windows\System32`、驱动、激活相关
- 浏览器登录配置、密码库、SSH 私钥
- 正在跑的数据库数据目录
- 用户文档/桌面/未确认的 Downloads 非安装包

### B4. 执行要点

1. 记录清理前 `Get-PSDrive C` 剩余空间  
2. 删除时用 `Remove-Item -Recurse -Force -ErrorAction SilentlyContinue`，忽略占用文件  
3. 清理后再次测剩余空间，用 **盘符 free 差值** 作为「实际释放」  
4. 报告：路径 → 清理前大小 → 状态（cleaned/locked/skip）

先运行不带 `-Execute` 的预览；用户确认报告中的具体清单后，才运行 `scripts/safe_clean.ps1 -Execute -ApprovedPath <获批路径>`。脚本必须跳过未列入 `-ApprovedPath` 的默认目标。

---

## 阶段 C：安全迁移（migrate）

### C1. 原则

- **优先 junction（`mklink /J`）**：软件仍访问原 C 路径，数据在目标盘。  
- **流程必须是**：`robocopy 源→目标` → `重命名源为 *.pre-migrate-bak` → `mklink /J 源 目标` → `删除 bak`。  
- 若复制失败、旧备份存在或源被锁：停止迁移，保留源和已复制目标并报告；不要自动清空源或强杀未知进程。  
- 已是 Junction 的路径：跳过并记入报告。  
- 目标盘空间不足：中止该项。

### C2. 默认可迁移清单（开发机）

| 标签 | 源（C） | 建议目标 | 配套环境变量/配置 |
|------|---------|----------|-------------------|
| Docker | `%LOCALAPPDATA%\Docker` | `D:\Docker` | Docker Desktop 设置/已有数据根；WSL vhdx |
| Android | `%LOCALAPPDATA%\Android` | `D:\Android` | `ANDROID_HOME`, `ANDROID_SDK_ROOT` |
| gradle | `%USERPROFILE%\.gradle` | `D:\DevCache\gradle` | `GRADLE_USER_HOME` |
| m2 | `%USERPROFILE%\.m2` | `D:\DevCache\m2` | Maven settings 可选 |
| jdks | `%USERPROFILE%\.jdks` | `D:\DevCache\jdks` | IDE 内 JDK 表通常跟路径 |
| miniconda3 | `%USERPROFILE%\miniconda3` | `D:\DevCache\miniconda3` | 用户 PATH 若写死需仍指向原路径（junction 可保） |
| npm-global | `%APPDATA%\npm` | `D:\DevCache\npm-global` | `npm config set prefix` 可仍用原路径 |
| npm-cache | `%LOCALAPPDATA%\npm-cache` | `D:\DevCache\npm-cache` | `npm_config_cache` / `npm config set cache` |
| pub | `%LOCALAPPDATA%\Pub` | `D:\DevCache\pub` | `PUB_CACHE`（常见 `...\pub\Cache`） |
| pip | `%LOCALAPPDATA%\pip` | `D:\DevCache\pip` | `PIP_CACHE_DIR` |
| develop | `%USERPROFILE%\develop` | `D:\DevCache\develop` | PATH 中 flutter 等保持原路径即可 |
| TEMP | 用户 TEMP/TMP | `D:\Temp\<user>` | 用户级 `TEMP`,`TMP` |

详细与风险见 `references/migration-map.md`。

### C3. 用户级环境变量（迁移后必配）

```powershell
[Environment]::SetEnvironmentVariable('GRADLE_USER_HOME','D:\DevCache\gradle','User')
[Environment]::SetEnvironmentVariable('PUB_CACHE','D:\DevCache\pub\Cache','User')
[Environment]::SetEnvironmentVariable('npm_config_cache','D:\DevCache\npm-cache','User')
[Environment]::SetEnvironmentVariable('PIP_CACHE_DIR','D:\DevCache\pip\Cache','User')
[Environment]::SetEnvironmentVariable('TEMP','D:\Temp\MECHREVO','User')  # 按用户名改
[Environment]::SetEnvironmentVariable('TMP','D:\Temp\MECHREVO','User')
[Environment]::SetEnvironmentVariable('ANDROID_HOME','D:\Android','User')
[Environment]::SetEnvironmentVariable('ANDROID_SDK_ROOT','D:\Android','User')

npm config set cache 'D:\DevCache\npm-cache' --location=user
# prefix 建议保持原 AppData\Roaming\npm（经 junction 落到 D）
```

提醒用户：**新开终端/重启 IDE** 后 env 才全局生效。

### C4. pagefile（可选，高影响）

- 可同时存在多盘 pagefile；改系统分页文件需管理员，且常需重启。  
- 不要直接删 `C:\pagefile.sys`。  
- 仅在用户明确要求时，指导：系统属性 → 高级 → 性能 → 虚拟内存，或等价 PowerShell/WMI；改完验证 `ExistingPageFiles`。

### C5. 不建议默认迁移

- 整个 `%LOCALAPPDATA%` 或整个 User Profile  
- 正在运行的 MongoDB/MySQL 数据目录（先停服务并改配置）  
- 微信/飞书等强依赖绝对路径且会自愈重建的目录（可迁但要验收更新）  
- Python 官方安装目录（升级/卸载器敏感；优先迁 cache/venv）

---

## 阶段 D：专项说明

### D1. Claude 桌面端 vs Claude Code CLI

| 组件 | 典型位置 | 清理影响 |
|------|----------|----------|
| Claude Code CLI | WinGet `Anthropic.ClaudeCode`、`C:\Windows\claude.exe`、`%USERPROFILE%\.claude` | **不要当桌面缓存清掉** |
| Claude 桌面 Store 包 | `%LOCALAPPDATA%\Packages\Claude_*` | 清缓存影响桌面端 |
| `vm_bundles`（rootfs.vhdx 等） | `...\Claude\vm_bundles` | 桌面本地 VM/Cowork；用户不用则可删腾约数 GB–十余 GB |

用户确认「几乎不用桌面端 VM/Cowork」后，可只删 `vm_bundles`，并验证 `claude --version` 仍可用。

### D2. Docker

- 数据常在 `%LOCALAPPDATA%\Docker` 下 `*.vhdx`（可已 junction 到 `D:\Docker`）  
- 先 `docker system df`，再按用户意图 `docker system prune` / 迁整目录  
- 迁移时尽量退出 Docker Desktop

### D3. 已存在联接

审计时检查 `Get-Item path | Select LinkType,Target`。已是 Junction 则不要重复 robocopy 覆盖目标（除非修复半成品）。

---

## 阶段 E：验收（verify）

必须检查：

1. `C`/`D` 剩余空间变化  
2. 每个迁移源：`LinkType=Junction` 且 Target 正确  
3. 无大量残留 `*.pre-migrate-bak`  
4. 冒烟：
   - `claude --version`（若安装）
   - `npm root -g` / `npm config get cache`
   - `java -version`（若使用）
   - Flutter/Android 路径是否仍存在（若使用）
5. 用户环境变量已写入（`[Environment]::GetEnvironmentVariable(...,'User')`）

脚本：`scripts/verify_junctions.ps1`。

---

## 对话中的执行规范（给 Agent）

1. **先 audit，只报告和建议**；把「预计释放」和「风险」说清楚后停止。  
2. 清理和迁移都必须列出具体路径、动作、影响；等待用户明确同意后，清理传入 `-Execute -ApprovedPath <获批路径>`，迁移才传入 `-Execute`。  
3. 使用 PowerShell 时注意：
   - `foreach { ... } | Sort-Object` 在部分版本会解析失败 → 先赋给 `$results = foreach ...` 再管道。  
   - 长拷贝用 `robocopy /E /COPY:DAT /R:1 /W:1 /XJ /MT:8`；exit code **0–7 成功**，≥8 失败。  
   - 清理 Temp 时 skill 任务日志可能被删，改用盘符 free 差值验收。  
   - **向 `pwsh -File script.ps1 -Param $array` 传数组会按元素拆成多个位置参数而报错**（实测：`A positional parameter cannot be found that accepts argument ...`）。当前会话内改用 `& script.ps1 -Param $array`；跨进程调用则把数组用逗号串成一个字符串参数再拆。
   - `safe_clean.ps1` 对无权删除的路径会返回 `Status=cleaned` 但 `FreedMB=0`（如非管理员清 `SoftwareDistribution\Download`）→ 以 FreedMB/盘符 free 差值为准，别只看 status。
4. 不要求管理员时：用户级 junction/env 通常足够；系统 pagefile/部分 Windows 目录需要提升权限。  
5. 结束后给用户一份简表：做了什么、释放多少、迁到哪里、需要重启哪些程序。  
6. 验收表必须包含「保留完好」项（如飞书 `app` 当前版、VS Code 新版扩展、Playwright 新版本仍在），证明只删了旧版本/缓存、没误删当前版本。  
7. 本 skill 持久化位置约定：源仓库 `E:\ojc-skills\windows-c-drive-cleanup`；Claude Code 安装为 `~\.claude\skills\windows-c-drive-cleanup` 的目录联接，DSH/其他 harness 安装（如 `~\.agents\skills\windows-c-drive-cleanup`）通常也是指向源仓库的目录联接，改源即自动同步。

## 参考文件

- `references/safe-targets.md` — 清理白名单/黑名单与命令片段  
- `references/folder-purposes.md` — 常见 Windows 目录的用途（作用）与默认处置分类知识库  
- `references/migration-map.md` — 迁移映射、env、PATH 注意点  
- `references/machine-profile-mechrevo.md` — 本机（MECHREVO）已落地配置快照  
- `scripts/*.ps1` — 可重复执行的审计/清理/迁移/验收脚本  
