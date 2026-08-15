# 本机配置快照（MECHREVO）

> 记录 2026-07-14 左右落地状态，便于复查。其他机器请以 audit 为准，不要照搬用户名。

## 磁盘

- C: 系统盘（约 200GB 级，曾低至 ~22GB 可用；2026-08-15 低风险清理后 ~43GB 可用）
- D: 数据盘（500GB+，空闲充足）
- E: 资料盘（含 `E:\ojc-skills`）

## 已做安全清理（示例）

- 用户 Temp、应用 updater 残留、npm/pip/.cache、VS Code CachedExtensionVSIXs
- Downloads 中安装包
- Claude 桌面 `vm_bundles`（用户确认几乎不用桌面 VM/Cowork）
- Claude Code CLI 保持可用（`claude --version`）

## 2026-08-15 会话：低风险快速清理（释放 7.08 GB）

C 盘剩余 36.06 → 43.14 GB。已删（全部为缓存/旧版本/日志）：

- 飞书旧版 `AppData\Local\Feishu\7.66.5`（963 MB，`app` 当前版保留）
- VS Code 旧扩展 `openai.chatgpt-26.715.61943`（891 MB）、`anthropic.claude-code-2.1.217`（256 MB）；新版保留
- Playwright 旧 Chromium `chromium-1208` + `chromium_headless_shell-1208`（≈600 MB；1234 保留）
- NVIDIA `DXCache`/`GLCache`（840 MB）
- `.codex\logs_2.sqlite` + wal/shm（755 MB）
- Yarn `Cache`（≈800 MB）
- 夸克 `QianwenInstaller\*.exe` 残留（209 MB）
- `.cache`（1431 MB）、`Code\CachedExtensionVSIXs`（433 MB）、QuarkUpdater（48 MB）、Code Cache/CachedData/logs（29 MB）、INetCache（8 MB）
- 回收站已清空

**未清完/待确认：**

- `C:\Windows\SoftwareDistribution\Download` 130 MB 未删掉（非管理员；`safe_clean.ps1` 显示 cleaned 但 FreedMB=0 的典型陷阱）
- `C:\$WinREAgent` 0.91 GB（8/12 更新回滚暂存，待确认更新完成）
- Claude 桌面 `Packages\Claude_*` 1.98 GB（用户未确认）
- `C:\pagefile.sys` 15 GB 仍在 C（迁 D 需管理员+重启，未做）

**经验沉淀：**

- 报告带「作用」列最受用户认可 → 已写入 SKILL.md A4/A5 + `folder-purposes.md`
- junction 大小陷阱：对 junction 根直测会穿透到 D 盘；robocopy `/L /E /XJ` 是测真实 C 占用的可靠方法
- 向 `pwsh -File script.ps1 -Param $array` 传数组会拆参报错 → 用 `&` 调用

## 已建立目录联接

| 源 | 目标 |
|----|------|
| `C:\Users\MECHREVO\AppData\Local\Docker` | `D:\Docker` |
| `C:\Users\MECHREVO\AppData\Local\Android` | `D:\Android` |
| `C:\Users\MECHREVO\.gradle` | `D:\DevCache\gradle` |
| `C:\Users\MECHREVO\.jdks` | `D:\DevCache\jdks` |
| `C:\Users\MECHREVO\.m2` | `D:\DevCache\m2` |
| `C:\Users\MECHREVO\miniconda3` | `D:\DevCache\miniconda3` |
| `C:\Users\MECHREVO\AppData\Roaming\npm` | `D:\DevCache\npm-global` |
| `C:\Users\MECHREVO\AppData\Local\npm-cache` | `D:\DevCache\npm-cache` |
| `C:\Users\MECHREVO\AppData\Local\Pub` | `D:\DevCache\pub` |
| `C:\Users\MECHREVO\develop` | `D:\DevCache\develop` |
| `C:\Users\MECHREVO\AppData\Local\pip` | `D:\DevCache\pip` |

## 用户环境变量

```text
GRADLE_USER_HOME=D:\DevCache\gradle
PUB_CACHE=D:\DevCache\pub\Cache
npm_config_cache=D:\DevCache\npm-cache
PIP_CACHE_DIR=D:\DevCache\pip\Cache
TEMP=D:\Temp\MECHREVO
TMP=D:\Temp\MECHREVO
ANDROID_HOME=D:\Android
ANDROID_SDK_ROOT=D:\Android
```

## 说明文件

- `D:\DevCache\README-migration.txt`

## 未默认迁移（本机仍可能占 C）

- `C:\pagefile.sys`（注册表曾见 C/D 分页配置；改动需管理员+重启）
- `C:\mongodb`
- `AppData\Local\Programs\Python`
- 微信/飞书/百度等 Roaming 数据
- Claude 桌面端除 vm_bundles 外的其他缓存

## Skill 安装位置

- 源：`E:\ojc-skills\windows-c-drive-cleanup`（`E:\ojc-skills` 为 git 仓库）
- Claude Code：`C:\Users\MECHREVO\.claude\skills\windows-c-drive-cleanup` → 联接到源目录
- DSH/harness：`C:\Users\MECHREVO\.agents\skills\windows-c-drive-cleanup` → 联接到源目录（2026-08-15 确认）
- 两者均为目录联接：**改源即自动同步**
