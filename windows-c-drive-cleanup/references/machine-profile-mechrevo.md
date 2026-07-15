# 本机配置快照（MECHREVO）

> 记录 2026-07-14 左右落地状态，便于复查。其他机器请以 audit 为准，不要照搬用户名。

## 磁盘

- C: 系统盘（约 200GB 级，曾低至 ~22GB 可用）
- D: 数据盘（500GB+，空闲充足）
- E: 资料盘（含 `E:\ojc-skills`）

## 已做安全清理（示例）

- 用户 Temp、应用 updater 残留、npm/pip/.cache、VS Code CachedExtensionVSIXs
- Downloads 中安装包
- Claude 桌面 `vm_bundles`（用户确认几乎不用桌面 VM/Cowork）
- Claude Code CLI 保持可用（`claude --version`）

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

- 源：`E:\ojc-skills\windows-c-drive-cleanup`
- Claude Code：`C:\Users\MECHREVO\.claude\skills\windows-c-drive-cleanup` → 联接到源目录
