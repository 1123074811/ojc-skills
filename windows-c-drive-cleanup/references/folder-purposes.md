# 常见 Windows 目录用途知识库

> 审计报告「作用」列的标注来源。处置分类：🗑️ 可清理 / 🔄 可迁移 / ⚠️ 勿动 / 🧩 卸载决策。
> 未收录的目录注明「未知，需人工判断」，**不要编造用途**。本表由 MECHREVO 实测整理，其他机器可参考并扩展。

## 一、C:\ 顶层

| 路径 | 作用 | 默认处置 |
|------|------|----------|
| `C:\Windows` | 操作系统本体（含 WinSxS、System32、Installer、SoftwareDistribution） | ⚠️ 勿手动删文件 |
| `C:\Program Files` | 64 位软件安装目录 | 🧩 卸载不用的软件 |
| `C:\Program Files (x86)` | 32 位软件安装目录 | 🧩 同上 |
| `C:\ProgramData` | 各程序公共数据（Package Cache、应用数据） | ⚠️ 保留 |
| `C:\pagefile.sys` | 虚拟内存页面文件 | 🔄 可迁 D（需管理员+重启） |
| `C:\hiberfil.sys` | 休眠文件（`powercfg /h off` 可关） | 🔄/🗑️ 关休眠后消失 |
| `C:\swapfile.sys` | UWP 应用交换文件，很小 | ⚠️ 保留 |
| `C:\$WinREAgent` | 系统功能更新的回滚暂存（Rollback/Scratch） | 🗑️ 确认更新完成后可删 |
| `C:\$Recycle.Bin` | 回收站（每用户一个子目录） | 🗑️ 清空回收站 |
| `C:\System Volume Information` | 系统还原点/卷影副本（非管理员不可读） | ⚠️ 通过「系统保护」管理，勿手删 |
| `C:\Recovery` | 恢复分区挂载内容 | ⚠️ 勿动 |
| `C:\PerfLogs` | 性能日志 | ⚠️ 保留（通常 ≈0） |
| `C:\Documents and Settings` | 旧系统兼容联接 → `C:\Users` | ⚠️ 系统联接，勿动 |
| `C:\Temp` / `C:\tmp` | 自定义临时目录 | 🗑️ 清空内容 |
| `C:\Users` | 用户配置文件根 | 下钻标注（见第二节） |
| 其他自定义目录（如 `C:\mongodb`、`C:\WCH.CN`） | 视安装内容而定（WCH.CN=沁恒芯片驱动） | 🧩/🔄 视情况 |

## 二、C:\Users\<user> 顶层

| 路径 | 作用 | 默认处置 |
|------|------|----------|
| `AppData` | 应用数据（Local/Roaming/LocalLow，见第三节） | 下钻标注 |
| `Documents` / `Desktop` / `Pictures` / `Videos` / `Music` | 个人文档 | ⚠️ 勿删 |
| `Downloads` | 下载目录 | 🗑️ 仅安装包（用户确认后） |
| `.vscode` | VS Code 扩展与数据（extensions 常藏旧版本） | 🗑️ 删旧版扩展目录 |
| `.codex` | OpenAI Codex CLI 数据（logs_*.sqlite、sessions、cache） | 🗑️ 旧日志库可删；配置保留 |
| `.cache` | 通用缓存（codex-runtimes 等） | 🗑️ 可重建 |
| `.gradle` `.m2` `.jdks` `miniconda3` `develop` `.npm` | 开发缓存/工具链（常已 junction 到 D） | 🔄 迁移（junction） |
| `.claude` | Claude Code CLI 配置与凭据 | ⚠️ 勿当缓存删 |
| `.config` `.local` `.nuget` `.conda` | 各工具配置/缓存 | 🗑️/⚠️ 视内容 |
| 各类应用点目录（`.workbuddy` `.campusmind` `.djl.ai` `.lingma` 等） | 对应工具的数据/模型（.campusmind 含本地 onnx 模型） | ⚠️ 保留为主 |
| `bin` | 用户自建工具目录 | ⚠️ 保留 |

## 三、AppData\Local（典型目录）

| 路径 | 作用 | 默认处置 |
|------|------|----------|
| `Programs` | 用户级安装的程序（Python、VS Code、Quark、Apifox、Trae 等） | 🧩 卸载不用的 |
| `Packages` | UWP/商店应用（Claude 桌面、终端等） | 🧩 卸载；Claude 桌面 VM 见下 |
| `Temp` | 用户临时目录（env TEMP 改到 D 后常为空） | 🗑️ 清内容 |
| `Docker` | Docker Desktop 数据（WSL vhdx，常已 junction 到 D） | 🔄 迁移 |
| `Microsoft\Edge\User Data` | Edge 浏览器缓存/数据 | 🗑️ 清缓存，⚠️ 留登录 |
| `Microsoft\WinGet\Packages` | winget 安装包副本 | 🗑️ 可清（重装时重下） |
| `NVIDIA` | 显卡 DXCache/GLCache 着色器缓存 | 🗑️ 清内容（自动重建） |
| `ms-playwright` | Playwright 浏览器（chromium-<版本> 多套并存） | 🗑️ 删旧版本目录 |
| `Yarn\Cache` | Yarn 缓存 | 🗑️ `yarn cache clean` |
| `pnpm` | pnpm store/缓存 | 🗑️ 视配置（确认非主 store） |
| `Feishu`（飞书） | 程序本体；`<版本号>` 为旧版、`app` 为当前版 | 🗑️ 删旧版号目录 |
| `Quark`（夸克） | 浏览器数据；`User Data\QianwenInstaller\*.exe` 为安装残留 | 🗑️ 删安装残留 exe |
| `微信开发者工具` | 微信小程序 IDE | ⚠️/🧩 视使用 |
| `JetBrains` | IDE 索引/缓存 | 🗑️ 清 caches（会重建） |
| `*-updater`（QuarkUpdater 等） | 应用更新器残留 | 🗑️ 可清 |
| `pip` `Pub` `npm-cache` | 包缓存（常已 junction 到 D） | 🔄 迁移 |
| `OpenAI` | Codex CLI 程序本体 | ⚠️ 保留 |

## 四、AppData\Roaming（典型目录）

| 路径 | 作用 | 默认处置 |
|------|------|----------|
| `npm` | npm 全局包（常已 junction 到 D） | 🔄 迁移 |
| `Tencent` | 微信/QQ/电脑管家数据（xwechat、QQPCMgr） | ⚠️ 保留（含登录/聊天数据） |
| `baidu` | 百度网盘模块/引擎 | ⚠️ 保留 |
| `JetBrains` | IDE 配置与索引 | ⚠️ 配置保留；索引可清 |
| `Code` | VS Code 配置 + `CachedExtensionVSIXs`/`Cache`/`logs` | 🗑️ 清缓存子目录 |
| `QoderCN` / `Qoder` | Qoder IDE 缓存库 | 🗑️ 缓存部分可清 |
| `adspower_global` | AdsPower 浏览器内核（chrome_<版本>） | 🗑️ 旧版本内核可删 |
| `bilibili` / `douyin` | 对应应用缓存（IndexedDB 等） | 🗑️ 缓存可清，⚠️ 登录保留 |
| `DingTalk` / `QQ` / `QQEX` | 应用数据 | ⚠️ 保留 |
| `greencore` | 安全软件 Chrome-bin 引擎 | ⚠️/🧩 视使用 |
| `Python` | pip 相关用户配置 | ⚠️ 保留 |

## 五、商店应用/桌面端专项

| 路径 | 作用 | 默认处置 |
|------|------|----------|
| `Packages\Claude_*` | Claude 桌面 Store 包 | 🧩 不用可卸载 |
| `...\Claude\vm_bundles` | 桌面端本地 VM/Cowork（rootfs.vhdx） | 🗑️ 仅用户确认不用桌面 VM；**不影响 Claude Code CLI** |
| `...\Claude\claude-code` | 桌面内嵌 claude-code 二进制 | ⚠️ 随桌面端 |

## 六、其他用户 profile（C:\Users 下）

| 路径 | 说明 | 默认处置 |
|------|------|----------|
| `<旧用户名>` / `<中文名>` / `<中文名>1` | 旧账号残留（常为几 MB 空壳） | 🧩 确认无数据后可用「系统属性→用户配置文件」删除 |
| `Public` `Default` `Default User` | 系统默认 profile | ⚠️ 勿动 |
| `WsiAccount` | Windows Store 服务账号 | ⚠️ 勿动 |
| `AppData`（直接挂在 C:\Users 下） | 异常目录——某程序把 AppData 写错位置 | 🧩 查清来源后清理 |
