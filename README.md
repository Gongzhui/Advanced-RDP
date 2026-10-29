# Advanced RDP

Windows 远程桌面客户端。2.0 用 [MyGo](https://github.com/egoist/mygo)（纯 Go、无 cgo）重写：界面由 MyGo 原生 UI 用 Direct3D 11 绘制，远程画面由系统自带的远程桌面控件（mstscax.dll）绘制，通过 ATL 控件宿主嵌入窗口。

## 特点

- **快**：单个 10 MB 可执行文件，不依赖 .NET 或 WebView2；主窗口约 80 ms 出现。
- **保存主机**：名称、地址（可带端口）、用户名、域、全屏、共享剪贴板、窗口分辨率；搜索；改名时密码跟着走。
- **密码在系统凭据管理器**：目标名 `AdvancedRdp/advancedrdp:{名称}`，与 1.x 相同。
- **会话**：每台主机一个窗口；再点连接会切换到已打开的会话。窗口模式下改变窗口大小，远程分辨率随之调整（服务器支持时；否则缩放）。
- **全屏**：使用系统连接栏（最小化、还原为窗口、断开，可拖动、可固定）；Windows 键只在全屏时发给远程；Ctrl+Alt+Break 切换全屏。
- **状态清楚**：连接中显示进度；失败时显示系统给出的原因，并自动检查域名、端口和 RDP 握手；可一键重连。网络中断时自动重连。
- **键盘**：Enter 连接，Ctrl+S 保存，Ctrl+N 新建，Ctrl+F 搜索，Ctrl+T 测试连接。

## 数据与 1.x 兼容

- 主机列表：`%LOCALAPPDATA%\AdvancedRdp\hosts.json`，字段与 1.x 相同（新增可选字段 `lastConnected`，1.x 会忽略）。
- 第一次保存前，原文件备份为 `hosts.v1-backup.json`。
- 1.x 只存在 `secrets.json`（DPAPI 加密）里的密码，读取时会迁移到凭据管理器。
- 日志：`%LOCALAPPDATA%\AdvancedRdp\logs\advancedrdp.log`（超过 1 MB 自动轮换）。

## 开发

需要 Go 1.27.3 或更新版本（Windows）。

```powershell
go tool mygo dev          # 开发运行，改代码自动重启
go test ./...             # 单元测试与界面测试（不打开窗口）
go tool mygo vet .        # 检查 MyGo 界面生命周期
go tool mygo build        # 生成 build\windows-amd64\ 下的程序和安装包
```

端到端测试会创建真实的远程桌面控件，并用一个不存在的用户连接本机的远程桌面服务（停在登录界面，不会登录）：

```powershell
$env:ARDP_E2E = 1; go test ./internal/rdp -v
```

试用开发版而不碰真实数据：设置 `ARDP_DATA_DIR`（数据目录）和 `ARDP_VAULT_APP`（凭据名前缀）。

## 目录

- `main.go`、`view_main.go`：主窗口（主机列表与表单）。
- `session.go`：会话窗口、连接状态、全屏与分辨率。
- `internal/rdp`：COM 与 ActiveX 宿主、事件接收器、远程桌面控件的封装。
- `internal/store`：`hosts.json` 的读写与校验。
- `internal/cred`：凭据管理器与 1.x 的 `secrets.json`。
- `internal/netprobe`：地址解析与连接诊断（DNS、TCP、RDP 握手）。
- `tools/genicon`：生成 `resources/icon.png`。
