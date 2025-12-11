# Advanced RDP (MVP)

一个极简的 Windows RDP 客户端 MVP，基于系统自带的 MsTscAx ActiveX 控件（通过 ProgID 动态托管，无需 VS/COM 包装）。支持为每台主机保存凭据（写入 Windows Credential Manager）、本地 JSON 配置的主机列表，嵌入原生 RDP 视图。

## 主要特点
- WPF (.NET 8) + MsTscAx ActiveX（系统自带），无第三方 UI 依赖
- 主机配置存于 `%LOCALAPPDATA%/AdvancedRdp/hosts.json`
- 密码存于 Windows Credential Manager（Generic Credentials，键名：`AdvancedRdp/advancedrdp:{名称}`)
- 左侧列表/表单（名称、地址、用户名、域、剪贴板重定向、保存密码），右侧显示 RDP 会话

## 目录
- `AdvancedRdp.sln`
- `src/AdvancedRdp/AdvancedRdp.csproj`
- `src/AdvancedRdp/MainWindow.xaml(.cs)`：UI 与连接逻辑
- `src/AdvancedRdp/Services/CredentialService.cs`：封装 Credential Manager
- `src/AdvancedRdp/Services/HostStore.cs`：加载/保存主机列表
- `src/AdvancedRdp/Models/HostEntry.cs`：主机数据模型

## 构建与运行
当前环境未安装 .NET SDK（仅有 runtime），需要先安装 SDK：
1) 安装 [.NET 8 SDK](https://dotnet.microsoft.com/download)（已装可忽略）
2) 命令行构建/运行：
   ```powershell
   dotnet build AdvancedRdp.sln
   dotnet run --project src/AdvancedRdp/AdvancedRdp.csproj
   ```
3) 首次运行时，添加主机并勾选“保存密码”，密码会写入 Credential Manager；之后可空密码直接连接。

## 已知事项与下一步
- UI 为 MVP：未做尺寸设置、标签页与重连策略，后续可补充
- 若系统禁用 ActiveX 或移除 MsTscAx，需要启用“远程桌面连接”组件
- 可考虑增加：UDP/GFX/H.264 细化开关、多窗口/多标签、导入导出配置、自更新
