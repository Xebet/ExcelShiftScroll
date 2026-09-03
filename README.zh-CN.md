# ExcelShiftScroll

ExcelShiftScroll 为 Windows 桌面版 Microsoft Excel 增加常见的 **Shift + 鼠标滚轮** 操作：

- Shift + 滚轮向上：向左横向滚动。
- Shift + 滚轮向下：向右横向滚动。
- 默认启用，每个滚轮刻度滚动 3 列。

[English](README.md)

![演示占位图：未来将展示在 Excel 中按 Shift 加滚轮进行横向滚动](docs/assets/demo-placeholder.svg)

> 完成并记录兼容性测试矩阵后，将以真实演示 GIF 替换此占位图。

## 不受影响的原生行为

普通滚轮仍纵向滚动；Ctrl + 滚轮仍缩放；Ctrl + Shift + 滚轮保留 Excel 原生处理。Alt/Windows 键组合、原生横向滚轮事件、其他应用、后台 Excel，以及功能区、公式栏、任务窗格、对话框和 VBA 编辑器区域都不会被转换。

加载项没有外部进程、不联网、不含遥测，也不读取工作簿内容。

## 支持范围

首要目标为 Windows 11、64 位 Microsoft 365 桌面版 Excel。构建同时生成 32 位 XLL，并按设计兼容 Windows 10 与 Excel 2019、2021、2024。Excel 网页版、macOS 和移动版不支持。当前实测与待测组合请见 [兼容性说明](docs/compatibility.md)。

## 下载与位数判断

从 [GitHub Releases](https://github.com/Xebet/ExcelShiftScroll/releases/latest) 下载与 Excel 位数匹配的 ZIP，并使用 `SHA256SUMS.txt` 校验。

在 Excel 中打开 **文件 → 帐户 → 关于 Excel**，第一行会标明 32 位或 64 位。不要用 Windows 位数代替判断；64 位 Windows 也可能安装 32 位 Excel。

## 安装

1. 退出所有 Excel 进程。
2. 把 ZIP 解压到稳定的本地目录，例如 `%LocalAppData%\ExcelShiftScroll`。
3. 右键 `.xll` → **属性**；如果看到 **解除锁定/Unblock**，勾选后确定。带“来自 Internet”标记的文件可能被 Office 阻止。
4. 启动 Excel，打开 **文件 → 选项 → 加载项**。
5. 底部选择 **Excel 加载项**，点击 **转到**，再点击 **浏览**。
6. 64 位 Excel 选择 `ExcelShiftScroll64.xll`；32 位 Excel 选择 `ExcelShiftScroll32.xll`。
7. 把鼠标放在工作表网格，按住 Shift 并转动纵向滚轮。

无需管理员权限，也无需另装 .NET 运行库。0.1.0 发布文件未做 Authenticode 代码签名，Windows SmartScreen 或 Office 可能提示未知发布者。不要关闭系统安全功能；只从本仓库 Release 获取文件，校验 SHA-256 后仅解除该文件的锁定。

## 功能入口

在 Excel 功能区的 **加载项** 选项卡中找到 **Shift Scroll**：

- 立即启用或暂停；
- 每刻度选择滚动 1、2、3、5 或 10 列；
- 反转方向；
- 恢复默认值；
- 查看版本和运行状态。

设置保存在 `%LocalAppData%\ExcelShiftScroll\settings.json`，不会写入工作簿。

## 卸载

1. 打开 **文件 → 选项 → 加载项**。
2. 选择 **Excel 加载项 → 转到**，取消勾选 ExcelShiftScroll。
3. 退出所有 Excel 进程。加载项关闭时会解除进程内钩子，不会留下后台进程。
4. 删除解压目录；如需清除设置和可选诊断日志，再删除 `%LocalAppData%\ExcelShiftScroll`。

## 常见问题

- **“不是有效的加载项”**：通常是 XLL 位数与 Excel 不匹配。
- **文件被阻止**：关闭 Excel，在 XLL **属性** 中解除锁定后重试；不要全局降低信任中心安全级别。
- **找不到功能区按钮**：检查 **文件 → 选项 → 加载项 → 禁用项目**。
- **没有横向滚动**：确认鼠标在工作表网格、只按下 Shift、加载项已启用，且工作表可以横向滚动。
- **公式栏/功能区/对话框/任务窗格不触发**：这是安全边界的预期行为。
- **企业策略禁用 XLL**：请管理员批准已校验文件或部署签名版本。本项目不绕过策略。
- **需要诊断**：退出 Excel 后把设置中的 `diagnosticsEnabled` 改为 `true`。日志只含时间、固定事件名和异常类型，排查后请关闭。

## 从源码构建

需要 Windows、.NET 8 SDK（仅作为构建 SDK）和 PowerShell。加载项目标框架为 Windows 自带的 .NET Framework 4.8。

```powershell
dotnet restore ExcelShiftScroll.sln --configfile NuGet.Config
dotnet build ExcelShiftScroll.sln --configuration Release --no-restore
dotnet test ExcelShiftScroll.sln --configuration Release --no-build --no-restore
./build/Package-Release.ps1 -Version 0.1.0
```

打包后的 XLL 位于 `src/ExcelShiftScroll/bin/Release/net48/publish`，发布 ZIP 和校验值位于 `artifacts/release`。详细设计和验证方式见 [架构](docs/architecture.md)、[测试](docs/testing.md) 和 [ADR](docs/adr/)。

## 隐私、安全和 clean-room 声明

ExcelShiftScroll 完全离线，不读取工作簿名称、路径、公式、值、选区或 VBA，不记录字符按键、鼠标坐标或滚轮历史，不含遥测、更新器或远程代码，不修改信任中心。线程级鼠标钩子只存在于加载它的 Excel 进程，并在关闭时解除。详见 [SECURITY.md](SECURITY.md)。

[OfficeScroll](https://github.com/T800G/OfficeScroll) 仅作为相关历史项目及用户体验/兼容性参考。本项目是独立 clean-room 实现，没有复制、翻译、反编译、改写或逐行移植 OfficeScroll 的 CPOL 源码、二进制或资源。记录见 [clean-room 说明](docs/clean-room.md)。

运行时依赖 Excel-DNA 1.9.0（zlib 许可证）。完整第三方声明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)；本项目源码采用 [MIT 许可证](LICENSE)。
