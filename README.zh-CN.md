# ExcelShiftScroll

ExcelShiftScroll 为 Windows 桌面版 Microsoft Excel 增加常见的 **Shift + 鼠标滚轮** 操作：

- Shift + 滚轮向上：向左横向滚动。
- Shift + 滚轮向下：向右横向滚动。
- 默认启用，使用 Excel 自身的平滑滚动通道，每个滚轮刻度默认为约 3 列的距离。

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

1. 退出所有 Excel 进程。不要从 ZIP 内部或 `%TEMP%` 临时目录加载 XLL。
2. 右键下载的 ZIP → **属性**；如有 **解除锁定/Unblock**，勾选后确定，然后再解压。
3. 在解压目录中右键 `Install-CurrentUser.ps1`，选择 **使用 PowerShell 运行**。它会先校验 XLL 的 SHA-256，在暂存文件上解除 Internet 标记，再安装到 `%LocalAppData%\ExcelShiftScroll\AddIn`。原有 XLL 会保留为带时间戳的 `.bak`；安装后校验失败会恢复原文件。不会修改注册表、信任中心或 Excel 加载项列表。
4. 启动 Excel，打开 **文件 → 选项 → 加载项**。
5. 底部选择 **Excel 加载项**，点击 **转到**，再点击 **浏览**。不要进入 Office 商店的“我的加载项”页面。
6. 从稳定安装目录选择 `ExcelShiftScroll64.xll`（64 位）或 `ExcelShiftScroll32.xll`（32 位）。
7. 把鼠标放在工作表网格，按住 Shift 并转动纵向滚轮。

手动方法：解压到稳定本地目录，右键 XLL → **属性 → 常规 → 解除锁定 → 确定**，然后在 Excel 中浏览这份文件。发布包内的 `INSTALL.txt` 有完整步骤。[Microsoft 官方说明当前 Excel 默认阻止来自不受信任位置的 XLL](https://support.microsoft.com/en-US/Excel/excel-is-blocking-untrusted-xll-add-ins-by-default)。

无需管理员权限，也无需另装 .NET 运行库。发布文件未做 Authenticode 代码签名，Windows SmartScreen 或 Office 可能提示未知发布者。不要关闭系统安全功能；只从本仓库 Release 获取文件，校验 SHA-256 后仅解除该文件的锁定。

### 升级与恢复旧版本

关闭 Excel 后运行新包的安装脚本。若此前已从稳定安装路径加载，重开 Excel 即可；否则先取消勾选旧 XLL，再浏览稳定目录中的新 XLL。“About / 关于”会显示实际加载路径与版本，不能仅凭文件夹名称判断版本。一个“Excel 加载项”条目加一个同名“COM 加载项”功能区辅助条目可能是正常现象，不要仅因同名就删除辅助组件。

如需恢复备份：关闭 Excel，另行保留当前 XLL，把安装目录中选定的带时间戳 `.bak` 复制回原来的 `ExcelShiftScroll64.xll` 或 `ExcelShiftScroll32.xll` 名称，再启动 Excel。设置保留；备份不会自动清理。SHA-256 只能校验完整性，不能代替发布者签名。本项目仍完全离线，不添加在线更新功能。

## 功能入口

在 Excel 功能区的 **加载项** 选项卡中找到 **Shift Scroll**：

- 立即启用或暂停；
- 每刻度选择约 1、2、3、5 或 10 列的滚动距离；
- 反转方向；
- 恢复默认值；
- 查看版本、运行状态、Excel 位数和实际加载的 XLL 路径；
- 打开当前加载项所在目录。

设置保存在 `%LocalAppData%\ExcelShiftScroll\settings.json`，不会写入工作簿。

配置缺失字段会使用默认值，明确保存的 `false` 不会被覆盖。保存失败时保留原设置并显示提示。多个 Excel 进程保存时会互斥写入，以最后一次成功保存的整份设置为准；已打开的进程不会实时同步其他进程的设置。

## 卸载

1. 打开 **文件 → 选项 → 加载项**。
2. 选择 **Excel 加载项 → 转到**，取消勾选 ExcelShiftScroll。
3. 退出所有 Excel 进程。加载项关闭时会解除进程内钩子，不会留下后台进程。
4. 运行发布包内的 `Uninstall-CurrentUser.ps1` 删除已安装的 XLL，设置与备份会保留。旧的手动解压副本需另外删除；如需清除设置、备份和可选诊断日志，再删除 `%LocalAppData%\ExcelShiftScroll`。

## 常见问题

- **“不是有效的加载项”**：通常是 XLL 位数与 Excel 不匹配。
- **显示“来源不受信任”**：浏览器下载或临时目录中的 XLL 通常带 Mark-of-the-Web。关闭 Excel，运行发布包内的当前用户安装脚本，或对稳定目录中的 XLL 执行 **属性 → 解除锁定**，再重新浏览加载。不要关闭 `BlockXLLFromInternet` 或降低信任中心安全级别。
- **列表仍有 `%TEMP%` 旧条目**：先尝试勾选这个已失效条目；Excel 若询问是否从列表删除，请选择“是”，然后浏览稳定目录中的新文件。
- **找不到功能区按钮**：检查 **文件 → 选项 → 加载项 → 禁用项目**。
- **没有横向滚动**：确认鼠标在工作表网格、只按下 Shift、加载项已启用，且工作表可以横向滚动。
- **仍然没有平滑过渡**：正常情况下，ExcelShiftScroll 会保留高精度滚轮增量并交给 Excel 自身的横向滚动引擎。Windows 横向滚轮配置为零/整屏或发送原生消息失败时，才使用整列回退路径。旧版 Excel 即使收到原生消息也可能没有平滑动画；加载项不会检测动画支持情况，也没有自行实现缓动。
- **公式栏/功能区/对话框/任务窗格不触发**：这是安全边界的预期行为。
- **企业策略禁用 XLL**：请管理员批准已校验文件或部署签名版本。本项目不绕过策略。
- **需要诊断**：退出 Excel 后把设置中的 `diagnosticsEnabled` 改为 `true`。日志只含时间、固定事件名和异常类型，排查后请关闭。

## 从源码构建

需要 Windows、.NET 8 SDK（仅作为构建 SDK）和 PowerShell。加载项目标框架为 Windows 自带的 .NET Framework 4.8。

```powershell
dotnet restore ExcelShiftScroll.sln --configfile NuGet.Config
dotnet build ExcelShiftScroll.sln --configuration Release --no-restore
dotnet test ExcelShiftScroll.sln --configuration Release --no-build --no-restore
./tests/install/Test-Installer.ps1
./build/Package-Release.ps1 -Version 0.2.1
```

打包后的 XLL 位于 `src/ExcelShiftScroll/bin/Release/net48/publish`，发布 ZIP 和校验值位于 `artifacts/release`。详细设计和验证方式见 [架构](docs/architecture.md)、[测试](docs/testing.md) 和 [ADR](docs/adr/)。

## 隐私、安全和 clean-room 声明

ExcelShiftScroll 完全离线，不读取工作簿名称、路径、公式、值、选区或 VBA，不记录字符按键、鼠标坐标或滚轮历史，不含遥测、更新器或远程代码，不修改信任中心。线程级鼠标钩子只存在于加载它的 Excel 进程，并在关闭时解除。详见 [SECURITY.md](SECURITY.md)。

[OfficeScroll](https://github.com/T800G/OfficeScroll) 仅作为相关历史项目及用户体验/兼容性参考。本项目是独立 clean-room 实现，没有复制、翻译、反编译、改写或逐行移植 OfficeScroll 的 CPOL 源码、二进制或资源。记录见 [clean-room 说明](docs/clean-room.md)。

运行时依赖 Excel-DNA 1.9.0（zlib 许可证）。完整第三方声明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)；本项目源码采用 [MIT 许可证](LICENSE)。
