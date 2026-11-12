# 📊 会计工作台（Excel-organize）

这个 Windows / Mac 桌面软件每天自动整理新收到的 Excel，用 DeepSeek 归纳内容，然后按 **日期 / 类别** 归档，需要时可以用一句中文让 AI 找到对应的表格。

## 下载安装

到 **[Releases](../../releases/latest)** 下载：

- `ExcelLibrary_Setup_x.y.z.exe`：Windows 安装包，双击安装，不需要管理员权限
- `ExcelLibrary_Portable_x.y.z.zip`：Windows 绿色版，解压后运行 `ExcelLibrary.exe`
- `ExcelLibrary_Mac_AppleSilicon_x.y.z.dmg` / `ExcelLibrary_Mac_Intel_x.y.z.dmg`：Mac 版，打开后拖进「应用程序」。第一次打开的方法见 [使用说明.md](使用说明.md)

首次打开：点顶部 **设置**，填写 DeepSeek API Key，再点 **保存并测试连接**；然后点 **打开收件箱**，把 Excel 放进去。

## 功能

- 自动处理收件箱里的新 Excel（xlsx / xlsm / xls / csv）：AI 生成标题、摘要、标签，并判断业务日期和类别
- 按 `文件库\类别\年-月\` 自动归档，并生成可点击打开的 `文件库总目录.xlsx`
- 按类别、月份、关键词、日期范围筛选，表格里的内容也能搜到
- **AI 找文件**：例如“上周华东区的销售日报”“10 月和某供应商的报价”
- **图片识别并登记**：在 AI 助手里粘贴或拖入发票、收据、回单等图片，AI 识别后整理到对应表格的末尾；你核对、修改并确认后才写入，写入前自动备份，可撤销
- **AI 问表格**：勾选几个 Excel 后直接提问其中的内容，求和、排名、筛选、跨文件对比都由本地精确计算
- 删除文件先进回收站，可以恢复或彻底删除；单个、批量删除都支持
- 开机后台运行、每天定时整理，内容相同的文件会自动去重

详细说明见 [使用说明.md](使用说明.md)。

## 构建

每次推送到 `main`，GitHub Actions 都会自动打包、做冒烟测试，并把安装包发布到与 `core.py` 中 `VERSION` 对应的 Release（发新版时先修改 VERSION）。在本地打包的方法：Windows 上装好 Python 3.12，双击 `打包成exe.bat`；Mac 上运行 `./打包Mac版.sh`。
