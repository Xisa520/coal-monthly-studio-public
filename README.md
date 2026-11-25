# 煤炭市场月报智能编制平台 · GitHub Pages 公开版 V1.1

基于已有本地FastAPI月报工作台的**公开适配版**。它是独立静态网页，不需要Python后端，适用于分享给所有人访问。

> **没有在此仓库放入任何实际内部月报、内部表格、经营数据、个人信息、AI API Key 或模型权重。** 系统不收集访问者上传的文件。原始文件仅在访问者浏览器内读取。程序从第三方CDN加载固定版本的Excel解析和Word压缩依赖；请按单位的信息安全要求决定是否允许用外部网站处理真实工作数据。

## 一、发布到 GitHub Pages：操作步骤

1. 登录 GitHub，点击右上角 `+` → **New repository**，仓库名推荐 `coal-monthly-studio-public`。选择 **Public**，点击 **Create repository**。
2. **不要将ZIP压缩包作为单个文件上传**。请先解压，把包含 `index.html`、`app.js`、`style.css`、`data/`、`scripts/`、`.github/` 等文件直接上传至仓库根目录。建议使用 **GitHub Desktop** 或 Git 命令，以保证 `.github/workflows/` 等目录完整上传。
3. 命令行方法（在本项目解压后的根目录打开PowerShell）：
   ```powershell
   git init
   git branch -M main
   git add .
   git commit -m "Publish public coal monthly studio"
   git remote add origin https://github.com/你的用户名/coal-monthly-studio-public.git
   git push -u origin main
   ```
4. 进入仓库 **Settings → Pages → Build and deployment → Source**，选择 **GitHub Actions**。
5. 到 **Actions** 页面，等待 `Deploy public GitHub Pages` 工作流成功；如未自动运行，打开工作流点击 **Run workflow**。
6. 发布后的网址一般为：`https://你的用户名.github.io/coal-monthly-studio-public/`。实际URL以 Pages 页面显示为准。

如果第一次运行 GitHub Actions 时提示权限受限，进入仓库 **Settings → Actions → General** 核实允许 Actions 运行。自动采集任务推送提交可能需要在同一页的 **Workflow permissions** 开启 `Read and write permissions`；如权限策略不允许，就保留人工上传/查看官网功能，不启用自动提交。

## 二、网页功能与限制

| 功能 | 公开版 V1.1 |
| --- | --- |
| XLSX、XLSM、XLS、CSV 文件读取 | ✅ 浏览器内读取，不执行VBA |
| 用户选择工作表、日期列、数据列、单位、核验状态 | ✅ |
| 记录数、趋势图、最近值、变化量 | ✅ |
| 标准化数据导出为Excel | ✅ |
| 煤炭月报六章节、摘要、预测表编辑 | ✅ |
| 本浏览器草稿保存 | ✅ 使用 localStorage，仅存报告文本 |
| A4 Word .docx 导出及简单附图 | ✅ 需人工核对分页和字体 |
| 国家统计局公开公告候选信息 | ✅ Actions定期尝试采集，失败则保持上次记录 |
| 大模型自动生成分析内容 | ⚠️ 公开版暂为提示词生成与复制；**不含可直接调用的安全多用户AI后端** |
| 多人账户、共享数据库、统一审批流 | ❌ 需要独立认证与后端服务 |
| 百分百复刻第54期Word版式、宏和全部图表 | ❌ 属于后续版本目标 |

**重要：** Excel文件只有选择日期列、数值列并点击“导入所选指标”后才进入浏览器内指标库。程序不会推测上年历史公布值来冒充国家统计局**可比口径同比**。

## 三、关于国家统计局定时更新

- `.github/workflows/official-data.yml` 默认每天北京时间约10:17尝试更新；GitHub计划任务不是实时任务，可能延迟或暂停。
- 使用 `scripts/nbs.py` 仅访问 `www.stats.gov.cn` 的HTTPS官方域名，解析“能源生产情况”新闻稿；它**不是**国家数据平台官方API。
- 解析结果只有“候选”，不自动归为已核验，也不自动插入任何人的报告。
- 如果官网网页结构变化、限制自动化访问或运行失败，旧数据文件不会被伪造覆盖。
- 工作流将公开候选公告写到 `data/nbs-public.json`。注意 Actions 按计划任务运行时，GitHub 仓库若长期无活动，可能停用计划工作流，需要在 Actions 界面手动启用。

## 四、关于隐私、安全及使用限制

1. 本项目代码可公开，**禁止直接将工作单位未批准公开的材料上传GitHub**。公开Pages站点的所有静态文件可被任何访问者读取。
2. 浏览器内导入的Excel不会由本项目上传到服务器；但是项目当前从CDN加载第三方脚本。若涉及机密文件，不应通过未经批准的公网网页处理。
3. 本版不提供API Key输入框，是为了避免将开发者统一密钥暴露给所有访问者，也避免共享AI额度被滥用。想自动调用大模型，请部署独立网关并加用户认证、配额、调用日志、数据外发授权。
4. GitHub Pages并非私有数据处理平台，`localStorage`只能存小体量报告草稿，不应当作数据库。浏览器清理数据后草稿可能丢失，请保存Word。
5. 公式和单元格值由 Excel解析库读取，**不执行Excel的VBA宏**；图表为新绘制，不会直接复现Excel原有复杂图表或宏按钮。
6. 所有报表、图表与AI提示词仅辅助研究，不保证事实准确；报送前必须复核。

## 五、进一步部署完整服务版

若要多人匿名使用且分别保存上传文件、后台集中存储、自动AI生成、自动抓取、共享研报及用户权限体系：需开发身份认证、多租户数据隔离、PostgreSQL或持久存储、文件隔离、队列任务和审核日志，再将Python后端部署到Render/Railway等云服务或单位授权服务器。**不能直接将原V1.0 FastAPI+SQLite包裸露在公网上。**

## 六、本地预览

静态网站可用 VS Code 的 Live Server 打开；也可在项目根目录运行 `python -m http.server 8000` 后打开 `http://localhost:8000`。联网时需可访问CDN。GitHub Pages上无需在本机启动服务。

## 开源许可

应用自有代码使用MIT许可，第三方Excel解析库SheetJS的许可证及CDN依赖按其各自声明执行；使用者应自行检查后续公开分发合规要求。
