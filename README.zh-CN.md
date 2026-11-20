# Polarion Requirement Assistant 网页版部署说明

这个文件夹是一套可直接部署到 GitHub Pages 的静态网页。无需安装 Node.js、配置服务器或运行构建命令。

## 一、网页能做什么

1. 导入 SharePoint 链接表（Excel 或 CSV）。表格应包含 `File Title`（文件名）和 `Hyperlink`（原始 SharePoint 链接）两列。
2. 选择已同步到电脑的文档文件夹。网页按文件名把本地文档与链接表匹配；点击文件名可打开本次选择的本地文件。
3. 在现有 VS Code 扩展中运行 GitHub Copilot 分析，导出结果 Excel，再将这份 Excel 导入网页，填入 Priority 和 Scope 建议。
4. 在网页中核对、修改各行，导出包含 `File Title`、`Hyperlink`、`Priority`、`Scope` 四列的 Excel。

网页不能直接调用 VS Code 里的 GitHub Copilot，也不会自行分析文档正文。网页不直接读取公司 SharePoint；原始链接必须来自你导入的链接表或手动填写。选中的文件和表格在浏览器内处理，本网页没有向服务器上传它们的代码。

## 二、部署到 GitHub Pages

1. 在 GitHub 新建一个仓库。
2. 解压 `polarion-github-pages.zip`，把解压后看到的 `index.html`、`app.js`、`styles.css`、`icon.png` 和 `vendor` 文件夹上传到仓库**根目录**。不要只上传 ZIP，也不要把整个 `Polarion Assiatant` 项目文件夹上传。
3. 打开该仓库的 **Settings → Pages**。
4. 在 **Build and deployment** 中，将 **Source** 设为 **Deploy from a branch**，分支选 `main`，文件夹选 `/(root)`，然后点 **Save**。
5. 等 GitHub Pages 完成发布，页面会显示访问地址，格式通常为 `https://你的用户名.github.io/仓库名/`。

网页已经使用相对资源路径，放在这种带仓库名的地址下也能加载。

**注意：**GitHub Pages 网站通常公开可访问，即使源码仓库是私有的。请先确认公司的发布规定；不要把客户文档、真实 SharePoint 链接表、账号资料或 VSIX 安装包上传到这个公开仓库。仓库只需要这套网页代码。

## 三、部署后如何使用

1. 打开网页，点击 **Import link list Excel**，选择含文件名和原始 SharePoint 链接的表格。
2. 点击 **Choose synced folder**，选择 SharePoint 同步到电脑的文档文件夹。前两步顺序可以互换。
3. 在 VS Code 的 Polarion Requirement Assistant 扩展中分析文档并导出 Excel；回到网页点击 **Import Copilot results** 选择这份结果。
4. 在 **Review & export** 核对 File Title、Hyperlink、Priority 和 Scope。缺少的原始链接可手动粘贴；无法确认时留空。
5. 点击 **Export Excel** 下载最终表格。正式导入 Polarion 前，还需在公司环境中核对字段映射。

如果不同子文件夹里有同名文件，网页无法可靠地自动判断链接属于哪一个，请逐条人工核对。

## 四、现有云端预览

除 GitHub Pages 上传包外，现有私人云端预览地址是：

https://polarion-assistant-haifeng-2026.haifeng-xu1127.chatgpt.site

该预览目前仅对本账号开放；要在公司电脑使用 GitHub Pages 版本，请按上面的步骤在获准的 GitHub 仓库发布。
