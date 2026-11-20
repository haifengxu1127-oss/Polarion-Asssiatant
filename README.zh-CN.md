# Polarion Requirement Assistant

这是一个带 Node.js 云端服务的网页应用。使用者只需打开云端网址并用 GitHub 登录，无需安装 VS Code 或 VSIX 扩展。GitHub Copilot 按每个登录用户自己的账号授权使用。

## 功能

1. 导入包含 `File Title` 和 `Hyperlink` 的 SharePoint 链接表（Excel 或 CSV）。
2. 选择本地同步的 PDF、DOCX、TXT 或 Markdown 文档文件夹。
3. 使用 GitHub 登录、填写允许的 Scope 值，在云端调用 Copilot 分析。
4. 核对和修改结果，导出 `File Title`、`Hyperlink`、`Priority`、`Scope` 四列的 Excel。

网页不会直接连接 SharePoint，链接需要从你确认过的链接表导入。当前云端版本不支持没有可提取文本的扫描版 PDF。

## 部署云端服务

GitHub Pages 只能托管静态文件，不能运行 Copilot 后端。请将整个仓库作为 **Docker Web 服务**部署（仓库根目录包含 `Dockerfile`）。Node 服务会在同一网址下同时提供网页和 API，用户可直接用浏览器访问。可选择任何支持 Docker 的云平台，但须符合组织的数据和托管政策。

1. 创建 GitHub OAuth App，将回调地址设为 `https://你的服务域名/auth/callback`。
2. 从仓库根目录使用附带的 `Dockerfile` 部署服务。
3. 在云平台的密钥/环境变量设置中配置：
   - `APP_BASE_URL`：公开 HTTPS 根网址，不含路径，例如 `https://requirements.example.com`
   - `GITHUB_CLIENT_ID`
   - `GITHUB_CLIENT_SECRET`
   - `GITHUB_ALLOWED_USERS`：允许使用服务的 GitHub 登录名，多个账号用逗号分隔
   - `PORT`：通常由云平台提供；默认值为 `3000`
4. 访问云服务网址。每位使用者都要用自己的 GitHub 账号登录，并且该账号需要有 Copilot 使用权限。

本地开发需要 Node.js `22.12` 或更新版本。将 `server/.env.example` 复制为 `server/.env`，填入 OAuth 配置；本地回调地址可使用 `http://localhost:3000/auth/callback`。然后在 `server/` 目录运行 `npm ci` 和 `npm start`。

不要提交 `.env` 文件，也不要把 OAuth 密钥写进浏览器代码。不要配置供所有用户共用的 `COPILOT_GITHUB_TOKEN`；SDK 会在每次分析时使用当前登录用户的 token。

## 数据、隐私和限制

- 选中的文档会上传至云服务。服务提取文档文本并以登录用户的 GitHub Copilot 身份发送给 Copilot。使用量和可用性受该用户的订阅、组织策略和使用限额约束。
- HTTP 服务代码不会把上传的原始文档写入磁盘；但 Copilot SDK 运行时可能会在云主机存储会话状态或对话记录。运行时配置为空闲 15 分钟后清理会话，因此不能将此原型视为“零留存”。上传真实客户资料前，请先确认 Copilot、云平台日志、数据留存和公司政策。
- OAuth token 和应用会话目前只保存在服务进程内存中，闲置后过期，服务重启后需要重新登录。初版适用于单实例；生产多实例部署还需要共享会话存储、更完善的运维监控和经批准的数据留存设计。
- 单个文件上限为 10 MB，提取文本上限为 60,000 字符。扫描版 PDF 需要 OCR，当前不支持。
- 应用不会从 SharePoint 下载文件，也不会直接写入 Polarion。

GitHub Pages 地址仍然只是静态版，不能执行云端 Copilot 分析。免安装入口是部署后的 Node 云服务网址。

## 第三方代码

`vendor/xlsx.full.min.js` 为 SheetJS Community Edition 0.18.5，采用 Apache 2.0 许可证。许可证文本见 `vendor/LICENSE-Apache-2.0.txt`。
