# 选款助手：上传到 Render

发布后，把 Render 生成的网址发给别人，对方用浏览器打开即可使用。整个过程通过网页完成，不需要安装开发软件或运行终端命令。

## 1. 解压上传包

双击 `选款助手-Render上传包.zip`，打开解压后的文件夹；也可以直接使用项目中现成的 `Render上传包` 文件夹。里面应有：

```text
public/
└── index.html
render.yaml
README.md
```

`public/index.html` 是已经打包完成的网站。上传包不包含商品 Excel；这里用到的 `render.yaml` 已准备好，按下面步骤操作时不需要编辑它。

## 2. 上传到 GitHub

1. 登录 [GitHub](https://github.com/)，右上角 **＋ → New repository**。
2. Repository name 填 `sku-selection`，可见性选择 **Private**，勾选 **Add a README file**，点击 **Create repository**。建立仓库的界面说明见 [GitHub 官方文档](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-new-repository)。
3. 在新仓库首页点击 **Add file → Upload files**。
4. 从上传包文件夹中，把 **`public` 文件夹、`render.yaml`、`README.md` 三项一起拖进去**。保留 `public` 文件夹这一层；上传的 README 可以替换刚才自动生成的 README。
5. 提交说明填写 `上线选款助手`，选择 **Commit directly to the main branch**，再点击 **Commit changes**。文件上传步骤见 [GitHub 官方文档](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository)。

上传完成后，仓库首页应直接看到 `public`、`render.yaml` 和 `README.md`，点进 `public` 能看到 `index.html`。**不要只上传 ZIP，也不要把外层 `Render上传包` 文件夹一起拖进去。**

## 3. 在 Render 创建网站

1. 登录 [Render 控制台](https://dashboard.render.com/)。
2. 点击 **New → Static Site**。
3. 连接 GitHub；在授权页面允许 Render 访问刚才的 `sku-selection` 仓库，然后选择该仓库并点击 **Connect**。
4. 按下表填写。Render 的各字段作用见 [官方部署说明](https://render.com/docs/your-first-deploy)。

| Render 字段 | 填写内容 |
| --- | --- |
| Name | `sku-selection`；若重名，可改成其他英文名称 |
| Branch | `main`；如果 GitHub 中的分支名不同，选择实际分支 |
| Root Directory | **留空** |
| Build Command | `test -s public/index.html` |
| Publish Directory | `public` |

Build Command 直接复制上表内容，前后不用加引号。网站已经打包完成，这条命令只检查首页文件是否存在。不需要设置 Start Command、数据库或环境变量。

5. 点击页面底部 **Deploy Static Site / Create Static Site**（按钮名称可能随界面版本变化）。
6. 等待部署成功，打开页面上显示的 `https://…onrender.com` 网址。这个网址就是发给别人使用的链接。Render 为静态网站提供网址，并会在所选分支更新后自动重新部署。[Render 静态网站说明](https://render.com/docs/static-sites)

## 4. 打开网址后怎么用

1. 每位使用者导入自己电脑上的一份或多份 `.xlsx`。
2. 搜索货号并逐个加入清单，或直接粘贴一批货号。
3. 点击导出，下载包含对应货号完整行、全部颜色及图片的 Excel。

Excel 在使用者自己的浏览器中处理，不会上传到 Render，也不会在不同使用者之间共享；刷新或关闭页面后，需要重新选择原文件。当前网站没有登录入口，拿到网址的人可以打开工具，即使 GitHub 仓库设置为 Private。

原表中的 WPS 单元格图片按原有形式复制。你目前这两份资料导出的图片表建议用 **WPS** 打开；不支持 `DISPIMG` 的软件可能显示公式，这属于原表图片格式的兼容性问题。

## 5. 以后更新网站

拿到新版本上传包后，在 GitHub 仓库中打开 `public` 文件夹，点击 **Add file → Upload files**，上传新版 `index.html` 覆盖同名文件，并提交到 Render 使用的分支。Render 会自动部署更新，原网址保持不变。[自动部署说明](https://render.com/docs/static-sites)

## 遇到问题时

- **Render 找不到仓库：**检查 GitHub 授权是否包含 `sku-selection`；Private 仓库需要明确授予 Render 访问权限。
- **部署失败，提示找不到首页或目录：**检查仓库路径应为 `public/index.html`，而不是 `Render上传包/public/index.html`；Root Directory 留空，Publish Directory 填 `public`。
- **看不到刚上传的版本：**检查文件已经提交成功，而且 Render 的 Branch 与 GitHub 文件所在的分支一致；等最新部署成功后再刷新网页。
