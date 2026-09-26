# 选款助手 · GitHub 上传文件

网站已经打包完成，只需上传以下三项：

```text
public/
└── index.html
render.yaml
README.md
```

## 上传到 GitHub

打开你的 `excel` 仓库，选择 **Add file → Upload files**，把 `public` 文件夹、`render.yaml` 和 `README.md` 拖进去，覆盖同名文件并提交。

保留 `public` 文件夹这一层，仓库中的网页路径必须是 `public/index.html`。若使用压缩包，请先解压，再上传其中这三项；不要直接上传 ZIP 或外层文件夹。

## Render 设置

现有网站保持以下设置即可：

- 类型：Static Site
- Root Directory：留空
- Build Command：`test -s public/index.html`
- Publish Directory：`public`

提交后等待 Render 部署完成，再打开原网站。无需安装依赖或运行 npm 命令。

## 使用

导入 `.xlsx` 商品表，搜索或粘贴货号，导出包含该货号全部颜色与完整行的选款 Excel。

本版已包含自动保存：文件旁显示“已保存在此浏览器”后，下次用同一浏览器打开同一网址会自动恢复，无需重新选文件。手机、电脑及不同浏览器分别保存，不互相同步。移除文件或清空本地资料会删除浏览器副本；清理网站数据后需要重新导入。原始 Excel 请保留，资料更新后移除旧表再添加新版。
