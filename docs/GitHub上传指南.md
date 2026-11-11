# 上传到 GitHub 指南

本指南把「本地已经整理好的项目」发布为 GitHub 仓库，并把打包好的 exe 通过 **Releases** 分发。
全程在 **项目根目录** `E:\work\思政培优项目\stat_analysis` 下执行，命令均在 **PowerShell** 中运行。

---

## 一、已经整理好的内容

上传前，仓库里应当只出现源码、文档与配置：

| 会上传 | 说明 |
| --- | --- |
| `report_generator.py`、`trend_analyzer.py` | 两个主程序 |
| `就业质量报告生成器.spec`、`就业变化趋势分析器.spec` | 打包配置（已改为相对路径，换电脑可直接打包） |
| `build.ps1` | 一键打包脚本 |
| `requirements.txt`、`环境配置清单.txt` | 依赖清单 |
| `README.md`、`LICENSE`、`.gitignore`、`.gitattributes` | 仓库基础文件 |
| `使用说明.docx` | 已完善的使用说明书 |
| `docs/GitHub上传指南.md` | 本文件 |
| `examples/数据模板.xlsx` | 仅表头 + 3 行示例的数据模板 |

| 不会上传（已被 `.gitignore` 排除） | 原因 |
| --- | --- |
| `dist/*.exe`（214 MB / 245 MB / 258 MB） | 超过 GitHub 单文件 100 MB 硬限制，改走 Releases |
| `build/` | PyInstaller 中间产物，可重新生成 |
| `log/` | 运行日志，含本机绝对路径 |
| `legacy/` | 本地历史归档（原 Word 说明原稿、旧版 `app.py` 归档） |
| `*.xlsx`（除 `examples/`） | 毕业生数据含个人信息，禁止上传 |

---

## 二、准备 Git 环境

### 1. 安装 Git

下载并安装 [Git for Windows](https://git-scm.com/download/win)，安装选项一路默认即可。装完新开一个 PowerShell 窗口验证：

```powershell
git --version
```

### 2. 首次配置身份与中文显示

```powershell
git config --global user.name "你的名字或昵称"
git config --global user.email "你的GitHub邮箱"
git config --global core.quotepath false      # 中文文件名不再显示成 \346\226\207
git config --global init.defaultBranch main
```

> `user.email` 与 GitHub 账号邮箱一致时，提交记录才会关联到你的头像下。

---

## 三、在 GitHub 创建空仓库

1. 登录 GitHub，点右上角 **+** → **New repository**。
2. **Repository name** 建议用英文，例如 `stat_analysis`（仓库名用英文更稳妥，描述可以写中文）。
3. **Description** 可填：`毕业生就业数据统计分析工具：一键生成带图表的 Word 就业质量报告与多年趋势分析`。
4. 可见性按需选择 **Public** 或 **Private**。
5. **不要**勾选 `Add a README file`、`.gitignore`、`license`——本地已经有这些文件，勾选会造成首次推送冲突。
6. 点 **Create repository**。

---

## 四、本地初始化并提交

```powershell
cd "E:\work\思政培优项目\stat_analysis"

git init
git add .

# 关键一步：确认暂存区里没有 exe / xlsx / log
git status
```

`git status` 里应当只有源码、文档、配置和 `examples/数据模板.xlsx`。
若看到 `dist/`、`build/`、`log/` 或其它 `*.xlsx`，说明 `.gitignore` 没生效，先执行 `git rm -r --cached .` 再 `git add .`。

```powershell
git commit -m "feat: 就业数据统计分析工具首次提交"
```

---

## 五、关联远程仓库并推送

在 GitHub 仓库页面复制 **HTTPS 地址**（形如 `https://github.com/你的用户名/stat_analysis.git`），然后：

```powershell
git remote add origin https://github.com/你的用户名/stat_analysis.git
git branch -M main
git push -u origin main
```

首次推送会弹出登录窗口（Git Credential Manager）：

- 选 **Sign in with your browser**，在浏览器里授权即可；
- 若要求输入密码，**不能用账号密码**，需要用 **Personal Access Token**：
  GitHub → 头像 → Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token，
  勾选 `repo` 权限，生成后把 token 当作密码粘贴（只显示一次，请先保存）。

推送成功后刷新页面，源码与文档即出现在仓库中。

---

## 六、把 exe 发布到 Releases（不要提交进仓库）

**为什么不能直接提交**：两个 exe 分别是 245 MB 与 258 MB，GitHub 对单个文件有 **100 MB 硬限制**，`git push` 会被直接拒绝；且 exe 是构建产物，不适合进版本库。

Releases 的单文件上限是 **2 GB**，适合放成品：

1. 打开仓库首页，右侧栏点 **Releases** → **Draft a new release**。
2. **Choose a tag** 填 `v1.0.0`，点 **Create new tag on publish**。
3. **Release title** 填 `v1.0.0 首个公开版本`。
4. 说明区可写：
   ```markdown
   - 就业质量报告生成器.exe：单年度就业质量报告
   - 就业变化趋势分析器.exe：多年就业趋势分析
   详细用法见 使用说明.docx
   ```
5. 把 `dist\就业质量报告生成器.exe` 和 `dist\就业变化趋势分析器.exe` 拖到附件区等待上传完成（约 500 MB，视网速可能需要几分钟）。
6. 点 **Publish release**。

> 也可以把两个 exe 打包成一个 zip 上传（本项目当前就是这么做的），但务必在 Release 说明里写清"解压后得到哪两个文件"，并同步修改 README 的下载说明，避免读者只看到 `default.zip` 不知道是什么。

之后 README 里“在 Releases 页面下载”的链接会自动指向你自己的仓库，无需修改。

---

## 七、后续更新流程

```powershell
git add .
git status                       # 习惯性确认
git commit -m "docs: 补充数据口径说明"
git push
```

需要发布新版本 exe 时：先 `.\build.ps1` 重新打包，再在 Releases 里 **Draft a new release**，tag 用 `v1.0.1`。也可以先在本地打 tag 再推送：

```powershell
git tag v1.0.1
git push origin v1.0.1
```

---

## 八、常见问题

| 现象 | 原因与处理 |
| --- | --- |
| `remote: error: File dist/xxx.exe is 245.00 MB; this exceeds GitHub's file size limit of 100.00 MB` | exe 被加进了提交。执行 `git rm --cached "dist/就业质量报告生成器.exe"`，确认 `.gitignore` 含 `dist/`，再重新 `git add .` 提交 |
| 中文文件名显示成 `\346\226\207` | `git config --global core.quotepath false` |
| 推送一直要求输入用户名密码 / 认证失败 | 用 Personal Access Token 当密码，或改用 SSH（见下） |
| `Failed to connect to github.com port 443` / 超时 | 网络问题。挂了代理则设置：`git config --global http.proxy http://127.0.0.1:7890`；用完后 `git config --global --unset http.proxy` 取消 |
| 首次推送报 `fetch first` / `rejected` | 建仓库时勾选了 README。执行 `git pull origin main --allow-unrelated-histories` 合并后再推送 |
| 误提交了 xlsx 或大文件且**尚未推送** | `git rm --cached 文件路径`，然后 `git commit --amend` |
| 误提交且**已经推送** | 最省事的做法：GitHub 上删除仓库重新创建并按本指南重来；否则需要用 `git filter-repo` 重写历史（会影响协作者） |
| `.gitignore` 改了但不生效 | 已被跟踪的文件不受 `.gitignore` 影响：`git rm -r --cached .` 后重新 `git add .` |
| 仓库体积过大 | 检查 `git count-objects -vH`，多半是 exe 或数据被提交过，按上面两条清理 |

### 可选：改用 SSH 免密推送

```powershell
ssh-keygen -t ed25519 -C "你的GitHub邮箱"      # 一路回车
cat ~/.ssh/id_ed25519.pub                       # 复制输出内容
```

把输出粘贴到 GitHub → Settings → **SSH and GPG keys** → New SSH key；然后：

```powershell
git remote set-url origin git@github.com:你的用户名/stat_analysis.git
```

---

## 九、上传前的检查清单

- [ ] `git status` 中没有 `*.xlsx`（`examples/数据模板.xlsx` 除外）、`*.exe`、`log/`
- [x] `README.md` 里 clone 示例已填入 `ShaoXian-Li`（若日后改仓库名需同步修改）
- [ ] 按需修改 `LICENSE` 的版权行
- [ ] 仓库 **About**（右侧齿轮）里填写描述，并加上 Topics：`statistics`、`python`、`education`
- [ ] 发布 Release 并上传两个 exe
- [ ] 用浏览器打开仓库首页，确认 README 的图片徽章、表格和链接正常显示
