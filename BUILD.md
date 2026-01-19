# 打包说明

## 使用 GitHub Actions 自动打包（推荐）

### 步骤 1：推送代码到 GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/你的用户名/你的仓库名.git
git push -u origin main
```

### 步骤 2：触发自动打包

有两种方式：

#### 方式 A：创建 Release Tag（推荐）

```bash
git tag v1.0.0
git push origin v1.0.0
```

#### 方式 B：手动触发

1. 进入 GitHub 仓库页面
2. 点击 "Actions" 标签
3. 选择 "Build macOS App" workflow
4. 点击 "Run workflow" → "Run workflow"

### 步骤 3：下载打包好的 App

打包完成后（约 5-10 分钟）：

1. 进入 "Actions" 页面
2. 点击最新的 workflow run
3. 在 "Artifacts" 部分下载 `知网Excel校验工具-macOS`
4. 解压后得到 `.app` 文件

## 本地打包（macOS 环境）

如果你在 macOS 上，也可以本地打包：

```bash
# 1. 安装依赖
pip install -r requirements.txt
pip install pyinstaller

# 2. 打包
pyinstaller --noconfirm \
  --windowed \
  --name "知网Excel校验工具" \
  --hidden-import=tkinter \
  --hidden-import=tkinterdnd2 \
  --hidden-import=pandas \
  --hidden-import=openpyxl \
  --hidden-import=xlrd \
  --hidden-import=selenium \
  --hidden-import=webdriver_manager \
  --collect-all=tkinterdnd2 \
  check_cnki_excel.py

# 3. 移除隔离属性（解决 Gatekeeper 问题）
xattr -dr com.apple.quarantine "dist/知网Excel校验工具.app"
```

打包结果在 `dist/知网Excel校验工具.app`

## 注意事项

1. **代码不需要修改**：GitHub Actions 会自动在 macOS 环境打包
2. **Chrome 浏览器**：用户需要自行安装 Chrome，脚本会自动管理 ChromeDriver
3. **首次运行**：可能需要联网下载 ChromeDriver（webdriver-manager 会自动处理）
4. **Gatekeeper**：如果遇到"无法打开"提示，右键选择"打开"即可

## 文件大小

打包后的 `.app` 文件约 100-200MB（包含所有 Python 依赖和库）
