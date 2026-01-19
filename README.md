# 知网 Excel 校验工具

自动校验知网文章标题与发布时间的匹配工具。

## 功能特点

- ✅ 支持拖拽 Excel 文件（.xlsx、.xls、WPS 格式）
- ✅ 自动选择日期并检索
- ✅ 分页查找标题（支持多页结果）
- ✅ 详细调试信息输出
- ✅ 控制台和 GUI 双重报错显示

## 支持的 Excel 格式

- **新版 Excel**: `.xlsx`、`.xlsm`（使用 openpyxl 引擎）
- **旧版 Excel**: `.xls`（使用 xlrd 引擎）
- **WPS**: `.xlsx`、`.xls`（WPS 保存的文件）

## 使用方法

### 方式一：直接运行 Python 脚本

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行脚本：
```bash
python check_cnki_excel.py
```

3. 将 Excel 文件拖拽到窗口中

### 方式二：使用打包好的 macOS App

#### 通过 GitHub Actions 自动打包

1. **推送代码到 GitHub**：
```bash
git add .
git commit -m "Initial commit"
git push origin main
```

2. **创建 Release Tag 触发自动打包**：
```bash
git tag v1.0.0
git push origin v1.0.0
```

或者直接在 GitHub 网页上：
- 进入 Releases → Create a new release
- 填写版本号（如 `v1.0.0`）
- 点击 "Publish release"

3. **下载打包好的 App**：
   - 自动打包完成后，在 Actions 页面下载 `知网Excel校验工具-macOS` artifact
   - 或者在 Releases 页面下载

#### 手动触发打包

1. 进入 GitHub 仓库的 Actions 页面
2. 选择 "Build macOS App" workflow
3. 点击 "Run workflow" → "Run workflow"

## Excel 文件格式要求

Excel 文件必须包含以下两列：
- **发布时间**：日期格式（如 2019-12-31）
- **标题**：文章标题文本

## 系统要求

- **macOS**: 10.14 或更高版本
- **Python**: 3.8+（如果直接运行脚本）
- **Chrome 浏览器**：需要用户自行安装（脚本会自动管理 ChromeDriver）

## 常见问题

### macOS "无法打开，因为来自身份不明的开发者"

1. 右键点击 `.app` 文件
2. 选择"打开"
3. 在弹出的对话框中再次点击"打开"

或者运行：
```bash
xattr -dr com.apple.quarantine "知网Excel校验工具.app"
```

### 读取旧版 .xls 文件失败

确保安装了 `xlrd<2.0`：
```bash
pip install "xlrd<2.0"
```

### ChromeDriver 相关问题

脚本使用 `webdriver-manager` 自动管理 ChromeDriver，首次运行需要联网下载。

## 开发说明

### 本地打包（macOS）

```bash
pip install pyinstaller
pyinstaller --noconfirm --windowed --name "知网Excel校验工具" check_cnki_excel.py
```

打包结果在 `dist/知网Excel校验工具.app`

## License

MIT
