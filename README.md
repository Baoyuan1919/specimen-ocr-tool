# 标本照片 → Excel 自动填表工具 v2.0

批量提取标本照片上的文字，按字段名自动填入 Excel。

## ✨ 功能

| 功能 | 说明 |
|------|------|
| **批量 OCR 识别** | 一次性处理文件夹内所有图片（jpg/png/bmp/tiff/webp） |
| **中文精准识别** | 基于 RapidOCR（ONNX Runtime），中英文混排无压力 |
| **字段完全自定义** | 增删改字段，与照片上的文字名称对应即可 |
| **从 Excel 同步字段** | 可直接读取现有 Excel 表头作为识别字段 |
| **文件夹自动监听** | 开启后自动检测新增图片，填入 Excel |
| **系统托盘后台** | 关闭窗口不退出，缩到托盘区持续监听 |
| **轻量无框架依赖** | 仅 3 个核心库，无需 PaddlePaddle 等大框架 |

## 识别的默认字段

采集号、采集时间、采集人、采集地点、经度、纬度、海拔、习性、生态环境、高度

> ⚙️ 这些字段可以在软件中自由增删改，也可以从 Excel 表头一键同步。

## 使用

1. 把标本照片按顺序编号放到一个文件夹
2. 打开软件 → 选图片文件夹 → 选/新建 Excel 文件
3. 可点击「管理字段」调整要识别的字段
4. 点 **「手动识别并填表」** 立即处理，或点 **「开始监听」** 自动处理新图片
5. 打开 Excel 查看结果

## 后台运行

- 点击「最小化到托盘」→ 软件缩到系统托盘区，右键可恢复/退出
- 开启「监听」后即使最小化，也会自动处理新放入的图片

## 获取 .exe

### 方法一：从 GitHub Actions 下载

1. 打开 [Actions 页面](https://github.com/Baoyuan1919/specimen-ocr-tool/actions)
2. 点击最新的绿色勾 ✅ 的 workflow
3. 往下翻到 **Artifacts** → 下载 **标本OCR填表工具** 或 **安装包**

### 方法二：自行打包

```bash
pip install onnxruntime==1.17.1 rapidocr-onnxruntime==1.3.22 openpyxl Pillow pystray pyinstaller
pyinstaller --onefile --noconsole --name "标本OCR填表工具" main.py
```
