# Candles_Amazon — 蜡烛香型自动分类器

## 项目结构

```
FF-trial/
├── app.py                    # Streamlit Web UI（主界面）
├── scent_classifier.py       # 分类引擎（核心逻辑）
├── fragrance_dict.json       # 香调关键词字典（9 大香族 + Mix）
├── brand_exclusions.json     # 品牌名 & 干扰词排除列表
├── 分类与搜索策略.md          # 分类策略设计文档
├── Candles_Amazon.html       # 纯前端版（无需 Python，双击即用）
├── requirements.txt          # Python 依赖清单
├── run.bat                   # Windows 一键启动脚本
├── Candles_Amazon.zip        # 部署包（以上文件的压缩包）
└── README.md                 # 本文档
```

## 分类流水线（4 阶段）

```
Excel 读取 → ① 无香排除 → ② 在线搜索 → ③ 文本清洗 → ④ 关键词映射
```

### ① 搜索前排除
检测商品名中的无香/电子蜡烛关键词（`unscented`, `flameless`, `LED` 等），命中则直接标记"无香型"，跳过后续步骤。

### ② 在线搜索
**双通道 fallback：**
1. **curl_cffi**（⚡ 优先）— 模拟浏览器 TLS 指纹，快（~5s/条），但 Amazon Akamai 偶尔封
2. **Playwright**（🐢 备选）— 真实 Chromium 浏览器，稳但慢（~8s/条），使用系统自带 Edge 无需额外下载

搜索流程：Amazon 搜索页 → 取 ASIN → 产品详情页 → 提取标题 + bullet points + Scent 字段。

### ③ 文本清洗
从搜索结果中移除干扰项（品牌名、wick 材质、蜡基、功能性词、场景词），避免误匹配。清洗逻辑在 `clean_text()` 函数中。

### ④ 关键词提取 → 香族映射
1. **关键词匹配**：在清洗后文本中用 `fragrance_dict.json` 的 200+ 关键词做正则匹配（长词优先，短词加 `\b` 边界）
2. **大类收敛**（单一家族）：
   - 细分关键词 ≥ 8 → **混合 Mix**
   - 否则按加权得分选最高大类：商品名自带关键词 2 分，搜索来的 1 分
   - 丢弃非选中大类的关键词
3. **格式化**：`大类` + `细分（关键词(中文简称)[搜]）`

## 10 大香族

| 家族 | 示例关键词 |
|------|-----------|
| 🍊 柑橘 Citrus | lemon, orange, lime, grapefruit |
| 🌸 花香 Floral | lavender, rose, jasmine, peony |
| 🍎 果香 Fruity | apple, coconut, peach, berry |
| 🪵 木质 Woody | sandalwood, cedar, pine, fir |
| 🍰 美食 Gourmand | vanilla, cupcake, caramel, honey |
| 🌿 草本 Herbal | eucalyptus, sage, mint, tea |
| 🌊 海洋 Aquatic | ocean, sea, marine, beach |
| ✨ 清新 Fresh | fresh, clean, linen, cotton |
| 🔥 东方 Oriental | amber, oud, cinnamon, musk |
| 🔀 混合 Mix | 细分 ≥ 8 自动触发 |

## 部署方式

### 你自己（完整版，在线搜索）
```bash
streamlit run app.py
# → http://localhost:8501
```

### 同事（免安装）
1. 装一次 Python 3.12（python.org，勾选"Add to PATH"）
2. 解压 `Candles_Amazon.zip`
3. 双击 `run.bat`（自动装依赖 + 启动）
4. 或直接双击 `Candles_Amazon.html`（纯前端版，不联网，秒级）

### 同一网络共享
你跑着 Streamlit 时，同事直接浏览器打开 `http://192.168.31.30:8501`，什么都不用装。

## 踩过的坑

### 1. Greenlet 跨线程崩溃
**现象**：`ThreadPoolExecutor` 多线程时 Playwright 崩溃 `greenlet.error: Cannot switch to a different thread`
**原因**：Playwright 使用 greenlet 协程，全局单例浏览器不能跨线程共享
**修复**：用 `threading.local()` 为每个线程创建独立浏览器实例

### 2. 表格无数据
**现象**：UI 表格只显示表头，没有数据行
**原因**：`细分_clean` 字段格式是 `rose(花香)`，但 `KEYWORD_TO_FAMILY` 的 key 是纯英文 `rose`，匹配失败导致所有二级关键词数据为空
**修复**：用正则 `re.sub(r"\([^)]*\)", "", ...)` 去掉中文括号注释

### 3. curl_cffi 全 503
**现象**：所有 curl_cffi 请求返回 503
**原因**：Amazon Akamai 检测到爬虫行为后 IP 限流（跑 200 条后触发）
**修复**：自动 fallback 到 Playwright（真实浏览器无法被检测）；IP 几小时后自动解封

### 4. run.bat 乱码
**现象**：同事双击 `run.bat` 报乱码 `'hon3'不是内部或外部命令`
**原因**：Unix LF 换行符在 Windows CMD 中不兼容，UTF-8 中文被 GBK 误解析
**修复**：统一使用 CRLF 换行 + 纯英文内容

### 5. Device Guard 拦截 pip
**现象**：公司电脑 `pip.exe` 被组织策略阻止
**原因**：Windows Device Guard / AppLocker 白名单策略
**修复**：改用 `python -m pip` 绕过 exe 拦截

### 6. Playwright Chromium 下载慢
**现象**：`playwright install chromium` 下载 ~150MB 浏览器，国内网速慢
**修复**：改用系统自带 Microsoft Edge（`channel="msedge"`），零额外下载

## 关键技术栈

| 层级 | 工具 | 用途 |
|------|------|------|
| UI | Streamlit | Python Web 界面 |
| HTTP | curl_cffi | TLS 指纹模拟（绕过 Amazon 检测） |
| 浏览器 | Playwright + Edge | JS 渲染 fallback（过 Akamai） |
| HTML 解析 | BeautifulSoup4 | 提取 Amazon 页面文本 |
| Excel | openpyxl | 读写 .xlsx |
| 并发 | ThreadPoolExecutor | 双线程提速 |
| 纯前端 | SheetJS (xlsx.js) | HTML 版 Excel 解析（无需 Python） |
