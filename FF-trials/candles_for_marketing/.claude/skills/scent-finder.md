---
name: scent-finder
description: 从Excel商品表中自动检索香薰蜡烛的香型（香族+香调），支持二级细粒度分类。
---

# 香型检索 Skill

## 触发条件

用户提供一份包含「商品」列的 Excel，需要自动获取每个商品的香型信息时调用。

## 执行流程

### 第一步：读取 Excel

```python
import pandas as pd
df = pd.read_excel(INPUT_FILE, header=None)
# 结构：row 0=合并表头，row 1=列名，row 2+=数据
# 商品列通常在第9列（索引8）
```

### 第二步：排除无香商品

检查「商品」列文本是否命中以下模式（来自 `brand_exclusions.json` → `unscented_detection`）：
- `unscented`, `fragrance[-]free`, `scent[-]free`, `non[-]scented`
- `flameless`, `battery[-]operated`, `led candle`, `fake candle`

命中 → 直接标记为「无香型」或「无香型（电子）」，**不搜索**。

### 第三步：在线搜索（所有有香商品）

按优先级搜索：

1. **品牌官网**：`site:{brand}.com {product_short_name}`
2. **Amazon 页面**：`site:amazon.com "{product_short_name}"`
3. **香水评测站**：`site:fragrantica.com OR site:nosetime.com {product_short_name}`
4. **通用 Bing**：`"{product_short_name}" candle scent fragrance notes`

提取搜索结果摘要文本（前5条，每条截取200字符），合并为搜索文本。

### 第四步：文本清洗

对「商品名 + 搜索结果文本」执行清洗（来自 `brand_exclusions.json` → `clean_patterns`）：

1. **品牌名** → 移除（如 `WoodWick` ≠ 木质, `Salt & Stone` ≠ 盐）
2. **Wick 材质** → 移除（`wood wick`, `cotton wick` ≠ 香调）
3. **蜡基** → 移除（`soy wax`, `beeswax` ≠ 香调）
4. **功能性词** → 移除（`smokeless` → 不是 smoke 香！）
5. **场景词** → 移除（`gifts for`, `birthday`, `wedding`）
6. **规格词** → 移除（`pack of`, `count`, `oz`）

### 第五步：关键词提取与香族映射

使用 `fragrance_dict.json` 中的关键词表进行匹配：

**匹配规则**：
- 长词优先匹配（如 `fireside` 先于 `fir`，`sea salt` 先于 `salt`）
- 短词使用 `\b` 词边界：`fir`, `tea`, `salt`, `rose`, `pine`, `cake`, `spring`

**输出**：
- **一级香族** = 命中关键词所属香族的并集（一个商品可多个）
- **二级香调** = 所有命中的关键词

**10大香族**（来自 `fragrance_dict.json`）：
| 香族 | 市场高频词示例 |
|------|---------------|
| 柑橘 Citrus | citrus(3), orange(4), lemon(5), grapefruit(1), bergamot(3) |
| 花香 Floral | lavender(22), jasmine(5), rose(2), lily(1), peony, gardenia |
| 果香 Fruity | apple(2), pear(3), coconut(2), fruit(2), champagne(3) |
| 木质 Woody | sandalwood(11), cedar(5), fir(8), mahogany(3), vetiver(1) |
| 美食 Gourmand | vanilla(13), cake(23), cupcake(5), caramel(1), pumpkin(2) |
| 草本 Herbal | eucalyptus(4), sage(5), mint(5), lemongrass(1), tea |
| 海洋 Aquatic | sea(3), ocean, marine, beach, sea salt(1) |
| 清新 Fresh | fresh(13), clean(13), rain(4), bamboo(1), linen |
| 东方 Oriental | amber(7), cinnamon(6), oud(3), clove(2), musk, spice(3) |

> 括号内为200条真实市场数据中的出现频次。

### 第六步：写入 Excel

在原 Excel 末尾追加列：
- `无香判定` — 是/否
- `一级香族` — 命中的香族（逗号分隔）
- `二级香调` — 命中的具体香调（逗号分隔）
- `来源` — 关键词匹配 / 在线搜索
- `搜索URL` — 在线搜索时记录来源

## 配置文件

- `fragrance_dict.json` — 关键词→香族映射表，可手动编辑扩充
- `brand_exclusions.json` — 品牌名、材质词、功能性词排除表
- `find_scent_type_v2.py` — 完整 Python 实现脚本

## 重要提醒

1. **品牌名只在文本清洗阶段排除**，不影响是否搜索的决策
2. **smoke 是最危险的误判源**：出现 22 次全部来自 `smokeless`，必须用清洗规则移除
3. **允许多标签**：一个商品可同时命中多个香族，不用纠结「混合香型」
4. **搜索限速 2-3 秒/次**，200 商品约需 5-10 分钟
5. **缓存搜索结果**到 `scent_cache.json`，中断后可续跑
