# 复利、方程与一段说明

这是一份用于演示转换效果的 Markdown。从 ChatGPT、Claude、Cursor 等工具复制后，可直接导入 md2word。

## 正文排版

转换后的正文默认：**首行缩进 2 字符**、**两端对齐**、**1.5 倍行距**。

> 引用块会保留。分节横线（`---`）在转换时会被去掉，避免 Word 里多出一条分隔线。

每年把利息并入本金，本金 $P$ 按年利率 $r$ 增长 $t$ 年后变为：

$$
A = P(1+r)^{t}
$$

## 列表与表格

- 无序列表
- **加粗** 与 *斜体*
- 行内代码：`print("hello")`

| 项目 | 说明 |
| --- | --- |
| Markdown | 源格式 |
| Word | 目标格式，含可编辑公式 |
| PDF | 需安装 Microsoft Word |

## 公式

行内公式：$E = mc^2$。AI 有时会写成不带 `$` 的括号形式，例如 (\mu) 或 (\hat{x})，md2word 会尽量识别并包成 Word 公式。

一元二次方程的求根公式：

$$
x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}
$$

也支持这种常见的 AI 输出（方括号包住的 LaTeX）：

[

\Delta = b^2-4ac

]

## 代码块

```python
def hello(name: str) -> str:
    return f"Hello, {name}"
```
