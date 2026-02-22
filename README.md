# AI产业链投研数据库系统（Equity Research Infra）

纯 Python 标准库实现的数据 pipeline（无需第三方包）。

## 功能
- 自动抓取公司历史财报（优先 Yahoo Finance；失败自动回退离线样本）
- 指标：收入、毛利率、净利润
- AI 相关收入识别（AI 暴露度参数）
- 产业链标签：光芯片 / InP材料 / CCL / T-Glass / AI PCB / 光模块
- 利润弹性情景：+50% / +100% / +150% / +200%
- 输出 Excel（xlsx，多 sheet）
- 输出 EPS Sensitivity 图（SVG）

## 覆盖公司
AXTI、Coherent、Sumitomo Electric、Ibiden、Shinko、Nittobo。

## 运行
```bash
python equity_research_pipeline.py --offline
```

或尝试在线抓取：
```bash
python equity_research_pipeline.py
```

## 输出
- `output.xlsx`
- `output/eps_sensitivity.svg`

## 关键模型
- `AIRevenue = Revenue * ai_exposure_ratio`
- `IncrementalProfit = AIRevenue * PriceIncrease * ai_incremental_margin`
- `NetIncome_Scenario = NetIncome_Base + IncrementalProfit`
- `NetIncome_2026E = NetIncome_2026_Base + IncrementalProfit`
- `EPS = NetIncome_Scenario / SharesOutstanding`

> 参数位于 `equity_research_pipeline.py` 的 `COMPANIES` 配置，可按研究假设持续更新。
