# LLM benchmark (pdf2pptx bench)

Machine: AMD64 Family 25 Model 117 Stepping 2, AuthenticAMD, 30.7 GB RAM. llama.cpp b12503 (Vulkan). Run 2026-11-18 15:51, 715 s in total.

Tasks: 10 judgment items (4 role pages, 2 reading-order pages, 3 tables, 1 region), 5 chart regions, 2 grounded text outlines. Accuracy 0-1 against the corpus truth; document models answer tables / charts only, text models text only.

RAM reserve during the run: 8.0 GB (default 8). 'fits 8 GB' = the model would start under the default reserve with the RAM that was free.

| rank | model | kind | load s | s / image req | prompt tok/s | gen tok/s | peak RAM GB | RAM need GB | fits 8 GB | --n-cpu-moe | judgment | table | chart | text | overall | total s |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Qwen3.5-9B | vlm | 5.4 | 24.2 | 56.0 | 10.6 | 7.77 | 7.67 | yes | 0 | 1.00 | 0.91 | 1.00 | 1.00 | 0.98 | 416 |
| 2 | Qwen3.5-4B | vlm | 2.8 | 13.8 | 90.4 | 18.1 | 5.04 | 4.95 | yes | 0 | 0.95 | 0.91 | 0.98 | 1.00 | 0.96 | 248 |
| 3 | PaddleOCR-VL-1.6 | docvlm | 1.6 | 5.2 | 80.2 | 70.1 | 2.25 | 2.52 | yes | 0 | - | 1.00 | 0.54 | - | 0.77 | 45 |
| 4 | granite-docling-258M | docvlm | 1.4 | 9.8 | 71.2 | 122.9 | 0.99 | 1.33 | yes | 0 | - | 1.00 | 0.32 | - | 0.66 | 88 |
| - | Qwen3.6-35B-A3B | vlm | not run: Qwen3.6-35B-A3B: needs 14.1 GB, budget 10.5 GB (available - 8 GB); 3.6 GB short of RAM (needs 14.1 GB + 8 GB reserve): c |||||||||||||
| - | gpt-oss-20b | text | not run: gpt-oss-20b: needs 12.4 GB, budget 10.5 GB (available - 8 GB); 1.9 GB short of RAM (needs 12.4 GB + 8 GB reserve): close |||||||||||||

## Recommended stack (written to pdf2pptx.toml)

- `llm_judge` = Qwen3.5-9B
- `llm_chart` = Qwen3.5-9B
- `llm_doc` = PaddleOCR-VL-1.6
- `llm_text` = Qwen3.5-9B

## Per item

| model | item | score | s | detail |
|---|---|---|---|---|
| Qwen3.5-9B | r01_report_two_column:p0 | 1.00 | 23.3 |  |
| Qwen3.5-9B | r02_report_tables_it:p0 | 1.00 | 41.9 |  |
| Qwen3.5-9B | r03_report_figure:p0 | 1.00 | 29.8 |  |
| Qwen3.5-9B | d17_footer_logo_number:p0 | 1.00 | 13.7 |  |
| Qwen3.5-9B | r01_report_two_column:p0 | 1.00 | 14.9 |  |
| Qwen3.5-9B | d04_three_columns:p0 | 1.00 | 13.0 |  |
| Qwen3.5-9B | d07_table_simple:t0 | 1.00 | 24.7 |  |
| Qwen3.5-9B | d08_table_merged:t0 | 0.73 | 25.4 |  |
| Qwen3.5-9B | r02_report_tables_it:t0 | 1.00 | 16.8 |  |
| Qwen3.5-9B | d10_arrows_chevrons:p0 | 1.00 | 10.1 | "process" |
| Qwen3.5-9B | d14_chart_bar:p0 | 1.00 | 27.9 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-9B | d14_chart_bar:p1 | 1.00 | 18.1 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-9B | d15_chart_line:p0 | 1.00 | 29.4 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-9B | d16_chart_pie:p0 | 1.00 | 25.3 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-9B | d27_italian_report_4x3:p1 | 1.00 | 24.2 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-9B | r01_report_two_column | 1.00 | 43.6 |  |
| Qwen3.5-9B | d27_italian_report_4x3 | 1.00 | 28.0 |  |
| Qwen3.5-4B | r01_report_two_column:p0 | 1.00 | 13.5 |  |
| Qwen3.5-4B | r02_report_tables_it:p0 | 0.88 | 23.8 |  |
| Qwen3.5-4B | r03_report_figure:p0 | 0.80 | 17.6 |  |
| Qwen3.5-4B | d17_footer_logo_number:p0 | 1.00 | 10.2 |  |
| Qwen3.5-4B | r01_report_two_column:p0 | 1.00 | 8.8 |  |
| Qwen3.5-4B | d04_three_columns:p0 | 1.00 | 7.8 |  |
| Qwen3.5-4B | d07_table_simple:t0 | 1.00 | 14.7 |  |
| Qwen3.5-4B | d08_table_merged:t0 | 0.73 | 13.8 |  |
| Qwen3.5-4B | r02_report_tables_it:t0 | 1.00 | 10.3 |  |
| Qwen3.5-4B | d10_arrows_chevrons:p0 | 1.00 | 8.9 | "process" |
| Qwen3.5-4B | d14_chart_bar:p0 | 1.00 | 18.3 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-4B | d14_chart_bar:p1 | 1.00 | 10.8 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-4B | d15_chart_line:p0 | 0.92 | 17.5 | {"type": 1.0, "categories": 1.0, "values": 0.75, "score": 0.9166666666666666} |
| Qwen3.5-4B | d16_chart_pie:p0 | 1.00 | 18.1 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-4B | d27_italian_report_4x3:p1 | 1.00 | 14.3 | {"type": 1.0, "categories": 1.0, "values": 1.0, "score": 1.0} |
| Qwen3.5-4B | r01_report_two_column | 1.00 | 25.5 |  |
| Qwen3.5-4B | d27_italian_report_4x3 | 1.00 | 11.3 |  |
| PaddleOCR-VL-1.6 | d07_table_simple:t0 | 1.00 | 4.8 |  |
| PaddleOCR-VL-1.6 | d08_table_merged:t0 | 1.00 | 4.7 |  |
| PaddleOCR-VL-1.6 | r02_report_tables_it:t0 | 1.00 | 3.0 |  |
| PaddleOCR-VL-1.6 | d14_chart_bar:p0 | 0.67 | 5.5 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
| PaddleOCR-VL-1.6 | d14_chart_bar:p1 | 0.67 | 4.9 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
| PaddleOCR-VL-1.6 | d15_chart_line:p0 | 0.04 | 5.8 | {"type": 0.0, "categories": 0.0, "values": 0.125, "score": 0.041666666666666664} |
| PaddleOCR-VL-1.6 | d16_chart_pie:p0 | 0.67 | 7.4 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
| PaddleOCR-VL-1.6 | d27_italian_report_4x3:p1 | 0.67 | 6.9 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
| granite-docling-258M | d07_table_simple:t0 | 1.00 | 9.0 |  |
| granite-docling-258M | d08_table_merged:t0 | 1.00 | 9.1 |  |
| granite-docling-258M | r02_report_tables_it:t0 | 1.00 | 12.8 |  |
| granite-docling-258M | d14_chart_bar:p0 | 0.67 | 9.4 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
| granite-docling-258M | d14_chart_bar:p1 | 0.17 | 9.1 | {"type": 0.0, "categories": 0.0, "values": 0.5, "score": 0.16666666666666666} |
| granite-docling-258M | d15_chart_line:p0 | 0.08 | 10.2 | {"type": 0.0, "categories": 0.0, "values": 0.25, "score": 0.08333333333333333} |
| granite-docling-258M | d16_chart_pie:p0 | 0.00 | 13.6 | {"type": 0.0, "categories": 0.0, "values": 0.0, "score": 0.0} |
| granite-docling-258M | d27_italian_report_4x3:p1 | 0.67 | 13.2 | {"type": 0.0, "categories": 1.0, "values": 1.0, "score": 0.6666666666666666} |
