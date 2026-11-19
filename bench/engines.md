# Local LLM engines on one deck-mode request

| config | engine | load s | cold s | warm s | prompt tok | gen tok | warm gen tok/s | quality | grounded bullets | numbers absent | peak server GB | lowest available GB |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| coli_vk_ram9 | colibri | 6.5 | 290.37 | 271.25 | 1851 | 452 | 1.67 | n/a | n/a | n/a | 9.36 | 8.91 |
| coli_vk_nochain_ram9 | colibri | 2.8 | 241.6 | 238.6 | 1851 | 406 | 1.7 | n/a | n/a | n/a | 8.11 | 10.3 |
| llama_qwen35_9b | llama.cpp | 22.7 | 58.98 | 47.36 | 1851 | 460 | 9.7 | 1.0 | 1.0 | none | 7.69 | 10.14 |
| coli_cpu_ram9 | colibri | 2.0 | 96.2 | 79.87 | 1851 | 448 | 5.61 | 1.0 | 1.0 | none | 4.33 | 13.02 |
