# Hardening (M10)

| case | expected | exit | seconds | peak RAM GB (converter / all POWERPNT.EXE) | traceback | outcome |
|---|---|---|---|---|---|---|
| corrupted | exit code 3, message says damaged | 3 | 1.8 | 0.05 / 0.21 | False | pdf2pptx: corrupted.pdf: the PDF is damaged and cannot be read (Failed to load document (PDFium: Data format error).); try re-saving it with a PDF viewer |
| damaged_page | converts; the broken page degrades, the others convert | 0 | 27.8 | 1.24 / 0.22 | False | {'pages': 3, 'slides': 3, 'seconds': 24.87, 'seconds_per_page': 8.29, 'fallback_pages': [], 'errors': [], 'tune': 'ran'} |
| encrypted | exit code 3, message names the password | 3 | 1.8 | 0.05 / 0.21 | False | pdf2pptx: encrypted.pdf: the PDF is encrypted (password protected); open it, remove the password (or print to PDF) and try again |
| images_50mb | converts; pictures kept | 0 | 46.9 | 0.82 / 0.28 | False | {'pages': 13, 'slides': 13, 'seconds': 42.481, 'seconds_per_page': 3.268, 'fallback_pages': [], 'errors': [], 'tune': 'ran'} |
| long_report_300p | converts; time and peak RAM logged | 0 | 1461.4 | 6.06 / 0.47 | False | {'pages': 307, 'slides': 307, 'seconds': 1457.198, 'seconds_per_page': 4.747, 'fallback_pages': [], 'errors': [], 'tune': 'ran'} |
| no_powerpoint_convert | converts without the tuner / checks | 0 | 8.5 | 0.16 / 0.22 | False | {'pages': 1, 'slides': 1, 'seconds': 6.394, 'seconds_per_page': 6.394, 'fallback_pages': [], 'errors': [], 'tune': 'PowerPoint unavailable: cannot start PowerPoint via COM: OSError: PowerPoint.Application: class not regi |
| no_powerpoint_deck | extractive deck, fit check skipped | 0 | 5.9 | 0.12 / 0.22 | False | {'pages': 1, 'slides': 4, 'seconds': 3.83, 'errors': []} |
| no_powerpoint_score | scores without the PowerPoint metrics | 0 | 5.9 | 0.15 / 0.22 | False | 17:41:22 WARNING pdf2pptx.verify.score: PowerPoint unavailable: PowerPointUnavailable: cannot start PowerPoint via COM: OSError: PowerPoint.Application: class not registered (simulated by PDF2PPTX_NO_POWERPOINT=1) |
| not_a_pdf | exit code 3, message says not a PDF | 3 | 1.7 | 0.05 / 0.21 | False | pdf2pptx: not_a_pdf.pdf: not a PDF file (no %PDF header) |
