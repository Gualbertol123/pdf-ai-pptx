"""Entry point of the packaged exe (PyInstaller).

``freeze_support`` must run first: the PowerPoint COM worker and the page-parsing pool
are ``multiprocessing`` children, which re-launch this exe.
"""

import multiprocessing
import sys

if __name__ == "__main__":
    multiprocessing.freeze_support()
    from pdf2pptx.cli import main

    sys.exit(main())
