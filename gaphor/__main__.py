import sys
import os

sys.path.append(sys.path[0])
del sys.path[0]

from gaphor.main import main

if __name__ == "__main__":
    sys.exit(main(sys.argv))
