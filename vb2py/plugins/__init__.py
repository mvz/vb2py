import glob
import os

from vb2py.utils import modulePath

mods = []
for fn in glob.glob(os.path.join(modulePath(), "plugins", "*.py")):
    name = os.path.splitext(os.path.basename(fn))[0]
    if not name.startswith("_"):
        mods.append(name)
