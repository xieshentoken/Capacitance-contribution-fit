import os
import re
import sys
import time
from collections import OrderedDict
from itertools import permutations
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog, ttk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from add_GUI import App
from func_XX import Orz




def main():
    root = Tk()
    root.title("A little App")
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
