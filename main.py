import re
import sys
from collections import OrderedDict
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.signal as signal
import xlrd

from add_GUI import App
from func_XX import Orz


def main():
    root = Tk()
    root.title("Nacher")
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
