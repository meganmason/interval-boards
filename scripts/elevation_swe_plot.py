import datetime
import glob
# import os
# import shutil
import numpy as np
import pandas as pd
from csv import writer
from pathlib import Path
import elevation
import utm
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import seaborn as sns

out_path = Path('/Users/mamason6/Documents/snowex/core-datasets/ground/interval-boards/output')

elevation.clip(bounds=(12.35, 41.8, 12.65, 42), output=Path.joinpath(out_path, 'Rome-DEM.tif'))
# elevation.clip(bounds=(-125, 32, -110, 50), output=Path.joinpath(out_path, 'westUS-DEM.tif'))
# elevation.clip(bounds=(-115.5, 39, -114.5, 45), output=Path.joinpath(out_path, 'westUS-DEM.tif'))
