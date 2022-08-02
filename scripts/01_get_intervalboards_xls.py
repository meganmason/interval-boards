import glob
from pathlib import Path
import shutil

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# SET UP

# paths
path_in = Path('/Users/mamason6/Documents/snowex/campaigns/TS-21/0_nsidc-raw-1Feb22/') # input path
path_out = Path('/Users/mamason6/Documents/snowex/core-datasets/ground/interval-boards/output/raw_xls/') # output path
path_out.mkdir(parents=False, exist_ok=True) # create directory if it doesn't exist

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FIND AND COPY FILES

# loop files and copy to output directory
for filename in path_in.rglob('*INTERVAL_BOARD.xlsx*'):
    # print('input file .../', str(filename)) # input
    # print('output file .../', str(path_out.joinpath(filename.name))) # output
    shutil.copy(str(filename), str(path_out.joinpath(filename.name)))


# check and remove unwanted (i.e. corrupt) files that transfered from nsidc
for filename in path_out.glob('*~$*'): # ~$ = corrupt file, rm
    print('\nchecking for corrupted files...')
    filename.unlink() # pathlib's os.remove()
    print('files removed ....: ', filename.name)

print('...script complete, open ../snowex/core-datasets/ground/interval-boards/output/raw_xls to see files')
