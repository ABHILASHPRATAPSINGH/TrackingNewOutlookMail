import pathlib
import os

def getSrcFolderPath():
    currentFilePath=pathlib.Path(os.path.abspath(__file__))
    currentFilename=pathlib.Path(os.path.basename(__file__))
    currentFolderPath=os.getcwd()
    return os.path.join(currentFolderPath,'src')


