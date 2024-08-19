import glob
import os


# Function to get files matching the patterns
def get_files(directory, patterns):
    files = []
    for pattern in patterns:
        files.extend(glob.glob(os.path.join(directory, pattern)))
    return files