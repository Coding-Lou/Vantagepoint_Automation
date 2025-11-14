from tqdm import tqdm
import sys

class DualOutput:
    def __init__(self, file_path):
        self.terminal = sys.stdout
        self.log = open(file_path, 'w', encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()

    def tqdm_write(self, message):
        tqdm.write(message)
        self.log.write(message + "\n")
        self.log.flush()
