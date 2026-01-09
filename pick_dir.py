#!/usr/bin/env python
import tkinter as tk
from tkinter import filedialog
import json
import sys

def main():
    try:
        initial = None
        if len(sys.argv) > 1:
            initial = sys.argv[1]
        root = tk.Tk()
        root.withdraw()
        # if initial directory provided, pass it to askdirectory
        if initial:
            selected = filedialog.askdirectory(title='Select template directory', initialdir=initial)
        else:
            selected = filedialog.askdirectory(title='Select template directory')
        try:
            root.destroy()
        except Exception:
            pass
        # print JSON to stdout
        print(json.dumps({'selected': selected or ''}))
        return 0
    except Exception as e:
        print(json.dumps({'error': str(e)}))
        return 2

if __name__ == '__main__':
    sys.exit(main())
