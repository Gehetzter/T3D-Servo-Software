import tkinter as tk
from gui import App

# Toggle transport debug logging (prints raw frames):
# Set to True to enable logs like "01:28:07.063110 COM -> : 02 06 00 62 00 00 28 27"
TRANSPORT_DEBUG = True


def main():
    xml_path = 'config/parameter_SpindleHS1.xml'
    # Window size constants
    DEFAULT_WIDTH = 1200
    DEFAULT_HEIGHT = 800
    MIN_WIDTH = 900
    MIN_HEIGHT = 500

    root = tk.Tk()
    # set a larger default window size so parameter tables are visible without manual resize
    root.geometry(f'{DEFAULT_WIDTH}x{DEFAULT_HEIGHT}')
    root.minsize(MIN_WIDTH, MIN_HEIGHT)
    app = App(root, xml_path, transport_debug=TRANSPORT_DEBUG)
    root.mainloop()


if __name__ == '__main__':
    main()
