"""
Kalkulator Amortisasi
Check my other projects at https://github.com/bintanngg
"""

from gui import AmortizationApp

def main():
    app = AmortizationApp(themename="journal")  # Initialize the GUI application with a theme
    app.mainloop() # Start the GUI event loop

if __name__ == "__main__":
    main()
