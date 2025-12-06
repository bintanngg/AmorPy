"""
AmorPy - Amortization & Depreciation Calculator
A simple tool to generate amortization schedules.
Version: 1.5
Check my other projects at https://github.com/bintanngg
"""

from gui import AmortizationApp

def main() -> None:
    """Start the GUI application."""
    app = AmortizationApp()
    app.mainloop()

if __name__ == "__main__":
    main()
