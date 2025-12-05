# AmorPy

AmorPy is a desktop application designed to calculate and create asset amortization schedules with a clean and easy-to-use interface.

## App Release

Download latest app in [release page](https://github.com/bintanngg/AmorPy/releases)

## Main Features v1.3

*   **Modern Interface:** Built using `ttkbootstrap` for a modern and responsive look.
*   **Two Amortization Methods:** Supports the two most commonly used calculation methods:
    *   Garis Lurus (Straight-Line)
    *   Saldo Menurun Ganda (Double Declining Balance)
*   **Export to Excel:** Save the generated amortization schedule to an `.xlsx` file format with number formatting already adjusted for readability.
*   **Input Validation:** Equipped with input data validation to ensure accuracy, including flexible date formats and correct numeric values.
*   **High-Precision Calculation:** Uses the `Decimal` data type for all financial calculations, ensuring accuracy down to the last decimal place.

## Application Screenshot

*(You can add application screenshots here to clarify the appearance)*

## Prerequisites

Before you begin, ensure you have met the following requirements:
*   Python 3.x
*   Windows/Linux/Mac Operating System

## Installation and Usage

1.  **Clone repository:**
    ```sh
    git clone https://github.com/bintanngg/AmorPy.git
    ```

2.  **Navigate to the v1.3 project directory:**
    ```sh
    cd AmorPy/v1.3
    ```

3.  **Install the required dependencies:**
    It is recommended to create a *virtual environment*.
    ```sh
    pip install -r requirements.txt
    ```

4.  **Run the application:**
    ```sh
    python app.py
    ```

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.