# AmorPy

AmorPy is a desktop application designed to calculate and create asset amortization & depreciation schedules with a clean and easy-to-use interface.

## App Release

Download latest app in [release page](https://github.com/bintanngg/AmorPy/releases)

## Main Features v1.5

*   **Clean Interface:** Built with Python's standard `tkinter` library for a native and lightweight user experience.
*   **Three Amortization Methods:** Supports three of the most commonly used calculation methods:
    *   Straight-Line
    *   Double Declining Balance
    *   Sum of the Years' Digits
*   **Export to Excel:** Save the generated amortization schedule to an `.xlsx` file format with number formatting already adjusted for readability.
*   **Input Validation:** Equipped with input data validation to ensure accuracy, including flexible date formats and correct numeric values.
*   **High-Precision Calculation:** Uses the `Decimal` data type for all financial calculations, ensuring accuracy down to the last decimal.

## Application Screenshot

***AmorPy UI***

<img width="692" height="578" alt="1" src="https://github.com/user-attachments/assets/6cfa0765-f1d2-4a14-b93c-2877a3345be6" />

#

***Output xlsx File***

<img width="721" height="545" alt="2" src="https://github.com/user-attachments/assets/841840bd-138a-4503-bdf8-419ebf87e120" />

## Prerequisites

Before you begin, ensure you have met the following requirements:
*   Python 3.x
*   Windows/Linux/Mac Operating System

## Installation and Usage

1.  **Clone repository:**
    ```sh
    git clone https://github.com/bintanngg/AmorPy.git
    ```

2.  **Navigate to the project directory:**
    ```sh
    cd AmorPy/v1.5
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