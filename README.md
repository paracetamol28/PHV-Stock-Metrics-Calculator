
# Stock Metrics Calculator

This project is a Stock Metrics Calculator that fetches financial data for a given stock symbol and calculates various metrics such as YTD return, revenue growth, net income growth, short interest of float, last closing price, EV/(EBITDA-capex) and operating margin. The calculated metrics are then saved in an Excel file.





## Prerequisites

- Python 3.x
- Required Python packages: pandas, yfinance, openpyxl

## Installation

1. Clone the repository:

```bash
    git clone https://github.com/your-username/stock-metrics-calculator.git
```
2. Change to the project directory:

```bash
    cd stock-metrics-calculator
```
3. Install the required Python packages using pip:

```bash
    pip install -r requirements.txt
```


   




    
## Usage

1. Open the main.py file and set the stock_symbol variable to the desired stock symbol.

2. Run the script:

```bash
python main.py
```
This will fetch the financial data, calculate the metrics, and generate an Excel file named stock_metrics.xlsx.

3. Open the Excel file to view the calculated metrics.


## Running on AWS server

1. Set up an AWS EC2 instance with an appropriate operating system (e.g., Amazon Linux, Ubuntu).

2. Connect to the EC2 instance using SSH.

3. Install Python and pip on the EC2 instance. You can use the following commands for Amazon Linux:

```bash
    sudo yum update -y
    sudo yum install python3 -y
    sudo yum install python3-pip -y
```

4. Clone the repository on the EC2 instance:

```bash
    git clone https://github.com/your-username/stock-metrics-calculator.git
```
5. Change to the project directory:

```bash
    cd stock-metrics-calculator
```

6. Install the required Python packages using pip:

```bash
    pip3 install -r requirements.txt
```

7. 7. Open the `main.py` file and set the `stock_symbol` variable to the desired stock symbol.

8. Run the script:

```bash
python3 main.py
```
This will fetch the financial data, calculate the metrics, and generate an Excel file named `stock_metrics.xlsx` on your AWS EC2 instance.

9. Download the Excel file from the EC2 instance to your local machine using SCP or any other method.

10. Open the Excel file to view the calculated metrics.

## Software design 

- The modular design allows for flexibility and easy extensibility. Additional functionality or calculations can be added by modifying the respective components without affecting the rest of the codebase.

- The object-oriented design promotes code organization, reusability, and maintainability. Each component has a well-defined responsibility and can be independently tested and modified.

## Customization

- You can modify the code in main.py to include additional metrics or customize the calculations based on your requirements.

- If you want to change the decimal digits or add specific symbols (e.g., %, x) to the Excel file, you can modify the create_excel_sheet() function in calculator.py.

## Future Improvements

- Implement unit tests to ensure the correctness of the calculations and functionality.
- Add support for fetching data from other financial APIs or data sources.
- Enhance error handling and logging for better debugging and error reporting.
- Implement caching mechanisms to reduce API calls and improve performance.
- Provide a command-line interface (CLI) or a web-based interface for user interaction and customization.

The software design aims to balance simplicity, readability, and functionality, allowing for further enhancements and improvements in the future.

## License

[MIT](https://choosealicense.com/licenses/mit/)



