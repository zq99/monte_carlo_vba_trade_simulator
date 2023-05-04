# Excel VBA Monte Carlo Trade Simulator

This Excel VBA Monte Carlo Trade Simulator is a helpful tool that assists you in evaluating the potential outcomes of your trading strategies using Monte Carlo Simulation techniques. It allows you to simulate multiple trading scenarios and analyze their performance to gain a better understanding of the risks and returns of your strategies.

![Screenshot](/screenshots/screenshot.PNG)

## Features

- Evaluate the performance of your trading strategies using Monte Carlo simulation
- Calculate the risk of ruin, median profit, median drawdown, and other key metrics for different starting equity amounts
- Analyze the impact of different lot sizes and number of trades per year
- Generate detailed equity curve data for further analysis

## Installation

To use the Excel VBA Monte Carlo Trade Simulator, simply download and open the Excel file `monte_carlo_trade_simulator.xlsm`. The other files in this repository are the VBA code files that already exist in the tool, but have been exported for versioning purposes on GitHub.

## Usage

1. Prepare a list of trade PNL (profit and loss) data for your trading strategy.
2. Paste the trades into the sheet labeled "InputData". 
3. On the "Control" sheet, set the required parameters for the simulation, such as lot size, number of trades per year, total number of runs, starting equity, and margin.
3. Press the button "Run" on the same sheet to run the Monte Carlo simulation and obtain the simulation results.
4. Analyze the results to evaluate the performance of your trading strategy.

## Disclaimer

Use this spreadsheet at your own risk, I make no claim as to the accuracy or validity of the results obtained. The project is open source for you to do
your own analysis and due dilligence.

## Contributing

Contributions are welcome, please feel free to open an issue or submit a pull request.

The module 'Test_Module_All" contains a subroutine called "RunAllTests" which can be used to test for any breaking changes after an amendment.

## License

This project is released under the MIT License. Please refer to the `LICENSE` file for more information.

## Acknowledgements

This project is an original implementation of an idea proposed by Kevin Davey.
