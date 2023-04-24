# Excel VBA Monte Carlo Trade Simulator

This Excel VBA Monte Carlo Trade Simulator is a helpful tool that assists you in evaluating the potential outcomes of your trading strategies using Monte Carlo Simulation techniques. It allows you to simulate multiple trading scenarios and analyze their performance to gain a better understanding of the risks and returns of your strategies.

![Screenshot](./screenshots/screenshot.png)

## Features

- Evaluate the performance of your trading strategies using Monte Carlo simulation
- Calculate the risk of ruin, median profit, median drawdown, and other key metrics for different starting equity amounts
- Analyze the impact of different lot sizes and number of trades per year
- Generate detailed equity curve data for further analysis

## Installation

To use the Excel VBA Monte Carlo Trade Simulator, simply download and open the Excel file `monte_carlo_trade_simulator.xlsm`. The other files in this repository are the VBA code files that already exist in the tool, but have been exported for versioning purposes on GitHub.

## Usage

1. Prepare a list of trade PNL (profit and loss) data for your trading strategy.
2. Paste the trades into the sheet labeled "InputData" and set the required parameters for the simulation, such as lot size, number of trades per year, total number of runs, starting equity, and margin.
3. Press the button "Run" on the "Control" worksheet to run the Monte Carlo simulation and obtain the simulation results.
4. Analyze the results to evaluate the performance of your trading strategy.

## Contributing

We welcome contributions to improve the Excel VBA Monte Carlo Trade Simulator. Please feel free to open an issue or submit a pull request.

## License

This project is released under the MIT License. Please refer to the `LICENSE` file for more information.
