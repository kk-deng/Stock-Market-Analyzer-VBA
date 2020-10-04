# Backgound
We have a workbook with three years of stock data for hundreds of tickers for the whole year. A VBA script was built to analyze the data, 
getting open price, close price, total volume for each ticker.

A summary table of stock data was generated for each year, along with the greatest increase and decrease percentage, total volume in the year.

# Script Functions
This script contains four subs in the code:
1. Main run: the main run to call functions from the subs below. A user can choose to reset cells prior to processing the data.
<a href="https://imgbb.com/"><img src="https://i.ibb.co/MG3ZqJ2/8b217f30d7877632551e81185c046f7.png" alt="8b217f30d7877632551e81185c046f7" border="0"></a><br />

2. StockYearlyChangeFast() sub performs the main analysis on the data. It loops the whole table to get every ticker name, yearly price change percentage and total volume.
List them on the summary table side to view (Yearly Change)	(Percent Change) (Total Stock Volume).

3. GreatestYearlyChange() sub reads summary data generated from the no.2 function. It compares all price change to get the highest, lowest price change and highest volume tickers.

4. ResetCells() sub is an optional function for users to reset summary data on the right side of excel worksheets for a new run.

# Features:

* Once user starts the run, a progress update will be displayed on the bottom left corner of the excel file to update the current progress. 
  (Only update every 5000 rows to lower the CPU usage)

<a href="https://imgbb.com/"><img src="https://i.ibb.co/ZX2Wr62/d3763c55805031c599ef68fee4fedcf.png" alt="d3763c55805031c599ef68fee4fedcf" border="0"></a>

* An elapsed time will be shown on the status bar to estimate the overall time used.

<a href="https://imgbb.com/"><img src="https://i.ibb.co/Fmy1yJm/8535d724b67bbed5ebfe592c3479828.png" alt="8535d724b67bbed5ebfe592c3479828" border="0"></a>

* The yearly price change cells are conditionally formatted with green and red colours to represent positive and negative changes.

<a href="https://imgbb.com/"><img src="https://i.ibb.co/jDh0mQH/image.png" alt="image" border="0"></a>

* Column O and Q were set as auto adjustment to fit the cell for a better view.

        .Columns("O").AutoFit
        .Columns("Q").AutoFit

# Results

This script was optimized for the given excel files. A sorted data assumption was made to get the best efficieny. 

Overall time used to analyze the test results of the multi-year datasets (2014, 2015 and 2016) is between <b>37s to 45s</b> depending on the specs of the computer.

Please check enclosed screenshots of detailed results in the folder.
