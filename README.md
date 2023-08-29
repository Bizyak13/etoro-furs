# Etoro -> FURS csv converter

Etoro to FURS conversion of excel file to csv, to easily import the data when submitting the `doh_div` to the Slovenian FURS. This is more of a "quick & dirty" implementation, but it should work fine to submit.

There are still some improvements that can be made, which I will do if/when time permits.

-   Load data files from github on every run to get current data
-   Write up some tests
-   Support other platforms??

## Install

1. Have [Python3](https://www.python.org/downloads/) installed
2. Download this package to your computer and extract it
    - Click the green code button and select "Download ZIP"
    - Exctract the zip to a folder on your computer
3. Open `CONFIG.cfg` and input your national(slovenian) `TAX_ID` number and `DIVIDEND_TYPE`
4. Install the required packages from the `requirements.txt`
    - `pip install -r requirements.txt` Run this command in the terminal

## Usage

First get the dividends report file from etoro [account statement page](https://www.etoro.com/documents/accountstatement) and save it in the same folder. (It is best to create reports for only 1 year at a time)

Then, run the app with `python etoro-furs.py` by opening a new terminal window in the folder you extracted in the previous step.

Use -h to display the help text, to see what commands are available.

```
etoro-furs: Running etoro-furs
usage: etoro-furs.py [-h] [-v] input output

positional arguments:
  input          Input file from etoro. Must be in .xlsx format.
  output         Output file from csv. Must be in .csv format.

options:
  -h, --help     show this help message and exit
  -v, --verbose  Verbose output
```

## Issues/Troubleshooting

If the edavki.durs.si is returning an error when submitting, it means that some data in the csv is either missing or wrong. Check if your TAX_ID is correctly set, and that the company data is present. (If not, keep reading)

If you are having any other issues, or a company you are using isn't on the list, create a new issue here on Github and I will take a look.
