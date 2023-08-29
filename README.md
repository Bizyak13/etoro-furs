# Etoro -> FURS csv converter

Etoro to FURS conversion of excel file to csv, to easily import the data when submitting the `doh_div` to the Slovenian FURS. This is more of a "quick & dirty" implementation, but it should work fine to submit.

There are still some improvements that can be made, which I will do if/when time permits.

-   Load data files from github on every run to get current data
-   Write up some tests
-   Support other platforms??

## Requirements

1. Have Python 3 installed
2. Download this package to your computer and extract it
3. Open `CONFIG.cfg` and input your national(slovenian) `TAX_ID` number and `DIVIDEND_TYPE`
4. Install the required packages from the `requirements.txt` (currently none)

## Usage

Run the app with `python etoro-furs.py`

Use -h to display the help text, to see what commands are available.

## Issues/Troubleshooting

If you are having issues, or a company you are using isn't on the list, create a new issue here on Github and I will take a look.
