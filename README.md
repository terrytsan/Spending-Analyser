<h1 align="center">Spending Analyser</h1>

<div align="center">A python script to analyse spending data stored in .xlsx format.</div>

## About

* Imports spending data (.xlsx) and outputs a summary report in a new .xlsx file.
    * Configure file paths in `settings.toml`.
* Analysis determines average cost per day of an item over its lifetime and uses this to determine cost within a given
  set of dates.
* Specify list of dates in `settings.toml` for which analysis should be performed. For each date range, the following is
  included in the resulting report:
    * Total cost by item category.
    * List of items used within the date range (retains the same columns that are in input .xlsx).
    * Summary table comparing cost of each category across all date ranges.

* Written for python 3.8.0

## Prerequisites

Spreadsheet under analysis should have the following:

* Sheet name called 'Spending' with a table named  'Spending'.
* 'Spending' table should contain at least the following columns:
    * Category (category of the item - string)
    * Cost (cost of the item - number)
    * Date Started (date the item was first used - date)
    * Date Finished (date the item was finished - date)

## Usage

1. Install python packages listed in `requirements.txt`.
2. Set up `settings.toml` with the following:
    1. `InputFileName` - path of .xlsx file to be analysed.
    2. `OutputFileName` - path of output file.
    3. `Queries` - list of date ranges determining how the data will be analysed.
        * Dates can optionally include time component e.g. `2022-03-01T01:11:00` or `2022-03-01` are both valid formats.
          If time component is omitted, 00:00:00 is assumed.