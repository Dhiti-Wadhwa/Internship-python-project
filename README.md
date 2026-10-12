# Science Birthday Greetings Project

## Files
- `main.py`: cleans the data, matches birthdays, and generates greetings.
- `members.xlsx`: sample member list.
- `scientists.xlsx`: scientist birthdays displayed as readable dates, plus original `birth_year`.
- `birthday_greetings.xlsx`: generated output after running the program.
- `requirements.txt`: dependencies.

## Run
Activate your virtual environment in VS Code, then run `python main.py`.

Each run updates `scientists.xlsx` with readable dates such as `04-Jan-1643` and the separate original `birth_year` column, then generates `birthday_greetings.xlsx`. Matching uses month and day only. Historical dates are stored as readable text because some dates are too old for standard Excel/Pandas datetime support.
