# Pokebinder

A tool for retrieving data and prices for your Pokemon TCG Collection.

## Getting Started

### Prerequisites

- Pythom 3.10 or higher
- Internet Connection
- An excel file 

### Installation

- Windows Executable (portable) file. No installation required

## Functions and Features

This project requires an excel file that has stored in it Card Name, Expansion Set and Collector's Number of the cards. Then using the Pokemon TCG Api, it retrieves market average price from TCG Player's real-time database and stores it inside the excel file. You can also retrieve the specific Set-ID and your card's rarities. 

Other features include:
- Make multiple new card entries to your binder
- Update All cards prices
- Update only missing prices
- Last time's total binder value
- Total value after modifications
- Display of your TOP 10 most valuable cards 
- Retrieve Card Rarities 
- Counts for Card Rarities 
- Binder total value

## Built With

- Python Idle
- Jupyter Notebook
- Pokemon TCG API

This project was built using the Pokemon TCG API, which allows us to retrieve information about Pokemon trading cards.

### Requirements

This project requires the following libraries to be installed:

- pandas
- requests
- beautifulsoup4
- tqdm
- openpyxl
- pyfiglet
- pokemontcgsdk

 Can be installed using pip:

```bash
pip install pandas requests beautifulsoup4 tqdm openpyxl pyfiglet pokemontcgsdk

```

### Additional Libraries
Additionally, this project requires the following standard Python libraries:

- datetime
- time
- os
- tkinter
- re
- decimal

These libraries are included in the Python standard library and do not need to be installed separately.


### Contact
If you have any questions, feedback, or suggestions for this project, please contact me:

- Email: vachtsavanisk@hotmail.com

- Twitter: @constantinevac

- Instagram: @constantinevac
