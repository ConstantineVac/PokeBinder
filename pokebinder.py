import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time
import os
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm #progress bar
import math
import openpyxl
import pyfiglet
from pyfiglet import Figlet

global df
global total_cards

df = None
total_cards = 0

#Function1 : Add Cards
# Option 1: Add new card entry
#ask number of the same car
def add_cards():
    global df
    global total_cards

    while True:
        try:
            num_cards = int(input("How many duplicates would you like to add? "))
            break
        except ValueError:
            print("Please enter a valid number.")
        
    # Get card info from user
    
    while True:
        card_name = input("Please enter the card name(as written on Card): ").replace('_V',' V')
        if card_name.isalpha():
            break
        else:
            print("Please enter a valid card name containing only letters.")
        
    while True:
        set_name = input("Please enter the expansion set(ex.Lost Origin): ")
        if set_name.isalpha():
            break
        else:
            print("Please enter a valid set name containing only letters.")

    while True:    
        card_version = input("Please enter the card version (ex.V1 or None):")
        if card_version.isalpha():
            break
        else:
            print("Please enter a valid card version")

    while True:    
        set_id = input("Please enter the set ID (ex.LOR): ").upper()
        if set_id.isalpha():
            break
        else:
            print("Please enter a valid set ID")
        
    while True:
        try:    
            card_number = input("Please enter the card number(ex.131): ")
            break
        except ValueError:
            print("Please enter a valid number.")
        
    # Add new row to DataFrame
    for i in range(num_cards):
        new_row = {'Card Name': card_name.replace('_V',' V'), 'Expansion Set': set_name.replace(' ','-'), 'Set ID': set_id.upper(), 'Card Version': card_version, 'Set Number': card_number.zfill(3)}
        df = df.append(new_row, ignore_index=True)

        # Generate URL and add to DataFrame
        card_version = df.at[df.index[-1], 'Card Version']
        card_number = df.at[df.index[-1], 'Set Number'].zfill(3)
        set_id = df.at[df.index[-1], 'Set ID']
        set_name = df.at[df.index[-1], 'Expansion Set']
        
        # Construct cardmarket URL
        if card_version=='None':
            card_version_str = ''
        else:
            card_version_str = f'{card_version}-'
            
        url = f"https://www.cardmarket.com/en/Pokemon/Products/Singles/{set_name}/{card_name.replace(' ', '-')}-{card_version_str}{set_id.replace(' ','')}{card_number}"
        df.at[df.index[-1], 'URL'] = url

    # Save changes to Excel file
    df.to_excel(excel_path, index=False)
    print("\nNew card entry has been added to the Excel file!")
        
    #Binder Status Info after the new entries
    total_cards = len(df)
    total_na_rows = df["Price"].isna().sum()
    now = datetime.now()
    last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
    
    if not math.isnan(total_price) and total_price > 0:
        total_price = df['Price'].sum(skipna=True)
        print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
    else:
        print(f'Your Binder has in total {total_cards} cards. Run 2 to get a total value !')
        
    if total_na_rows > 0:
        print('Warning Missing Prices !')
        if total_na_rows == 1:
            print(f'There was found {total_na_rows} card with missing price')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing price!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
        else:
            print(f'There were found {total_na_rows} cards with missing prices')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing prices!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
    print(last_checked)
        
    pass




#Function2 Update Card Prices 
#Update Cards Prices using CardMarket's 7-day average!
def update_card_prices():
    global df
    global total_cards
    try:    
        print("\nCard prices will be updated based on the cardmarket.com 7-day average.")    
        #total_cards = df.shape[0]
        #completed_cards = 0
    
        print("Updating Prices:") 
        with tqdm(total=total_cards) as pbar:
            for index,row in df.iterrows():
                card_name = row['Card Name']
                set_name = row['Expansion Set']
                set_id = row['Set ID']
                card_version = row['Card Version']
                card_number = row['Set Number']  
    
                # Extract the URL for the card from the DataFrame
                url = row['URL']

                # Make a GET request to the URL and parse the HTML with BeautifulSoup
                response = requests.get(url)
                soup = BeautifulSoup(response.content, 'html.parser')
    
                # Find the 'dd' element containing the 7-day average price
                avg_price_dt = soup.find('dt', string='7-days average price')
                avg_price_dd = avg_price_dt.find_next_sibling('dd')
                price_span = avg_price_dd.find('span')
                avg_7_day_price = price_span.text.strip()
    
                # Remove and Replace both comma and euro sign
                avg_7_day_price_nosign = avg_7_day_price.replace("€", "").strip()
                avg_7_day_price_comma = avg_7_day_price_nosign.replace(",", '.').strip()
                avg_7_day_float_price = float(avg_7_day_price_comma)
            
    
    
                # Update the 'Price' and 'Timestamp' values in the DataFrame
                df.at[index, 'Price'] = avg_7_day_float_price
                df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                pbar.update(1)
        
        # Save the updated data back to the Excel file
        df.to_excel(excel_path, index=False)
        print("\nPrices updated successfully!")
    
        total_cards = len(df)
        total_price = df['Price'].sum(skipna=True)
        total_na_rows = df["Price"].isna().sum()
        now = datetime.now()
        last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
        print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
        if total_na_rows > 0:
            print('Warning Missing Prices !')
            if total_na_rows == 1:
                print(f'There was found {total_na_rows} card with missing price')
                print("")
                print('It is suggested that you run Option 2 in order to retrieve missing price!')
                nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
                print(nan_card_names)
            else:
                print(f'There were found {total_na_rows} cards with missing prices')
                print("")
                print('It is suggested that you run Option 2 in order to retrieve missing prices!')
                nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
                print(nan_card_names)
        print(last_checked)
        
    except Exception as e:
        print(f"An error occurred: {e}")
        pass



# Option 3: Add new card entry and update prices
def both():
    
    global df
    global total_cards
    num_cards = int(input("How many duplicates would you like to add? "))
    # Get card info from user
    card_name = input("Please enter the card name(as written on Card): ").replace('_V',' V')
    set_name = input("Please enter the expansion set(ex.Lost-Origin): ")
    card_version = input("Please enter the card version (ex.V1):")
    set_id = input("Please enter the set ID (ex.LOR): ")
    card_number = input("Please enter the card number(ex.131): ")

    # Add new row to DataFrame
    for i in range(num_cards):
        new_row = {'Card Name': card_name.replace('_V',' V'), 'Expansion Set': set_name.replace(' ','-'), 'Set ID': set_id.upper(), 'Card Version': card_version, 'Set Number': card_number.zfill(3)}
        df = df.append(new_row, ignore_index=True)

        # Generate URL and add to DataFrame
        card_version = df.at[df.index[-1], 'Card Version']
        card_number = df.at[df.index[-1], 'Set Number'].zfill(3)
        set_id = df.at[df.index[-1], 'Set ID']
        set_name = df.at[df.index[-1], 'Expansion Set']
        
        # Construct cardmarket URL
        if card_version=='None':
            card_version_str = ''
        else:
            card_version_str = f'{card_version}-'
            
        url = f"https://www.cardmarket.com/en/Pokemon/Products/Singles/{set_name}/{card_name.replace(' ', '-')}-{card_version_str}{set_id.replace(' ','')}{card_number}"
        df.at[df.index[-1], 'URL'] = url

    # Save changes to Excel file
    df.to_excel(excel_path, index=False)
    print("\nNew card entry has been added to the Excel file!")
    print("")
    print("\nCard prices will be updated based on the cardmarket.com 7-day average.")
    print("")
    print("")
    print("Updating Prices:") 
    with tqdm(total=total_cards) as pbar:
        for index,row in df.iterrows():
            card_name = row['Card Name']
            set_name = row['Expansion Set']
            set_id = row['Set ID']
            card_version = row['Card Version']
            card_number = row['Set Number']  
    
            # Extract the URL for the card from the DataFrame
            url = row['URL']

            # Make a GET request to the URL and parse the HTML with BeautifulSoup
            response = requests.get(url)
            soup = BeautifulSoup(response.content, 'html.parser')
    
            # Find the 'dd' element containing the 7-day average price
            avg_price_dt = soup.find('dt', text='7-days average price')
            avg_price_dd = avg_price_dt.find_next_sibling('dd')
            price_span = avg_price_dd.find('span')
            avg_7_day_price = price_span.text.strip()
    
            # Remove and Replace both comma and euro sign
            avg_7_day_price_nosign = avg_7_day_price.replace("€", "").strip()
            avg_7_day_price_comma = avg_7_day_price_nosign.replace(",", '.').strip()
            avg_7_day_float_price = float(avg_7_day_price_comma)
            
    
    
            # Update the 'Price' and 'Timestamp' values in the DataFrame
            df.at[index, 'Price'] = avg_7_day_float_price
            df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            pbar.update(1)
        
    # Save the updated data back to the Excel file
    df.to_excel(excel_path, index=False)
    print("\nPrices updated successfully!")
    
    
    #Binder Status Info after the new entries
    total_cards = len(df)
    total_na_rows = df["Price"].isna().sum()
    now = datetime.now()
    last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
    total_price = df['Price'].sum(skipna=True)
    print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
        
    if total_na_rows > 0:
        print('Warning Missing Prices !')
        if total_na_rows == 1:
            print(f'There was found {total_na_rows} card with missing price')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing price!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
        else:
            print(f'There were found {total_na_rows} cards with missing prices')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing prices!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
    print(last_checked)
        
    pass



#4 Update only cards with missing card values!
def update_missing():
    global df
    global total_cards
    
    #find missing prices
    rows_missing_price = df[df['Price'].isna()]
    total_cards = len(rows_missing_price)
    
    print("Updating Prices based on the cardmarket.com 7-day average:")
    # Update missing prices
    with tqdm(total=total_cards) as pbar:
        for index, row in rows_missing_price.iterrows():
            card_name = row['Card Name']
            set_name = row['Expansion Set']
            set_id = row['Set ID']
            card_version = row['Card Version']
            card_number = row['Set Number']  
    
            # Extract the URL for the card from the DataFrame
            url = row['URL']

            # Make a GET request to the URL and parse the HTML with BeautifulSoup
            response = requests.get(url)
            soup = BeautifulSoup(response.content, 'html.parser')
    
            # Find the 'dd' element containing the 7-day average price
            avg_price_dt = soup.find('dt', string='7-days average price')
            avg_price_dd = avg_price_dt.find_next_sibling('dd')
            price_span = avg_price_dd.find('span')
            avg_7_day_price = price_span.text.strip()
    
            # Remove and Replace both comma and euro sign
            avg_7_day_price_nosign = avg_7_day_price.replace("€", "").strip()
            avg_7_day_price_comma = avg_7_day_price_nosign.replace(",", '.').strip()
            avg_7_day_float_price = float(avg_7_day_price_comma)
            
    
    
            # Update the 'Price' and 'Timestamp' values in the DataFrame
            df.at[index, 'Price'] = avg_7_day_float_price
            df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            pbar.update(1)
        
    # Save the updated data back to the Excel file
    df.to_excel(excel_path, index=False)
    print("\nPrices updated successfully!")
    
    
    #Binder Status Info after the new entries
    total_cards = len(df)
    total_na_rows = df["Price"].isna().sum()
    now = datetime.now()
    last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
    total_price = df['Price'].sum(skipna=True)
    print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
        
    if total_na_rows > 0:
        print('Warning Missing Prices !')
        if total_na_rows == 1:
            print(f'There was found {total_na_rows} card with missing price')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing price!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
        else:
            print(f'There were found {total_na_rows} cards with missing prices')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing prices!')
            nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
            print(nan_card_names)
    print(last_checked)
        
    pass

        
#Function 6
# Calculate total binder value

def calculate_binder_price():
    df = pd.read_excel(excel_path)
    total_price = df['Price'].sum(skipna=True)
    print('')
    print(f'The Total Value of your binder is {total_price.round(2)}€') 
    pass
    
#Function 5
# Display top 10 most expensive cards

def show_top_10_expensive_cards():
    #top = df.sort_values(by='Price', ascending=False).head(10)
    top_10 = top_10 = df.nlargest(10, 'Price')[['Card Name', 'Price']]
    print('')
    print('Your Top 10 Most Expensive Cards are:')
    print(top_10)
    pass

now = datetime.now()
#custom_fig = Figlet(font='cybermedium')
print("*********************************************************")
print("*                                                       *")
print("*                  Thank you for using                  *")
print("*                                                       *")
print("*                       PokeBinder                      *")
#print(custom_fig.renderText('  PokeBinder'))              
print("*                                                       *")
print("*                         v1.1                          *")
print("*                Made by ConstantineVac                 *")
print("*                                                       *")
print("*                                                       *")
print("*          Consider Supporting me on Youtube            *")
print("*                                                       *")
print(f"*           Started: {now.strftime('%b %d, %Y at %H:%M')}              *")
print("*                                                       *")
print("*********************************************************")
print("")




#Main Body
def main():    
    

    if choice == '1':
        add_cards()
        df = pd.read_excel(excel_path)
        
    elif choice == '2':
        update_card_prices()
        df = pd.read_excel(excel_path)
        
    elif choice == '3':
        both()
        df = pd.read_excel(excel_path)
        
    elif choice == '4':
        update_missing()
        df=pd.read_excel(excel_path)

    elif choice == '5':
        show_top_10_expensive_cards()
        df = pd.read_excel(excel_path)
                           
    elif choice == '6':
        calculate_binder_price()
        df = pd.read_excel(excel_path)



# Greet the user and ask for the Excel file

print("   Welcome to the Pokemon Card Market Price Checker!")
print("")
print("Please select your Excel file containing your card information.")
print("")
print("")
input("Press Enter to continue...")

root = tk.Tk()
root.withdraw()
excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

# Load the Excel file into a pandas dataframe
try:
    df = pd.read_excel(excel_path)
    print("")
    print(f"* Successfully loaded {os.path.basename(excel_path)}. *")
    print("")
except Exception as e:
    print(f"Error loading {os.path.basename(excel_path)}: {e}")
    exit(1)
if len(df) == 0 :   
    df['Timestamp'] = df['Timestamp'].astype(str)

    # Insert new timestamp value
    index = 0
    df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')





#Binder Info.
total_cards = len(df)

if total_cards != 0:

    total_price = df['Price'].sum(skipna=True)
    total_na_rows = df["Price"].isna().sum()
    print("")
    print("")
    print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
    print("")
    print("")
    if total_na_rows > 0:
        print('Warning Missing Prices !')
        if total_na_rows == 1:
            print(f'There was found {total_na_rows} card with missing price')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing price!')
        else:
            print(f'There were found {total_na_rows} cards with missing prices')
            print("")
            print('It is suggested that you run Option 2 in order to retrieve missing prices!')
        nan_card_names = df.loc[df['Price'].isna(), 'Card Name'].reset_index()
        print(nan_card_names)
else:
    total_price = 0
    print("You don't have any cards yet :(")

input("Press Enter to continue...")



while True:
    print("\nWhat would you like to do today?")
    print("**********************************")
    print("1. Add a new card entry")
    print("2. Update card prices")
    print("3. Add new card entry and update prices")
    print("4. Update only missing card prices")
    print("5. Show Top 10 Most Expensive Cards")
    print("6. Show Total Binder Value")
    print("Enter 'exit' to quit the program")
    print("")
    print("")
    choice = input("Please enter your choice (1/2/3/4/5/6 or exit): ")
    
    if choice == '1':
        main()
        
    elif choice == '2':
        main()
        
    elif choice == '3':
        main()
        
    elif choice == '4':
        main()
        
    elif choice == '5':
        main()
    
    elif choice == '6':
        main()
        
    
    elif choice.lower()=='exit':
        #Exit the program
        print("Exiting program.")
        break
    
    else:
        #invalid choice
        print("Invalid choice. Please pick one of the choices!")





        

if __name__ == '__main__':
    main()
