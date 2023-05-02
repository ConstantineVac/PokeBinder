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
import re
from decimal import Decimal
from pokemontcgsdk import Card
from pokemontcgsdk import Set
from pokemontcgsdk import Type
from pokemontcgsdk import Supertype
from pokemontcgsdk import Subtype
from pokemontcgsdk import Rarity
from pokemontcgsdk import RestClient

RestClient.configure('Your_API_Key')

global df
global total_cards
global total_price

df = None
total_cards = 0
total_price = 0


        

    

#Function1 : Add Cards
# Option 1: Add new card entry
#ask number of the same car
def add_cards():
    global df
    global total_cards
    set_id = None
       
    # Get card info from user
    while True:
        try:
            num_cards = int(input("How many duplicates would you like to add? "))
            break
        except ValueError:
            print("Please enter a valid number.")
    
    while True:
        card_name = input("Please enter the card name (as written on the card): ").strip().replace('_V',' V')
        if all(c.isalpha() or c.isspace() for c in card_name):
            break
        else:
            print("Please enter a valid card name containing only letters and spaces.")

    while True:
        set_name = input("Please enter the expansion set(ex.Lost Origin): ").strip()
        if all(s.isalpha() or s.isspace() for s in set_name ):
            set_name = set_name
            print(set_name)
            break
            
    while True:
        card_number = input("Please enter the card number (up to 3 digits or TG01/GG20): ")
        if re.match("^\d{1,3}$|^[A-Z]{2}\d{2}$|^[A-Z]{2}\d{3}$", card_number):
            break
        else:
            print("Please enter a valid card number containing up to 3 digits or a 2-letter code followed by 2 or 3 digits.")
                          
    print("Fetching Card Info from your INPUTS. Please wait...")        
    for index, row in df.iterrows():
                card_name = f'"{row["Card Name"]}"'
                set_number = row["Set Number"]
                set_name = f'"{row["Expansion Set"]}"'
                query = f'name:{card_name} number:{set_number} set.name:{set_name}'
                results = Card.where(q=query)

                # Update the 'Price' and 'Timestamp' values in the DataFrame
                if results:
                    set_id = results[0].set.ptcgoCode
                    #rarity = results[0].rarity
                    df.at[index, 'Set ID'] = set_id
                    #df.at[index, 'Rarity'] = rarity
                    
    # Add new row to DataFrame
    for i in range(num_cards):
        new_rows = [{'Card Name': card_name.replace('"', ''), 'Expansion Set': set_name.replace('"', ''), 'Set ID': set_id.replace('"', ''), 'Set Number': card_number}]
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True,)
        

    # Save changes to Excel file
    df.to_excel(excel_path, index=False)
    print("\nNew card entry has been added to the Excel file!")
        
    #Binder Status Info after the new entries
    total_cards = len(df)
    total_na_rows = df["Price"].isna().sum()
    now = datetime.now()
    last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
    
    if not math.isnan(total_price) and total_price > 0:
        print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is {total_price.round(2)}€')
    else:
        print(f'Your Binder has in total {total_cards} cards. Run 2 to get an updated total value !')
        
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




#Function 2 Update Card Prices 
#Update Cards Prices using TCG Player AVG Price.
def update_card_prices():
    global df
    global total_cards
    try:    
        print("\nRetrieving Card Prices based on TCG Player API.")    
        print("Updating Prices:") 
        with tqdm(total=total_cards) as pbar:
            for index, row in df.iterrows():
                card_name = f'"{row["Card Name"]}"'
                set_number = row["Set Number"]
                set_name = f'"{row["Expansion Set"]}"'
                query = f'name:{card_name} number:{set_number} set.name:{set_name}'
                results = Card.where(q=query)

                # Update the 'Price' and 'Timestamp' values in the DataFrame
                if results:
                    avg_price = results[0].cardmarket.prices.averageSellPrice
                    df.at[index, 'Price'] = avg_price
                df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                pbar.update(1)
        
        # Save the updated data back to the Excel file
        df.to_excel(excel_path, index=False)
        print("\nPrices updated successfully!")
    
        total_cards = len(df)
        total_na_rows = df["Price"].isna().sum()
        now = datetime.now()
        last_checked = now.strftime("%Y-%m-%d %H:%M:%S")
        total_price = df['Price'].sum(skipna=True) 
        total_price = Decimal(total_price).quantize(Decimal('.01'))
        print(f'Your Binder has in total {total_cards} cards and its total value now is {total_price}€')
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



#3 Update only cards with missing card values!
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
            card_name = f'"{row["Card Name"]}"'
            set_number = row["Set Number"]
            set_name = f'"{row["Expansion Set"]}"'
            query = f'name:{card_name} number:{set_number} set.name:{set_name}'
            results = Card.where(q=query)

            # Update the 'Price' and 'Timestamp' values in the DataFrame
            if results:
                avg_price = results[0].cardmarket.prices.averageSellPrice
                df.at[index, 'Price'] = avg_price
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
    total_price = Decimal(total_price).quantize(Decimal('.01'))
    print(f'Your Binder has in total {total_cards} cards and its total value now is ${total_price}')
        
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

        
   
#Function 4
# Display top 10 most expensive cards

def show_top_10_expensive_cards():
    global df
    #top = df.sort_values(by='Price', ascending=False).head(10)
    top_10 = top_10 = df.nlargest(10, 'Price')[['Card Name', 'Price']]
    print('')
    print('Your Top 10 Most Expensive Cards are:')
    print(top_10)
    pass


#Function 5
# Calculate total binder value

def calculate_binder_price():
    
    df = pd.read_excel(excel_path)
    total_price = df['Price'].sum(skipna=True)
    print('')
    print(f'The Total Value of your binder is ${total_price.round(2)}') 
    pass

#Function 6
# Update Missing Rarity
def update_rarity():
    total_cards = len(df)
    rows_missing_rarity = df[df['Rarity'].isna()]
    mis_count = df['Rarity'].isna().sum()
    print(f'There are {mis_count} card/s missing rarity/ies')
    print("* Do you want to update 'All' or 'Missing' Rarities only ?")
    while True:
        choice7 = input("Update 'all' or 'missing' or 'exit' :").lower()
    
        if choice7 == 'all':
            print("Updating Card Rarities:") 
            with tqdm(total=total_cards) as pbar:       
                for index, row in df.iterrows():
                    card_name = f'"{row["Card Name"]}"'
                    set_number = row["Set Number"]
                    set_name = f'"{row["Expansion Set"]}"'
                    query = f'name:{card_name} number:{set_number} set.name:{set_name}'
                    results = Card.where(q=query)
                                                     
                    # Update Rarities
                    if results:
                        rarity = results[0].rarity
                        df.at[index, 'Rarity'] = rarity
                                                     
                    pbar.update(1)                     
            df.to_excel(excel_path, index=False)
            print("\nRarities updated successfully!")
            print("")
            break

        elif choice7 == 'missing' and mis_count == 0:
            print('')
            print('All cards have Rarities, If necessary update all')
            
        elif choice7 == 'missing':
            print("Updating Missing Card Rarities:") 
            with tqdm(total=mis_count) as pbar:
                for index,row in rows_missing_rarity.iterrows():
                    card_name = f'"{row["Card Name"]}"'
                    set_number = row["Set Number"]
                    set_name = f'"{row["Expansion Set"]}"'
                    query = f'name:{card_name} number:{set_number} set.name:{set_name}'
                    results = Card.where(q=query)
                    # Update Rarities
                    if results:
                        rarity = results[0].rarity
                        df.at[index, 'Rarity'] = rarity
                        pbar.update(1)    
            df.to_excel(excel_path, index=False)
            print("\nRarities updated successfully!")
            print("")
            break
            
        elif choice7 == 'exit':
            print('Aborting Rarity Update...')
            break
        
        else:    
            print("Type either 'all' or 'missing'!")
    pass


# Function 7
# Rarities Summary
def rarity_sum():

    df = pd.read_excel(excel_path)
    # Group the DataFrame by the 'Rarity' column and count the number of cards for each rarity
    rarity_counts = df.groupby('Rarity')['Card Name'].count()
    
    # Print the rarity counts
    print(f"Rarity Counts of {total_cards} cards:")
    for rarity, count in rarity_counts.items():
        print(f"{rarity}: {count}")
    pass

now = datetime.now()
custom_fig = Figlet(font='cybermedium')
print("*********************************************************")
print("*                                                       *")
print("*                  Thank you for using                  *")
print("*                                                       *")
print(custom_fig.renderText('  PokeBinder'))              
print("*                                                       *")
print("*                         v1.4                          *")
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
        update_missing()
        df=pd.read_excel(excel_path)

    elif choice == '4':
        show_top_10_expensive_cards()
        df = pd.read_excel(excel_path)
                           
    elif choice == '5':
        calculate_binder_price()
        df = pd.read_excel(excel_path)
        
    elif choice == '6':
        update_rarity()
        df = pd.read_excel(excel_path)
        
    elif choice == '7':
        rarity_sum()
        df = pd.read_excel(excel_path)



# Greet the user and ask for the Excel file

print("   Welcome to Pokemon Binder. Your Pokemon TCG Price Tool!")
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
    # check if the DataFrame has the required columns
    required_columns = ['Card Name', 'Expansion Set', 'Set ID', 'Card Version', 'Rarity', 'Set Number', 'Price', 'Timestamp', 'URL']
    if not all(column in df.columns for column in required_columns):
    # if any required column is missing, initialize a new DataFrame with these columns
        df_new = pd.DataFrame(columns=required_columns)
    # concatenate the new DataFrame with the existing DataFrame
        df = pd.concat([df_new, df], ignore_index=True)
    
    print("")
    print(f"* Successfully loaded {os.path.basename(excel_path)}. *")
    print("")
except Exception as e:
    print(f"Error loading {os.path.basename(excel_path)}: {e}")
    exit(1)
if len(df) == 0 :   
    df['Timestamp'] = df['Timestamp'].astype(str)

    # Insert new timestamp value
    #index = 0
    #df.at[index, 'Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')





#Binder Info.
total_cards = len(df)
total_price = df['Price'].sum(skipna=True)
if total_cards != 0 and total_price >0:

    
    total_na_rows = df["Price"].isna().sum()
    print("")
    print("")
    print(f'Your Binder has in total {total_cards} cards and its last time total value, since last update is ${total_price.round(2)}')
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

print("")    
print("")

input("Press Enter to continue...")



while True:
    print("\nWhat would you like to do today?")
    print("**********************************")
    print("1. Add NEW cards")
    print("2. Update ALL card prices")
    print("3. Update ONLY missing card prices")
    print("4. Show Top10 Most Expensive Cards")
    print("5. Show Total Binder Value")
    print("6. Update Rarities")
    print("7. Count Rarities")
    print("Enter 'exit' to quit the program")
    print("")
    print("")
    choice = input("Please enter your choice (1/.../8 or exit):")
    
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

    elif choice == '7':
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
