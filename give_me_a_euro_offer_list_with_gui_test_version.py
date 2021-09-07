#!/usr/bin/env python
# coding: utf-8

# # About "Give me a euro offer list"
# 
# This simple program will be used as a tool for showing the actual prices from the offer list. 
# Because the offer list is in `PLN` (Polish Złoty) currency and some of my clients need to get my offer list in `EUR` (Euro), I decided to build this tool. 
# 
# The functionalities:
# 1. The program will read `xls` file that I've been using on daily basis for my work as a head of sales in the gaming wholesale industry.
# 2. Then convert it to `Pandas` version.
# 3. It will web-scrap actual currency exchange rate `PLN/EUR` from [Kantor Cent](https://www.centkantor.pl/).
# 4. It will convert `PLN` price and give a complete offer list that I can use to offer my clients.\
# 5. Then, It will allow me to directly send these offers using SkPy API, to Skype. So the whole process of creating and sending offers to skype will be automated. 

# About `Submit offers` button (and module) in part 1087 to 1333:
# Api client will work properly only when clients with the check button will have login names.
# First two clients called: client 1, client 2 - need to have not only login name, but skype group login name.
# The skype group login name has to be this. When placed skype login name(not the group one), the process will stop.
# If there is no login name linked to the particular client don't tick the box of this client name.
# This rule doesn't apply to the "other clients" tick-box. There, there's only a need for one login.





# The first two steps will be done below:

# =============================================================================
# Part 1:
# Processing input data and web scrapping
# =============================================================================

import pandas as pd
# Needed for display log with the error exeption function:
# https://realpython.com/the-most-diabolical-python-antipattern/
import logging


#import sys
#import win32com.client 

"""
If we will open the file via shortcut, then we need to use win32com.client.
Without it, the program stops with an error.
More about (Answer -> nick: SaoPauloooo):
https://stackoverflow.com/questions/397125/reading-the-target-of-a-lnk-file-in-python
"""

#shell = win32com.client.Dispatch("WScript.Shell")
#shortcut = shell.CreateShortCut("bazowy cennik na psn dla wszystkich klientów 06.2021.xlsx — skrót .lnk")
#excel_file = shortcut.Targetpath

# direct opening from the same folder where this code is saved:
df = pd.read_excel("Offer List, Clients, logins/the offer list from soure.xlsx")

"""
Because I need to see entire dataframe for locating certain segments
I will use method that will allows me to see all
https://www.kite.com/python/answers/how-to-print-an-entire-pandas-dataframe-in-python
"""

pd.set_option("display.max_rows", None, "display.max_columns", None)

df.loc[18,'indeks'] = 1
#print(df.loc[18,'indeks']) # Test

"""

This part is hashed.
The idea was to update the core .xlsx file with the new currency exchange rate.
BUT, when I do so, then I lost all formats and formulas in my file.
Because of that, I decided to only display output without saving.
If someone would like to save this file as I described here - with all the consequences,
just unhash the line below:

"""
# df.to_excel(excel_file, index = False) # Saving the file by using a shortcut from this file


# ## Python web scraping
# 
# For this part, I will use two libraries: `requests` and `BeautifulSoup`. The goal here is to download the currency exchange rates for our Pandas data frame. This data is needed for offer list preparation. It will allow me to generate an up-to-date offer list for my clients. As I mentioned before, I will use [Kantor Cent](https://www.centkantor.pl/kursy-walut) page. It's a page with very cheap prices for exchange and it constantly used by me on the daily basis.  

from bs4 import BeautifulSoup
import requests 

url = "https://www.centkantor.pl/kursy-walut"
page = requests.get(url)

soup = BeautifulSoup(page.content, "html.parser")

tbody = soup.find('tbody')
trs = tbody.find_all('tr')
a = []
for tr in trs:
    tds = tr.find_all('td')
    for td in tds:
        a.append(td.text)
        
df_currency = pd.DataFrame(a)
#df_currency ## Test


# Now, after I gather all data in pandas dataframe, the next thing to do is:
# - Create proper colums for each type of data. 
# - Remove "NaN"

df_currency.columns = ["Country"]

df_currency["curency name"] = df_currency.iloc[range(1,344,5)]

df_currency["broker buy"] = df_currency.iloc[range(3,344,5)]["Country"]
df_currency["broker sell"] = df_currency.iloc[range(4,344,5)]["Country"]
df_currency["Country"] = df_currency.iloc[range(2,344,5)]["Country"]
# print(df_currency) # Before cleaning
# # After cleaning:
currency_table = df_currency.apply(lambda x: pd.Series(x.dropna().values)) # https://stackoverflow.com/questions/43119503/how-to-remove-blanks-nas-from-dataframe-and-shift-the-values-up
#currency_table # test


# Before I will fill all necessairy values into final offer list, I need to create a series that will incloude all products that will be offered:

products = df.loc[16:26]['Produkt']
#products # Test


# There is a need to have all the date with reset index. Let's `use .reset_index(drop=True)` with all the data that will be processed below. With this approach there won't be any `NaN` values.

products.reset_index(drop=True, inplace=True)
#products # Test


# Because I have to add some profit to input offer list, I will build a dataframe that will include these, plus polish prices. The goal is to have:
# - Builded the full offer list with 6% income for some clients, with the prices converted from PLN to EUR.
# - Builded the full offer list with 7% income for some clients, with the prices converted from PLN to EUR.
# - Builded two full offer lists with same thing as above but in Polish currency: PLN
# "sprzedaż 6% marży PLN" means "sell with 6% marigin PLN"


pln_006_no_psn_100 = df.loc[0:9]['sprzedaż 6% marży PLN']
pln_006_psn_100 = df.loc[12]['sprzedaż 6% marży PLN']
pln_006 = (
    pln_006_no_psn_100
    .append(pd.Series(pln_006_psn_100))
    .reset_index(drop=True)
)

#print(pln_006) # Test: print polish PLN prices with 6% profit

pln_007_no_psn_100 = df.loc[0:9]['sprzedaż 7% marży PLN']
pln_007_psn_100 = df.loc[12]['sprzedaż 7% marży PLN']
pln_007 = (
    pln_007_no_psn_100
    .append(pd.Series(pln_007_psn_100))
    .reset_index(drop=True)
)

#print(pln_007) # Test: print polish PLN prices with 6% profit

# Computing exchange rate from PLN to EUR:

eur_006_no_psn_100 = (
    df.loc[0:9]['sprzedaż 6% marży PLN'] 
    / float(currency_table.loc[0]["broker buy"])
)# print Euro EUR prices with 6% profit

eur_006_psn_100 = (
    df.loc[12]['sprzedaż 6% marży PLN'] 
    / float(currency_table.loc[0]["broker buy"])
)# print Euro EUR prices with 6% profit (a sencond part)

# type(eur_006_psn_100)
eur_006 = (
    eur_006_no_psn_100
    .append(pd.Series(eur_006_psn_100))
    .reset_index(drop=True)
)


eur_007_no_psn_100 = (
    df.loc[0:9]['sprzedaż 7% marży PLN'] 
    / float(currency_table.loc[0]["broker buy"])
)# print Euro EUR prices with 6% profit

eur_007_psn_100 = (
    df.loc[12]['sprzedaż 7% marży PLN'] 
    / float(currency_table.loc[0]["broker buy"])
)# print Euro EUR prices with 6% profit (a sencond part)

# type(eur_006_psn_100)
eur_007 = (
    eur_007_no_psn_100
    .append(pd.Series(eur_007_psn_100))
    .reset_index(drop=True)
)

#print(eur_006) # Test
#print('\n')
#print(eur_007) # Test


# Builiding dataframe for each output:

"""
Additionally I will round each value with decimals = 2,
because this is adequate with both currencies: EUR and PLN:
"""

frame = {
    'Produkty': products,
    'sprzedaż 6% marży (EUR)':eur_006.astype(float).round(decimals=2),
    'sprzedaż 7% marży (EUR)':eur_007.astype(float).round(decimals=2),
    'sprzedaż 6% marży (PLN)':pln_006.astype(float).round(decimals=2),
    'sprzedaż 7% marży (PLN)':pln_007.astype(float).round(decimals=2)
} 

total_offer_list = pd.DataFrame(data=frame)

#total_offer_list


# To make this dataframe easier to read and use, I will add the name of the currency after each value:

total_offer_list['sprzedaż 6% marży (EUR)'] = (
    total_offer_list['sprzedaż 6% marży (EUR)']
    .apply(lambda x: str(x) + " EUR")
)

total_offer_list['sprzedaż 7% marży (EUR)'] = (
    total_offer_list['sprzedaż 7% marży (EUR)']
    .apply(lambda x: str(x) + " EUR")
)

total_offer_list['sprzedaż 6% marży (PLN)'] = (
    total_offer_list['sprzedaż 6% marży (PLN)']
    .apply(lambda x: str(x) + " PLN")
)

total_offer_list['sprzedaż 7% marży (PLN)'] = (
    total_offer_list['sprzedaż 7% marży (PLN)']
    .apply(lambda x: str(x) + " PLN")
)


#
# =============================================================================

# Part 2:
# Creating GUI

# =============================================================================
#


import tkinter as tk
import tkinter.font as font

#global root
root = tk.Tk()
root.geometry("1770x870") # "1740x850"
root.columnconfigure(0, weight=3)
root.rowconfigure(0, weight=3)
# https://stackoverflow.com/questions/51591456/can-i-use-rgb-in-tkinter/51592104
def from_rgb(rgb):
    """translates an rgb tuple of int to a tkinter friendly color code
    """
    return "#%02x%02x%02x" % rgb   

root.title('Give me a euro offer list')

content = tk.Frame(root)
content.configure(bg='LightCyan2')


content.columnconfigure(0, weight=8)
content.columnconfigure(1, weight=8)
content.columnconfigure(2, weight=8)
content.columnconfigure(3, weight=8)
content.columnconfigure(4, weight=8)

content.rowconfigure(0, weight=8)
content.rowconfigure(1, weight=8)
content.rowconfigure(2, weight=8)
content.rowconfigure(3, weight=8)
content.rowconfigure(4, weight=8)
content.rowconfigure(5, weight=8)
content.rowconfigure(6, weight=8)
content.rowconfigure(7, weight=8)
content.rowconfigure(8, weight=8)
content.rowconfigure(9, weight=8)

# =============================================================================
# Button's font pattern:
# =============================================================================
font_button_big = font.Font(family='Helvetica', size="11", weight='bold')
font_button_small = font.Font(family='Helvetica', size="10", weight='bold')

# =============================================================================
# making "Currency exchange rate ALL" button
# =============================================================================
"""
I will change the data frame into a string,
then separate each line and I will add an interline for better visibility.
"""
string_currency_table = currency_table.to_string()

#print(string_currency_table) # Test
output_currency_table = ""
for line in string_currency_table.splitlines(): #https://stackoverflow.com/questions/15422144/how-to-read-a-long-multiline-string-line-by-line-in-python
#    # Test:
#    print(line+"\n"+"=====================================================================================")
    output_currency_table += line + "\n" + "============================================================================" + '\n'

def show_text_from_entry_all_curr(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    output_text = text_box.insert(
            "0.0", output_currency_table)
    print(output_text)
    
    return None

button_all_curr = tk.Button(
    content,
    command=show_text_from_entry_all_curr,
    text="PLN - all currency ",
    width=20,
    height=2,
    bg=from_rgb((23,175,231)),
    fg="black",
    font=font_button_small
    )

string_currency_table = currency_table.to_string()

print(string_currency_table)

for line in string_currency_table.splitlines(): #https://stackoverflow.com/questions/15422144/how-to-read-a-long-multiline-string-line-by-line-in-python
    print(line+"\n"+"=====================================================================================")

len(currency_table)

# =============================================================================
# making "Currency exchange rate PLN/EUR" button
# =============================================================================
def show_text_from_entry(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    output_text = text_box.insert(
            "0.0",
            "Currency exchange rate PLN/EUR:\n\n" +
            "Broker buy: " +
            str(float(currency_table.loc[0]["broker buy"])) +
            " PLN" +
            "\nBroker Sell: " +
            str(float(currency_table.loc[0]["broker sell"])) +
            " PLN"
            )
    print(output_text)
    return None

button_pln_eur = tk.Button(
    content,
    command=show_text_from_entry,
    text="PLN to EUR",
    width=20,
    height=2,
    bg=from_rgb((23,175,231)),
    fg="black",
    font=font_button_small
    )

# =============================================================================
# making ""Offer: EUR 6% margin" button
# =============================================================================

#string_eur_6 = []

new_list = ""
for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (EUR)"]].values:
    x = repr(x)
    a = x.find("'")
    b = x.rfind("'")
    new_str = x[a+1:b].replace("', '", "  @ ")
    new_list += new_str + '\n'
new_list = "sprzedaż 6% marży (EUR):\n\n\n" + new_list + '\n'
print(new_list)

def show_text_from_entry_eur_6pct(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    new_list = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list += new_str + '\n'
    
    new_list = "\n=============================================================================\nsprzedaż 6% marży (EUR):\n\n\n" + new_list + '\n'
    output_text = text_box.insert("0.0", new_list)
    print(output_text)
    return None

button_eur_6pct = tk.Button(
    content,
    command=show_text_from_entry_eur_6pct,
    text="Offer: EUR 6% margin",
    width=18,
    height=2,
    bg="lawn green",
    fg="black",
    font=font_button_big
    )


# =============================================================================
# making "Offer: EUR 7% margin" button
# =============================================================================

def show_text_from_entry_eur_7pct(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    new_list = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list += new_str + '\n'
    
    new_list = "\n=============================================================================\nsprzedaż 7% marży (EUR):\n\n\n" + new_list + '\n'
    output_text = text_box.insert("0.0", new_list)
    print(output_text)
    return None

button_eur_7pct = tk.Button(
    content,
    command=show_text_from_entry_eur_7pct,
    text="Offer: EUR 7% margin",
    width=18,
    height=2,
    bg="lawn green",
    fg="black",
    font=font_button_big
    )


# =============================================================================
# making ""Offer: PLN 6% margin" button
# =============================================================================

def show_text_from_entry_pln_6pct(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    new_list = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list += new_str + '\n'
    
    new_list = "\n=============================================================================\nsprzedaż 6% marży (PLN):\n\n\n" + new_list + '\n'
    output_text = text_box.insert("0.0", new_list)
    print(output_text)
    return None

button_pln_6pct = tk.Button(
    content,
    command=show_text_from_entry_pln_6pct,
    text="Offer: PLN 6% margin",
    width=18,
    height=2,
    bg="lawn green",
    fg="black",
    font=font_button_big
    )


# =============================================================================
# making ""Offer: PLN 7% margin" button
# =============================================================================

def show_text_from_entry_pln_7pct(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
    new_list = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list += new_str + '\n'
    
    new_list = "\n=============================================================================\nsprzedaż 7% marży (PLN):\n\n\n" + new_list + '\n'
    output_text = text_box.insert("0.0", new_list)
    print(output_text)
    return None

button_pln_7pct = tk.Button(
    content,
    command=show_text_from_entry_pln_7pct,
    text="Offer: PLN 7% margin",
    width=18,
    height=2,
    bg="lawn green",
    fg="black",
    font=font_button_big
    )

# =============================================================================
# making "Adrian display all offers" button
# =============================================================================

def show_text_from_entry_adrian(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
        
    new_list_1 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list_1 += new_str + '\n'
    
    new_list_2 = "\n=============================================================================\nsprzedaż 6% marży (EUR):\n\n\n" + new_list_1 + '\n'
    output_text_2 = text_box.insert("0.0", new_list_2)
#    print(output_text_2)    
    
    new_list_3 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list_3 += new_str + '\n'
    
    new_list_4 = "\n=============================================================================\nsprzedaż 7% marży (EUR):\n\n\n" + new_list_3 + '\n'
    output_text_4 = text_box.insert("0.0", new_list_4)
#    print(output_text_4)
    
    new_list_5 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list_5 += new_str + '\n'
    
    new_list_6 = "\n=============================================================================\nsprzedaż 6% marży (PLN):\n\n\n" + new_list_5 + '\n'
    output_text_6 = text_box.insert("0.0", new_list_6)
#    print(output_text_6)
     
    new_list_7 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        new_list_7 += new_str + '\n'
    
    new_list_8 = "\n=============================================================================\nsprzedaż 7% marży (PLN):\n\n\n" + new_list_7 + '\n'
    output_text_8 = text_box.insert("0.0", new_list_8)
#    print(output_text_8)

    eur_pln_curr_output = text_box.insert(
            "0.0",
            "Currency exchange rate PLN/EUR:\n\n" +
            "Broker buy: " +
            str(float(currency_table.loc[0]["broker buy"])) +
            " PLN" +
            "\nBroker Sell: " +
            str(float(currency_table.loc[0]["broker sell"])) +
            " PLN" +
            '\n'
            )
    print("to jest ten output: ", eur_pln_curr_output)   
    total_list = (
            output_text_2 +
            output_text_4 +
            output_text_6 +
            output_text_8 +
            eur_pln_curr_output
            )
    ''.join(total_list)
    
    return total_list

button_adrian_display = tk.Button(
    content,
    command=show_text_from_entry_adrian,
    text="Adrian: display all offers",
    width=20,
    height=3,
    bg=from_rgb((38,199,62)),
    fg="black",
    font=font_button_big
    )

# =============================================================================
# creating a window that will allow a user to place skype login
# =============================================================================
label_entry_skype_login = tk.Label(
        content, text="Skype Name: ",
        width=18,
        bg=from_rgb((193,254,252)),
        font=font_button_big,
        anchor="e"
        )

# This line allows to display text in entry: `entry_skype_login' from txt file: 'Offer List, Clients, logins\skype_login.txt'
skype_login_var=tk.StringVar()

entry_skype_login = tk.Entry(
        content,
        textvariable=skype_login_var,
        bd =5,
        width=40
        )
text_skype_name = entry_skype_login.get()


# https://stackoverflow.com/questions/68807822/tkinter-how-to-detect-checkbutton-when-gui-starts?noredirect=1#comment121631633_68807822        
with open('Offer List, Clients, logins\skype_login.txt') as file:
    saved_name = file.readline()
    
def skype_login_remember_and_quit():
    # get data from entry
    text_skype_name = entry_skype_login.get()
    with open('Offer List, Clients, logins\skype_login.txt', 'w') as f:  # open file in write mode
        f.write(text_skype_name if skype_login_remembervar.get() else '')  # if user wanted to "remember login" then write to file
        # the entered data else write nothing (note that this will overwrite everything in the file)
    content.quit()  # or use `exit()` depending on what you need

# https://stackoverflow.com/questions/68807822/tkinter-how-to-detect-checkbutton-when-gui-starts?noredirect=1#comment121631633_68807822
root.protocol('WM_DELETE_WINDOW', skype_login_remember_and_quit) # this protocol
# is triggered when user presses the "X" button in the top right corner of window
entry_skype_login.insert(0, saved_name)        
    
skype_login_remembervar = tk.BooleanVar(value=True)

skype_bt_login_remeber = tk.Checkbutton(
        content,
        text="remember login",
        variable=skype_login_remembervar,
        onvalue=True,
        bg=from_rgb((193,254,252)),
        font=font_button_big,
        )

# =============================================================================
# creating a window that will allow a user to place skype pass
# =============================================================================

label_entry_skype_pass = tk.Label(
        content, text="Skype pass: ",
        width=18,
        bg=from_rgb((193,254,252)),
        font=font_button_big,
        anchor="e"
        )

# https://stackoverflow.com/questions/10989819/hiding-password-entry-input-in-python:
entry_skype_pass = tk.Entry(content, bd =5, width=40, show="*")

text_skype_pass = entry_skype_pass.get()


# =============================================================================
# Send all offers to Adrian:
# =============================================================================
"""
The code below seems to be duplicated, but it's not. 
I needed to create an extra version that will be properly prepared
for the Skype input message. This is message dedicated to my boss.
"""
def show_text_from_entry_adrian(): # https://www.youtube.com/watch?v=ITaDE9LLEDY
     
    output_text_1 = (
        "Currency exchange rate PLN/EUR:\n\n" +
        "Broker buy: " +
        str(float(currency_table.loc[0]["broker buy"])) +
        " PLN" +
        "\nBroker Sell: " +
        str(float(currency_table.loc[0]["broker sell"])) +
        " PLN" +
        '\n'
        )
    
    output_text_2 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        output_text_2 += new_str + '\n'
    
    output_text_3 = "\n=============================================================================\nsprzedaż 6% marży (EUR):\n\n\n" + output_text_2 + '\n'  
    
    output_text_4 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (EUR)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        output_text_4 += new_str + '\n'
    
    output_text_5 = "\n=============================================================================\nsprzedaż 7% marży (EUR):\n\n\n" + output_text_4 + '\n'

    
    output_text_6 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        output_text_6 += new_str + '\n'
    
    output_text_7 = "\n=============================================================================\nsprzedaż 6% marży (PLN):\n\n\n" + output_text_6 + '\n'
   
    output_text_8 = ""
    for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (PLN)"]].values:
        x = repr(x)
        a = x.find("'")
        b = x.rfind("'")
        new_str = x[a+1:b].replace("', '", "  @ ")
        output_text_8 += new_str + '\n'
    
    output_text_9 = "\n=============================================================================\nsprzedaż 7% marży (PLN):\n\n\n" + output_text_8 + '\n'
    
    output_text_all = (
            output_text_1 +
            output_text_3 +
            output_text_5 +
            output_text_7 +
            output_text_9
            )
    ''.join(output_text_all)
    output_text_all = str(output_text_all)
    # https://github.com/Terrance/SkPy/issues/173
    import re
    re.sub(r"\\n", "<br/>", output_text_all) # skype needs html line breaks: "<br/>"
    return output_text_all
    

def offers_via_skype_to_adrian():
    from skpy import Skype
    sk = Skype(entry_skype_login.get(), entry_skype_pass.get()) # connect to Skypesk.user 
    # Single chat
    ch = sk.contacts["live:.cid.addcb0ebdebb491e"].chat # 1-to-1 conversationch.sendMsg(content) # plain-text message #
    # https://github.com/Terrance/SkPy/issues/173
    msg = ch.sendMsg(
            show_text_from_entry_adrian(),
            rich=True) # skype needs html line breaks: "rich=True"
    msg
    


button_adrian_offers_send = tk.Button(
    content,
    text="Adrian: send all offers",
    command=offers_via_skype_to_adrian,
    width=20,
    height=3,
    bg=from_rgb((38,199,62)),
    fg="black",
    font=font_button_big
    )

# =============================================================================
# delete button: It will delete the output from `text_box`
# =============================================================================

def delete_output():
    text_box.delete("0.0", tk.END)
    return None
delete_button = tk.Button(
    content,
    command=delete_output,
    text="Delete output text",
    width=20,
    height=3,
    bg="red",
    fg="black",
    font=font_button_big
    )

# =============================================================================
# making "exit window" button
# =============================================================================

exit_button = tk.Button(
        content, 
        text='Quit', 
        command=skype_login_remember_and_quit,
        width=30,
        height=3,
        bg=from_rgb((213,60,60)),
        fg="black",
        font=font_button_big
        )
# =============================================================================
# creating a text box for output:
# =============================================================================

text_box_description = tk.Label(
        content,
        text="Output:",
        foreground="forest green",
        background=from_rgb((30,81,100)),
        width=70,
        height=1,
        font=font.Font(family='Helvetica', size="12", weight='bold')
        )

text_box = tk.Text(
        content,
        width=80,
        height=40,
        bg=from_rgb((38,15,2)),
        fg="gainsboro"
        )

# =============================================================================
# making image button
# https://blog.furas.pl/python-tkinter-how-to-load-display-and-replace-image-on-label-button-or-canvas-gb.html
# https://pythonexamples.org/python-pillow-show-display-image/
# https://pythonexamples.org/python-pillow-resize-image/
# =============================================================================
#from PIL import ImageTk, Image
# https://stackoverflow.com/questions/38180388/tkinter-how-to-insert-an-image-to-a-text-widget
# https://stackoverflow.com/questions/35924690/tkinter-image-wont-show-up-in-new-window

def add_image():
    toplevel_about = tk.Toplevel()
#    toplevel_about.columnconfigure(0, weight=2)
    toplevel_about.title('About')
    canvas = tk.Canvas(toplevel_about, width = 1000, height = 650)
#    canvas.grid(sticky="S"+"N"+"E"+"W")
#    canvas.columnconfigure(0, weight=3)
    canvas.pack(expand = tk.YES, fill = tk.BOTH)
    img1 = tk.PhotoImage(file = 'Pictures\picture1.png')
                                #image not visual
    canvas.create_image(50, 10, image = img1, anchor = tk.NW)
    #assigned the gif1 to the canvas object
    canvas.img1 = img1
    canvas.create_text(500,530,fill="darkblue",font="Times 20 italic bold",
                        text="Give me a euro offer list v. 0.1.0\n\n" +
                        "Author: Paweł Pedryc\n pawelpedryc@gmail.com")
    
# canvas + grid: https://stackoverflow.com/questions/20149483/python-canvas-and-grid-tkinter

button_about = tk.Button(
    content,
    command=add_image,
    text="About",
    width=8,
    height=1,
    bg="red",
    fg="black",
    font=font.Font(family='Helvetica', size="14", weight='bold')
    )

def select_all(): # select all `tk.Checkbutton`
    check_buttons_list = [twovar, threevar, fourvar, fivevar, sixvar, sevenvar, eightvar]
    for name in check_buttons_list:
        if onevar.get() == True:
            for name in check_buttons_list:
                name.set(value=True) # instead of `value=True` can be `1`
        if onevar.get() == False:
            pass

    
onevar = tk.BooleanVar(value=False)
twovar = tk.BooleanVar(value=False)
threevar = tk.BooleanVar(value=False)
fourvar = tk.BooleanVar(value=False)
fivevar = tk.BooleanVar(value=False)
sixvar = tk.BooleanVar(value=False)
sevenvar = tk.BooleanVar(value=False)
eightvar = tk.BooleanVar(value=False)


# creating a label frame for checkbuttons:

label_frame_for_checkbuttons = tk.LabelFrame(
        content, text='Send offer to client(s):',
        width = 10,
        height = 65,
        bg=from_rgb((137,212,209)),
        font=font.Font(family='Helvetica', size="10", weight='bold')
        )

label_frame_for_checkbuttons.grid(
        column=2,
        row=6,
        columnspan=3,
        rowspan=8,
        sticky="N"+"E"
        )

one = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Match ALL",
        variable=onevar,
        onvalue=True,
        width = 20,
        bg=from_rgb((255,121,121)),
        fg="black",
        command=select_all,
        font=font.Font(family='Helvetica', size="9", weight='bold')
        )

two = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="NL games",
        variable=twovar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

three = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Czech games",
        variable=threevar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

four = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Russian Players",
        variable=fourvar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

five = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Izrael pc gamer",
        variable=fivevar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

six = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Usa games wholesale",
        variable=sixvar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

seven = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="IGo",
        variable=sevenvar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

eight = tk.Checkbutton(
        label_frame_for_checkbuttons,
        text="Other clients",
        variable=eightvar,
        onvalue=True,
        bg='LightCyan2',
        command=select_all
        )

one.grid(column=2, row=5, columnspan=3, pady=4)
two.grid(column=2, row=6, sticky=(tk.W))
three.grid(column=2, row=7, sticky=(tk.W))
four.grid(column=2, row=8, sticky=(tk.W))
five.grid(column=3, row=6, sticky=(tk.W))
six.grid(column=3, row=7, sticky=(tk.W))
seven.grid(column=3, row=8, sticky=(tk.W))
eight.grid(column=3, row=9, sticky=(tk.W))

# =============================================================================
# Submit offers for clients:
# =============================================================================
import random

"""
The `random` module and `skype_welcome_messages` will allow
to create a random "hello" message for clients, so they will assume that
they are speaking with someone real.
"""

skype_welcome_messages = [
        "Hello, how are you?<br/>I have a new PSN Poland offer list for you! :)<br/>",
        "Hello, I'm having another day with great weather! :D Hope you're having one too!<br/>See below new PSN Poland offer list<br/>",
        "Hello, what's new on your end? :)<br/>A new PSN Poland offer attached below<br/>",
        "Hello, check my new PSN Poland offer list! ;)<br/>",
        "Hi :), new PSN Poland offer list:<br/>",
        "Hi, just sending you the PSN Poland offer list, and I'm off. It's a busy day. I will check in later - if you need anything: let me know ;)",
        "Hi, I'm collecting orders for PSN Poland, do you need anything today? :)<br/>",
        "Hi, rainy day today but I have more PSN Poland requests, so there's profit :P If you need anything from the new list, let me know ;)<br/>"
        ]        


new_list_6_pln = ""
for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (PLN)"]].values:
    x = repr(x)
    a = x.find("'")
    b = x.rfind("'")
    new_str = x[a+1:b].replace("', '", "  @ ")
    new_list_6_pln += new_str + '\n'

new_list_6_pln = "\n" + new_list_6_pln + '\n'


new_list_7_pln = ""
for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (PLN)"]].values:
    x = repr(x)
    a = x.find("'")
    b = x.rfind("'")
    new_str = x[a+1:b].replace("', '", "  @ ")
    new_list_7_pln += new_str + '\n'

new_list_7_pln = "\n" + new_list_7_pln + '\n'


new_list_6_eur = ""
for x in total_offer_list.loc[:][["Produkty","sprzedaż 6% marży (EUR)"]].values:
    x = repr(x)
    a = x.find("'")
    b = x.rfind("'")
    new_str = x[a+1:b].replace("', '", "  @ ")
    new_list_6_eur += new_str + '\n'

new_list_6_eur = random.choice(skype_welcome_messages) + "\n" + new_list_6_eur + '\n'


new_list_7_eur = ""
for x in total_offer_list.loc[:][["Produkty","sprzedaż 7% marży (EUR)"]].values:
    x = repr(x)
    a = x.find("'")
    b = x.rfind("'")
    new_str = x[a+1:b].replace("', '", "  @ ")
    new_list_7_eur += new_str + '\n'

new_list_7_eur = random.choice(skype_welcome_messages) + "\n" + new_list_7_eur + '\n'

#print("new_list_6_eur: ", new_list_6_eur) # Test.
#print("new_list_6_eur type: ", type(new_list_6_eur)) # Test.
#print("new_list_6_eur type repr: ", repr(new_list_6_eur)) # Test.

    
df_clients = pd.read_excel("Offer List, Clients, logins/Clients names and skipe adreses.xlsx", index_col=0) #https://stackoverflow.com/questions/36606931/how-to-set-in-pandas-the-first-column-and-row-as-index/51936996

## Tests:
##print(df_clients)
###df_clients = df_clients.loc[1:7]
#print(df_clients.loc["NL games", "Skype Adress"])
###df_clients.loc[0]
  

# =============================================================================
# Search module for finding skype logins from recent chats. 
# It's the only tool for finding group chats:
# https://skpy.t.allofti.me/guides/chats.html#finding-specific-chats    

                        
#from skpy import Skype
#sk = Skype(Place your login with quotes, place your skype with quotes) # connect to Skypesk.user #
#while True:
#    for chat in sk.chats.recent():
#        print(chat)
#    else: # No more chats returned.
#        break
# =============================================================================



def submit_offers_to_selected_clients():
    
    """ This Part is still in building """
    
    
    """
    Because I need a login and a password to access
    the Skype platform before sending messages to clients,
    I need to confirm that these are proper. If they aren't I need to create
    the error message box
    """
#    # https://stackoverflow.com/questions/47676319/how-to-create-a-tkinter-error-message-box
##    tk.messagebox.showerror("We can’t sign you!", "Your login or password is not correct. Try to fix it.")
#    try:
#        tk.messagebox.showerror("We can’t sign you!", "Your login or password is not correct. Try to fix it.")
#
##    except SkypeApiException:
#    except SkypeApiException:     
#         tk.messagebox.showerror("We can’t sign you!", "Your login or password is not correct. Try to fix it.")
#    """
#    To find exceptions:
#        https://stackoverflow.com/questions/18176602/how-to-get-the-name-of-an-exception-that-was-caught-in-python
#    """
#    
    
    
    
#    import sys    
#    except Exception:
#        exc_type, value, traceback = sys.exc_info()
#        assert exc_type.__name__ == 'NameError'
#        print("Failed with exception [%s]" % exc_type.__name__)
    
    """ END OF BUILDING PART """
    
    
    from skpy import Skype
    sk = Skype(entry_skype_login.get(), entry_skype_pass.get()) # connect to Skypesk.user 
    login_from_clients_dict = {}
        
    check_buttons_dict = {
            "NL games":[twovar, new_list_6_eur, new_list_6_pln],
            "Czech games":[threevar, new_list_7_eur],
            "Russian Players":[fourvar, new_list_6_eur],
            "Izrael pc gamer":[fivevar, new_list_6_eur],
            "Usa games wholesale":[sixvar, new_list_7_eur],
            "IGo":[sevenvar, new_list_7_eur],
            "Other clients":[eightvar, new_list_7_eur]
            }
    check_buttons_list = [twovar, threevar, fourvar, fivevar, sixvar, sevenvar, eightvar]
    
    #print(check_buttons_dict["NL games"])
   
    for name in check_buttons_list:
        if name.get() == True:
            for dict_name in check_buttons_dict:
                if name == check_buttons_dict[dict_name][0]:
                    
                    login_from_clients_dict[dict_name] = df_clients.loc[dict_name, "Skype Adress"]
    
                    # Single chat
    #                    ch = sk.contacts[df_clients.loc[str(dict_name),"Skype Adress"]].chat # 1-to-1 conversationch.sendMsg(content) # plain-text message
                    
                    if dict_name == "Other clients":
                        # https://github.com/Terrance/SkPy/issues/174
    #                    print(repr(login_from_clients_dict[dict_name])) # test 
                        """
                        "Other clients" will get a 7% margin offer. All messages for clients
                        from this category - in the excel file painted blue - will be sent at once. 
                        """
                        for client_login in df_clients.loc["Other clients":, "Skype Adress"].dropna():
                            
                            # Single chat
                            ch = sk.contacts[client_login].chat
                            # https://github.com/Terrance/SkPy/issues/173
                            msg = ch.sendMsg(
                                    check_buttons_dict[dict_name][1],
                                    rich=True
                                    ) # skype needs html line breaks: "rich=True"
                            msg
                    
                    
                                
                        for client_login in df_clients.loc["Other clients":, "group skype adress"].dropna():        
                            try:
                                # Single chats:
                                ch1 = sk.contacts[client_login].chat
                                msg = ch1.sendMsg(
                                        check_buttons_dict[dict_name][1],
                                        rich=True
                                        ) # skype needs html line breaks: "rich=True"
                                msg 
                                
                            except AttributeError as ae:
                                logging.exception("Caught an error [in code used: except AttributeError]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell." + str(ae))
                                print("Caught an error [in code used: except Exception]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell.")
                            
                            try:
                                # Group chats:
                                ch2 = sk.chats[client_login]
                                msg = ch2.sendMsg(
                                        check_buttons_dict[dict_name][1],
                                        rich=True
                                        ) # skype needs html line breaks: "rich=True"
                                msg      
                            except AttributeError as ae:
                                logging.exception("Caught an error [in code used: except AttributeError]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell." + str(ae))
                                print("Caught an error [in code used: except Exception]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell.")

                       
                    """
                    If there is a situation where the client has his own Skype account
                    and he/she participate in group chat, then I need to generate two 
                    messages: for both accounts:
                    """
                    if (dict_name != "NL games" and 
                        dict_name == "Czech games" and
                        dict_name != "Other clients" and
                        df_clients.loc[dict_name, "group skype adress"] != "NaN"
                        ):
                        ch = sk.contacts[login_from_clients_dict[dict_name]].chat
                        msg = ch.sendMsg(
                                check_buttons_dict[dict_name][1],
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg

                        # Group chats:
                        ch1 = sk.chats[df_clients.loc[dict_name, "group skype adress"]]
                        msg = ch1.sendMsg(
                                check_buttons_dict[dict_name][1],
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg
                        
                    elif (dict_name != "NL games" and 
                        dict_name != "Czech games" and
                        dict_name != "Other clients" and
                        df_clients.loc[dict_name, "group skype adress"] != "NaN"
                        ):
                        
                        ch = sk.contacts[login_from_clients_dict[dict_name]].chat
                        msg = ch.sendMsg(
                                check_buttons_dict[dict_name][1],
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg
                        
                        try:
                            # Single chats:
                            ch1 = sk.contacts[df_clients.loc[dict_name, "group skype adress"]].chat
                            msg = ch1.sendMsg(
                                    check_buttons_dict[dict_name][1],
                                    rich=True
                                    ) # skype needs html line breaks: "rich=True"
                            msg 
                            
                        except Exception as e:
                            logging.exception("Caught an error [in code used: except AttributeError]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell." + str(e))
                            print("Caught an error [in code used: except Exception]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell.")
                        
                        try:
                            # Group chats:
                            ch2 = sk.chats[df_clients.loc[dict_name, "group skype adress"]]
                            msg = ch2.sendMsg(
                                    check_buttons_dict[dict_name][1],
                                    rich=True
                                    ) # skype needs html line breaks: "rich=True"
                            msg      
                        except Exception as e:
                            logging.exception("Caught an error [in code used: except AttributeError]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell." + str(e))
                            print("Caught an error [in code used: except Exception]: 'NoneType' object has no attribute 'chat'. There is no skype login (or a proper one) in current excel cell.")
                       
                            
                    """
                    "NL games" client needs to get a specific offer list
                    """
                    if (dict_name == "NL games" and 
                        dict_name != "Other clients" and
                        df_clients.loc[dict_name, "group skype adress"] != "NaN"
                        ):
                        """
                        Because this client need to get two offers 
                        and the output will be list-type, we need to convert it to string:
                        """
                        ch = sk.contacts[login_from_clients_dict[dict_name]].chat
                        offer_string = ' '.join([str(elem) for elem in check_buttons_dict[dict_name][1:]])
                        msg = ch.sendMsg(
                                offer_string,
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg
                        
                        # Group chats:
                        ch1 = sk.chats[df_clients.loc[dict_name, "group skype adress"]]
                        offer_string = ' '.join([str(elem) for elem in check_buttons_dict[dict_name][1:]])
                        msg = ch1.sendMsg( 
                                offer_string,
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg
                        
                       
                        # Tests:
#                        print("regular print  ", check_buttons_dict[dict_name][1:])
#                        print("repr print ", repr(check_buttons_dict[dict_name][1:2]))
#                        print("repr print type ", type(repr(check_buttons_dict[dict_name][1:2])))
                    """
                    All clients who want to get this offer, but are less important.
                    The offer will be 7% margin.
                    """
                    if (dict_name != "Other clients" and
                        dict_name != "NL games" and
                        dict_name != "Czech games" and
                        df_clients.loc[dict_name, "group skype adress"] == "NaN"
                        ):
                        # https://github.com/Terrance/SkPy/issues/174
    #                    print(repr(login_from_clients_dict[dict_name])) # test 
                        """
                        If the excel file included logins with quotes
                        then we need to add the `replace` function,
                        because we have double quotes like '"skype:login"'
                        """
#                        import time
#                        time.sleep(2)
                        
                        # Single chat
                        ch = sk.contacts[login_from_clients_dict[dict_name]].chat #.replace('"', "")
                        # https://github.com/Terrance/SkPy/issues/173
                        msg = ch.sendMsg(
                                check_buttons_dict[dict_name][1],
                                rich=True
                                ) # skype needs html line breaks: "rich=True"
                        msg
                        


submit_button = tk.Button(
        label_frame_for_checkbuttons,
        text="Submit offers",
        width=40,
        height=2,
        bg=from_rgb((255,24,24)),
        fg="black",
        font=font_button_big,
        command=submit_offers_to_selected_clients
        )

# =============================================================================
# Main grid:
# =============================================================================

content.grid(column=0, row=0)
#frame.grid(column=0, row=0, columnspan=3, rowspan=6)
text_box_description.grid(column=0, row=0)
text_box.grid(column=0, row=1, rowspan=10)
# =============================================================================
# # create a Scrollbar and associate it with `text_box`:
# =============================================================================


# Vertical (y) Scroll Bar
scroll = tk.Scrollbar(content)
# scroll bar remember position after "scroll":
text_box.config(yscrollcommand=scroll.set) 
scroll.config(command=text_box.yview)

scroll.grid(column=0, row=1, rowspan=10, sticky="N"+"S"+"E")
# Configure the scrollbars
scroll.config(command=text_box.yview)

#hello_message.grid(column=3, row=9, )#columnspan=2)
button_all_curr.grid(column=3, row=4)
button_pln_eur.grid(column=2, row=4)
button_eur_6pct.grid(column=1, row=4)
button_eur_7pct.grid(column=1, row=5)
button_pln_6pct.grid(column=1, row=6)
button_pln_7pct.grid(column=1, row=7)

delete_button.grid(column=1, row=8)

button_adrian_offers_send.grid(column=1, row=9)
button_adrian_display.grid(column=1, row=10)

label_entry_skype_login.grid(column=1, row=1,  sticky=(tk.N))
entry_skype_login.grid(column=2, row=1, sticky=(tk.N), rowspan=3)
skype_bt_login_remeber.grid(column=3, row=1,  sticky=(tk.N))
label_entry_skype_pass.grid(column=1, row=2, sticky=(tk.N))
entry_skype_pass.grid(column=2, row=2, sticky=(tk.N), rowspan=3)

button_about.grid(column=4, row=0, sticky=(tk.E))
submit_button.grid(column=2, row=10, columnspan=3, sticky=tk.S, pady=4)
exit_button.grid(column=3, row=10, columnspan=4, sticky=(tk.N))

""" Weight configure doesn't work, don't know why """
# =============================================================================
# Weight configure:
# =============================================================================
#root.grid_columnconfigure(0, weight=1)
#content.rowconfigure(0, weight=1)

# https://stackoverflow.com/questions/53665173/tkinter-columnconfigure-weight-not-adjusting

#root.columnconfigure(0, weight=10)
#root.columnconfigure(1, weight=10)
#root.columnconfigure(2, weight=10)
#root.columnconfigure(3, weight=10)
#root.columnconfigure(4, weight=10)
#
#bttn_list = [label_frame_for_checkbuttons,
#             #content,
#             text_box_description,
#             text_box, 
#             scroll,
#             hello_message,
#             button_all_curr,
#             button_pln_eur,
#             button_eur_6pct,
#             button_eur_7pct,
#             button_pln_6pct,
#             button_pln_7pct,
#             delete_button,
#             button_adrian_offers_send,
#             button_adrian_display,
#             label_entry_skype_login,
#             entry_skype_login,
#             label_entry_skype_pass,
#             entry_skype_pass,
#             button_about,
#             submit_button,
#             exit_button
#             ]
#
#for i in range(len(bttn_list)):
#    root.columnconfigure(i, weight=4) ## Not the button, but the parent
#    root.rowconfigure(i, weight=5) ## Not the button, but the parent


# https://stackoverflow.com/questions/41979656/is-it-possible-to-make-tkinter-remember-variables-when-you-close-it
# save the text after shutdown
with open('Offer List, Clients, logins\skype_login.txt','w') as text_file_skype_login:
    text_file_skype_login.write(text_skype_name)
    

root.mainloop()
# =============================================================================
# Question:
#    I am just starting with Python.
#    When I use Tk, a blank dialog always opens up behind tkMessageBox or whatever other GUI element that one executes.
#    Is their a way of disabling this?
#    
#    Answer:
#    
#    The "blank dialog" is Tkinter's root window.
#    To eliminate that, explicitly create a root and withdraw it before proceeding:
#    
#           root = Tkinter.Tk()
#           root.withdraw()
# =============================================================================
root.withdraw()

text_file_skype_login.close() # Closes the .txt file

