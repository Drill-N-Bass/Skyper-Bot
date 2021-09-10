# Skyper-Bot

### The main functionality of Skyper-Bot is:
- processing input offer list from suppliers from Xls to pandas.
- getting actual currency exchange rate (PLN/EUR) and computing the price in EUR.
- adding proper margin (6% or 7%).
- using SkPy API for communication between code and Skype for sending offers with a proper margin to adequate clients.
- the message mentioned above has a random additional 'hello' message. The message looks human-produced, not bot. 


### Additional features:
- displaying offers in PLN and EUR ( with 6% and 7% margin).
- remembering login in for the next sessions.
- displaying actual PLN/EUR currency exchange rate.
- displaying all currency exchange rates for all currencies.
- sending all versions of offers to people in charge.
- displaying all versions of offers .to people in charge.

#### Additional informations:
- for using this software you need to provide a skype nick and password
- the `.xls` file in location `" Offer List, Clients, logins/Clients names and skype adreses.xlsx includes test accounts that can receive messages"`.
- the input data are located in `"Offer List, Clients, logins/the offer list from soure.xlsx"`.
- the `"Offer List, Clients, logins/skype_login.txt"`secure the login name (if this feature was chosen in GUI).
- the software will be called **Skyper Bot**, but because it's still in the process of developing, it has a draft name in few places (like an executable file) "give_me_a_euro_offer_list_with_gui_test_version".
