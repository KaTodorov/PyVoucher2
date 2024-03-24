# PyVoucher 

The code was made with the idea of fast use python for printing a hardcopy stickers-vouchers with a promo code generated from os.urandom().
The project was made only for practical use for a local bussiness, generates 3 exel files in form of tables to store & make the stickers.

Every parameter, including "starting_number","promo code","validity until", "promotion" can be changed.
exept the total number of vouchers(it is 100) and the number of parameters.

## Install

    pip install -r requirements.txt
    python exel.py

It is supposed to generate 3 files: one for printing (starting_number-(starting_number+100)\_print.xlsx),
one raw (starting_number-(starting_number+100)\_raw.xlsx), one for verification(starting number-(starting_number + 100)\_print_verification).
### Printing File Sample
![Alt text](https://github.com/KaTodorov/PyVoucher2/blob/master/Untitled.png?raw=true 'Title')
The file used to print out the stickers and stamp them.
### Verification File Sample
![Alt text](https://github.com/KaTodorov/PyVoucher2/blob/master/Untitled1.png?raw=true 'Title')

The file is printed to check the validity of each vaucher.
### Raw File Sample
![Alt text](https://github.com/KaTodorov/PyVoucher2/blob/master/Untitled3.png?raw=true 'Title')

The raw format for 1 column results.

