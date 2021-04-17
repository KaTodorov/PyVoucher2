# PyVoucher

The code was made with the idea of fast use python for printing a hardcopy stickers-vouchers with a promo code generated from os.urandom().
The project was made only for practical use for a local bussiness.

Every parameter, including "starting_number","promo code","validity until", "promotion" can be changed.
exept the total number of vouchers(it is 100) and the number of parameters.

## Install

    pip install -r requirements.txt
    python exel.py

If any libs are missing, install them. Start the file, it is supposed to generate 3 files: one for printing (starting_number-(starting_number+100)\_print.xlsx),
one raw(starting_number-(starting_number+100)\_raw.xlsx), one for verification(starting number-(starting_number + 100)\_print_verification).

![Alt text](https://github.com/KaTodorov/PyVoucher2/blob/master/Untitled.png?raw=true 'Title')
printing file
![Alt text](relative/path/to/img.jpg?raw=true 'Title')
verification file
![Alt text](relative/path/to/img.jpg?raw=true 'Title')
raw file
