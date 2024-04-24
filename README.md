# Sheets Translator

Sheets Translator is a program for translating your spreadsheet files into a multitude of languages quickly and easily.

## Installation

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install dependency needed.

Run this command.
``` bash
pip install -r requirements.txt
```

## Usage

First you need to fill in the spreadsheet file with which you want to translate:

- input_data_all: to translate a complete file containing all the sheets in your file

 - input_data_one_column : to translate one column at a time (there is no row limit)

After filling in the required file, dn your command terminal write the following command in the folder:

``` bash 
python translate.py 
```
You can choose the translation module you want from the following:

- Google translator
- Deepl
- MyMemory
- Qcri
- Pons
- Yandex
- LibreTranslator

For Deepl, Qcri and Yandex, you will need an api key.

Now we just have to wait for the translation to finish.

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

Sheets Translator © 2024 by Maxence Baffet is licensed under CC BY-NC-SA 4.0. To view a copy of this license, visit https://creativecommons.org/licenses/by-nc-sa/4.0/
