import pandas as pd
from deep_translator import GoogleTranslator, MyMemoryTranslator, DeeplTranslator, QcriTranslator, LingueeTranslator, PonsTranslator, YandexTranslator, LibreTranslator
import time
from tqdm import tqdm
import os

timestr = time.strftime("%Y.%m.%d-%H.%M.%S")

languages = {
    'af': 'afrikaans', 'sq': 'albanian', 'am': 'amharic', 'ar': 'arabic',
    'hy': 'armenian', 'az': 'azerbaijani', 'eu': 'basque', 'be': 'belarusian',
    'bn': 'bengali', 'bs': 'bosnian', 'bg': 'bulgarian', 'ca': 'catalan',
    'ceb': 'cebuano', 'ny': 'chichewa', 'zh-cn': 'chinese (simplified)', 'zh-tw': 'chinese (traditional)',
    'co': 'corsican', 'hr': 'croatian', 'cs': 'czech', 'da': 'danish',
    'nl': 'dutch', 'en': 'english', 'eo': 'esperanto', 'et': 'estonian',
    'tl': 'filipino', 'fi': 'finnish', 'fr': 'french', 'fy': 'frisian',
    'gl': 'galician', 'ka': 'georgian', 'de': 'german', 'el': 'greek',
    'gu': 'gujarati', 'ht': 'haitian creole', 'ha': 'hausa', 'haw': 'hawaiian',
    'iw': 'hebrew', 'he': 'hebrew', 'hi': 'hindi', 'hmn': 'hmong',
    'hu': 'hungarian', 'is': 'icelandic', 'ig': 'igbo', 'id': 'indonesian',
    'ga': 'irish', 'it': 'italian', 'ja': 'japanese', 'jw': 'javanese',
    'kn': 'kannada', 'kk': 'kazakh', 'km': 'khmer', 'ko': 'korean',
    'ku': 'kurdish (kurmanji)', 'ky': 'kyrgyz', 'lo': 'lao', 'la': 'latin',
    'lv': 'latvian', 'lt': 'lithuanian', 'lb': 'luxembourgish', 'mk': 'macedonian',
    'mg': 'malagasy', 'ms': 'malay', 'ml': 'malayalam', 'mt': 'maltese',
    'mi': 'maori', 'mr': 'marathi', 'mn': 'mongolian', 'my': 'myanmar (burmese)',
    'ne': 'nepali', 'no': 'norwegian', 'or': 'odia', 'ps': 'pashto',
    'fa': 'persian', 'pl': 'polish', 'pt': 'portuguese', 'pa': 'punjabi',
    'ro': 'romanian', 'ru': 'russian', 'sm': 'samoan', 'gd': 'scots gaelic',
    'sr': 'serbian', 'st': 'sesotho', 'sn': 'shona', 'sd': 'sindhi',
    'si': 'sinhala', 'sk': 'slovak', 'sl': 'slovenian', 'so': 'somali',
    'es': 'spanish', 'su': 'sundanese', 'sw': 'swahili', 'sv': 'swedish',
    'tg': 'tajik', 'ta': 'tamil', 'te': 'telugu', 'th': 'thai',
    'tr': 'turkish', 'uk': 'ukrainian', 'ur': 'urdu', 'ug': 'uyghur',
    'uz': 'uzbek', 'vi': 'vietnamese', 'cy': 'welsh', 'xh': 'xhosa',
    'yi': 'yiddish', 'yo': 'yoruba', 'zu': 'zulu'
}

api_keys_file = 'api_keys.txt'

def load_api_keys():
    if os.path.exists(api_keys_file):
        with open(api_keys_file, 'r') as file:
            return dict(line.strip().split('=') for line in file if line.strip())
    return {}

def save_api_key(service, key):
    api_keys = load_api_keys()
    api_keys[service] = key
    with open(api_keys_file, 'w') as file:
        for k, v in api_keys.items():
            file.write(f"{k}={v}\n")

def update_api_key_prompt(service):
    current_key = get_api_key(service)
    if current_key:
        update_key = input(f"Current API key found for {service.title()}. Would you like to update it? (yes/no): ").lower()
        if update_key == 'yes':
            new_key = input("Enter the new API key: ")
            save_api_key(service, new_key)
            return new_key
    else:
        new_key = input(f"Enter the API key for {service.title()}: ")
        save_api_key(service, new_key)
        return new_key
    return current_key

def get_api_key(service):
    return load_api_keys().get(service)

def translate_text(text, lang_code, translator):
    try:
        return translator.translate(text, timeout=10)
    except Exception as e:
        print(f"Failed to translate text due to: {e}")
        return None

translator_options = {
    1: GoogleTranslator,
    2: MyMemoryTranslator,
    3: DeeplTranslator,
    4: QcriTranslator,
    5: LingueeTranslator,
    6: PonsTranslator,
    7: YandexTranslator,
    8: LibreTranslator
}

def get_translator(service_number, target_lang, api_key):
    TranslatorClass = translator_options.get(service_number)
    if TranslatorClass is None:
        raise ValueError("Invalid translator selection.")
    if service_number in [3, 4, 7]:
        translator_instance = TranslatorClass(source='auto', target=target_lang, auth_key=api_key)
        if not test_api_key(translator_instance):
            raise ValueError("Invalid API key provided.")
        return translator_instance
    return TranslatorClass(source='auto', target=target_lang)

def test_api_key(translator_instance, sample_text="Hello"):
    try:
        result = translator_instance.translate(sample_text)
        return result is not None
    except Exception as e:
        print(f"Error testing the API key: {e}")
        return False

def translate_column(df, lang_code, translator_instance, column_name='Input Text'):
    failed_indices = []
    df['Translated Text'] = None
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Translating column"):
        original_text = row[column_name]
        if isinstance(original_text, str):
            try:
                translated_text = translate_text(original_text, lang_code, translator_instance)
                df.at[index, 'Translated Text'] = translated_text if translated_text else original_text
            except Exception as e:
                print(f"Failed to translate text in row {index}, column '{column_name}' due to: {e}")
                failed_indices.append(index)
    return df, failed_indices

def translate_all_sheets(xls, lang_code, translator_instance):
    failed_sheets = {}
    translated_dfs = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        failed_indices = []
        for index, row in tqdm(df.iterrows(), total=len(df), desc=f"Translating {sheet_name}"):
            for col in df.columns:
                original_text = row[col]
                if isinstance(original_text, str):
                    try:
                        translated_text = translate_text(original_text, lang_code, translator_instance)
                        df.at[index, col] = translated_text if translated_text else original_text
                    except Exception as e:
                        print(f"Failed to translate text in row {index}, column '{col}' due to: {e}")
                        failed_indices.append((index, col))
        if failed_indices:
            failed_sheets[sheet_name] = failed_indices
        translated_dfs[sheet_name] = df
    return translated_dfs, failed_sheets


def save_output(df, directory, sheet_name, writer, write_header=True):
    if not os.path.exists(directory):
        os.makedirs(directory)
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=write_header)

if __name__ == '__main__':

    print('-' * 20 + 'Create by Celmax' + '-' * 20)
    print("Welcome to the Sheets Translator!")
    print("Select a translator from the following options:")

    for num, translator in translator_options.items():
        print(f"{num}: {translator.__name__}")
    translator_number = int(input("Select a translator by number: "))
    
    api_key_needed = translator_number in [3, 4, 7]
    api_key = None
    if api_key_needed:
        api_key = update_api_key_prompt(list(translator_options.values())[translator_number - 1].__name__.replace('Translator', '').lower())

    while True:
        lang_choice = input("Enter the language you want to translate your text into or type 'help' to list all languages: ").lower()
        if lang_choice == "help":
            print("Available languages:")
            for code, lang in sorted(languages.items(), key=lambda item: item[1]):
                print(f"{lang.title()} ({code})")
        else:
            lang_code = next((code for code, lang in languages.items() if lang_choice in lang), None)
            if lang_code:
                break
            print("Sorry, the selected language isn't supported for translation. Please try again.")

    translator_instance = get_translator(translator_number, lang_code, api_key)

    file_mode = input("Do you want to translate the first column (1) or all sheets (2)?: ")
    t1 = time.perf_counter()

    if file_mode == '1':
        directory = "column"
    elif file_mode == '2':
        directory = "all"
    else:
        print("Invalid input. Please specify '1' for column or '2' for all sheets.")
        exit()

    output_name = input("Enter a name for the output file (without extension): ")
    output_path = os.path.join(directory, f"{output_name}_{timestr}.xlsx")
    with pd.ExcelWriter(output_path) as writer:
        if file_mode == '1':
            df = pd.read_excel('input_data_one_column.xlsx', usecols=[0])
            df.columns = ['Input Text']
            df, failed_indices = translate_column(df, lang_code, translator_instance)
            save_output(df, "column", 'Translated', writer, write_header=True)
            if failed_indices:
                print("Failed to translate the following rows:", failed_indices)
        elif file_mode == '2':
            xls = pd.ExcelFile('input_data_all.xlsx')
            translated_dfs, failed_sheets = translate_all_sheets(xls, lang_code, translator_instance)
            for sheet_name, df in translated_dfs.items():
                save_output(df, "all", sheet_name, writer, write_header=False)
            if failed_sheets:
                print("Failed to translate the following cells in sheets:", failed_sheets)
        else:
            print("Invalid input. Please specify '1' for column or '2' for all sheets.")
            exit()
    print("File saved successfully:", output_path)

    t2 = time.perf_counter()
    execution_time = t2 - t1
    print('-' * 40)
    print("Translation complete!")
    print(f"Execution time: {execution_time:.2f} seconds.")
    print("Thank you for using the Sheets Translator!")
    print("Goodbye! See you next time!")
    print('-' * 40)