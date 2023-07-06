from googletrans import Translator
from tqdm import tqdm

testing = ["Halo", "Apa Kabar", "Lagi ngapain kamu?"]
translator = Translator()
res = []
for text in tqdm(testing):
    translated = translator.translate(testing, dest='en')
    res = translated.text
print(res)
# langs_list = GoogleTranslator().get_supported_languages()  # output: [arabic, french, english etc...]
# print(langs_list)