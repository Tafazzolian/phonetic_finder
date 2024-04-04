import pandas as pd
import re
import nltk
from nltk.corpus import wordnet

nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('wordnet')

tag_map = {
    'N': wordnet.NOUN,
    'V': wordnet.VERB,
    'R': wordnet.ADV,
    'J': wordnet.ADJ
}

df = pd.read_excel('words.xlsx')

for index, row in df.iterrows():
    word = str(row[1]).lower()

    sentence = str(row[5])

    tokens = nltk.word_tokenize(sentence)
    tagged_words = nltk.pos_tag(tokens)

    found = False
    for tagged_word in tagged_words:
        word_pos = tagged_word[0].lower()
        pos_tag = tagged_word[1][:1]

        if word == word_pos or word == nltk.WordNetLemmatizer().lemmatize(word_pos, tag_map.get(pos_tag, wordnet.NOUN)):
            found = True
            break

    if found:
        new_sentence = re.sub(r"\b{}\b".format(re.escape(word)), r"<{}>".format(word), sentence, flags=re.IGNORECASE)
        df.at[index, 5] = new_sentence
    else:
        print(f"Row {index} skipped: word '{word}' not found in sentence '{sentence}'")

df.to_excel('A2_output.xlsx', index=False)
