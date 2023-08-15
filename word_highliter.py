import pandas as pd
import re
import nltk
from nltk.corpus import wordnet

# Download the required NLTK resources (if not already downloaded)
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('wordnet')

# Map NLTK POS tags to WordNet POS tags
tag_map = {
    'N': wordnet.NOUN,
    'V': wordnet.VERB,
    'R': wordnet.ADV,
    'J': wordnet.ADJ
}

# read in the Excel file
df = pd.read_excel('words.xlsx')

# iterate over each row in the DataFrame
for index, row in df.iterrows():
    # extract the word from the second row and make it lowercase
    word = str(row[1]).lower()

    # extract the sentence from the fifth row
    sentence = str(row[5])

    # tokenize and tag the words in the sentence
    tokens = nltk.word_tokenize(sentence)
    tagged_words = nltk.pos_tag(tokens)

    # find the word in the sentence using POS tags
    found = False
    for tagged_word in tagged_words:
        word_pos = tagged_word[0].lower()
        pos_tag = tagged_word[1][:1]  # Extract the first character of the tag

        # check if the word or its lemmatized form matches
        if word == word_pos or word == nltk.WordNetLemmatizer().lemmatize(word_pos, tag_map.get(pos_tag, wordnet.NOUN)):
            found = True
            break

    if found:
        # replace the word with the word surrounded by <>
        new_sentence = re.sub(r"\b{}\b".format(re.escape(word)), r"<{}>".format(word), sentence, flags=re.IGNORECASE)

        # update the DataFrame with the new sentence
        df.at[index, 5] = new_sentence
    else:
        # print out the row index for each row that is skipped
        print(f"Row {index} skipped: word '{word}' not found in sentence '{sentence}'")

# write the updated DataFrame to a new Excel file
df.to_excel('A2_output.xlsx', index=False)