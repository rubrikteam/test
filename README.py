# test

import string
text= "hellowsaikiraniam1993. my car blue"
clean_text= text.translate(str.maketrans('','',string.punctuation))
word_list= clean_text.split()
print(word_list)
