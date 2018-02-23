
import openpyxl
from itertools import permutations


def generate_permutations(s):

    perms = [''.join(p) for p in permutations(s.__str__())]
    return perms


def find_word(s, word):

    if word == s:
        return True
    return False


book = openpyxl.load_workbook('PersianWords.xlsx')
sheet = book.active
result = []
words = []
dic = input(" Enter A for Dehkhoda's Dictionary \nB for Moin's Dictionary \nC for openoffice.fa.dictionaries Persian Spellchecker Dictionary \nD for Fa.wikipedia's article's title \nand E for Total farsi Word ")
inp = input("Enter persion word")
for i in range(32640):
    s  = dic.__str__()+(i+5).__str__()
    words.append(sheet[s])
    sum = 0
    sumdef = 0
    for j in range(len(words[i].value.__str__())):
        sum += ord(words[i].value[j])
    for j in range(len(inp)):
        sumdef += ord(inp[j])
    if sum == sumdef:
        gp = generate_permutations(inp)
        for l in range(len(gp)):
            if find_word(words[i].value ,gp[l]):
                result.append(gp[l])
print("anagrams are :")
print (result.__str__())


