from yargy import rule, or_, Parser
from yargy.predicates import gram
import docx2txt
from natasha import (
    span,
    Segmenter,
    MorphVocab,

    NewsEmbedding,
    NewsMorphTagger,
    NewsSyntaxParser,
    NewsNERTagger,

    PER,
    NamesExtractor,
    DatesExtractor,
    MoneyExtractor,
    AddrExtractor,

    Doc
)

segmenter = Segmenter()
morph_vocab = MorphVocab()

emb = NewsEmbedding()
morph_tagger = NewsMorphTagger(emb)
syntax_parser = NewsSyntaxParser(emb)
ner_tagger = NewsNERTagger(emb)

names_extractor = NamesExtractor(morph_vocab)
dates_extractor = DatesExtractor(morph_vocab)
money_extractor = MoneyExtractor(morph_vocab)
addr_extractor = AddrExtractor(morph_vocab)


FIRST = gram('Name')
LAST = gram('Surn')
MIDDLE = gram('Patr')
ABBR = gram('Abbr')

NAME = or_(
    rule(FIRST),
    rule(LAST),
    rule(FIRST, LAST),
    rule(LAST, FIRST),
    rule(FIRST, MIDDLE, LAST),
    rule(LAST, FIRST, MIDDLE),
    rule(ABBR, '.', ABBR, '.', LAST),
    rule(LAST, ABBR, '.', ABBR, '.'),
)
parser = Parser(NAME)


text = ('1 599 059, 38 Евро, 420 долларов, 20 млн руб, 20 т. р., 881 913 '
        '(Восемьсот восемьдесят одна тысяча девятьсот тринадцать) руб. 98 коп.')
def findMoney(text):
    #result = docx2txt.process(text)  возможность загрузить файл
    text = str(text)
    for money in list(money_extractor(text)):
      money = str(money.fact)
      indx1 = money.index('amount=') + 7
      indx2 = money.index(', c')
      indx3 = money.index('currency=') + 9
      indx4 = money.index(')')

      summ = money[indx1:indx2].lstrip().rstrip()
      valute = money[indx3:indx4].lstrip().rstrip()
      print('сумма:', summ, ',', 'валюта:', valute)
#findMoney(text)


text = '''Россия, Вологодская обл. г. Череповец, пр.Победы 93 б',
   692909, РФ, Приморский край, г. Находка, ул. Добролюбова, 18
   ул. Народного Ополчения д. 9к.3'''

def findAdress(text):
    #result = docx2txt.process(text)   возможность загрузить файл
    text = str(text)
    for addr in list(addr_extractor(text)):
        print(addr.fact)
#findAdress(text)



dateFile = "file.docx"
def findDate(nameOfFile):
    result = docx2txt.process(nameOfFile)
    text = str(result)
    sp_date = []
    for date in list(dates_extractor(text)):
        date = str(date.fact)
        indx1 = date.index('year=') + 5
        indx2 = date.index(', m')
        indx3 = date.index('month=') + 6
        indx4 = date.index(', d')
        indx5 = date.index('day=') + 4
        indx6 = date.index(')')
        dt_year = date[indx1:indx2].lstrip().rstrip()
        dt_month = date[indx3:indx4].lstrip().rstrip()
        dt_day = date[indx5:indx6].lstrip().rstrip()
        date_main = ('год:', dt_year, ',', 'месяц:', ',', dt_month, ',', 'день:', dt_day)
        sp_date.append(date_main)

    num = 0
    for match in parser.findall(text):
        start, stop = match.span
        main_name = text[start:stop]
        if len(main_name) >= 3 and main_name.istitle():
            sp = [main_name.split(' ')]
            if len(sp[0]) > 1:
                print(*sp[0], ':',   *sp_date[num])
                num += 1


#findDate(dateFile)

def findNames(nameOfFile):
    result = docx2txt.process(nameOfFile)
    text = str(result)
    names = list(names_extractor(text))
    for name in names:
       name = str(name)
       if 'first=None' in name and 'middle=None' in name:
           continue
       elif 'last=None, middle=None' in name:
           continue
       else:
           print(name)


#findNames(dateFile)



