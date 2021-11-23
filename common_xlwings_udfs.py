import re
import xlwings as xw
import pandas as pd
import time


class Address:
    def __init__(self):
        self.origin_string = ""
        self.region = ""
        self.city = ""
        self.street = ""
        self.house = ""


class AddressManager:

    @classmethod
    def filter_joined_text1(cls, inp):
        ptns_rules = [
            {"ptn": r'(.+)\s*,\s*к(?:орпус)?[\s\.]*(\d+)$', "sub": r'\1/\2'},
            {"ptn": r'(.+)\s*,\s*стр(?:оение)?[\s\.]*([\d/]+[а-яА-Я]?)$', "sub": r'\1,\2'},
            {"ptn": r'(.+)\s*,\s*соор(?:ужение)?[\s\.]*([\d/]+[а-яА-Я]?)$', "sub": r'\1,\2'},
            {"ptn": r'(.+[\d/]+[а-яА-Я]?)\s*кв(?:артира)?[\s\.]*([\d/]+[а-яА-Я]?)$', "sub": r'\1'},
            {"ptn": r'(.+?)\s*\(.*\)$', "sub": r'\1'},
            {"ptn": r'(.+)\s*т[\\/-]с$', "sub": r'\1'},
            {"ptn": r'(.+,\d+[а-ик-юА-ИК-Ю]?)[\s,]+(\d+[а-яА-Я]?)$', "sub": r'\1/\2'},
        ]
        txt = inp.strip()
        for rule in ptns_rules:
            p = re.compile(rule["ptn"], re.IGNORECASE)
            txt = p.sub(rule["sub"], txt)
        return txt

    @classmethod
    def filter_text1(cls, inp):
        ptns1 = [
            r"^дом[\s№\.]+(.+)",
            r"^г(?:ор|ород)?[\.\s]+(.+)",
            r"^п(?:ос|оселок)?[\.\s]+(.+)",
            r"^с(?:ел|ело)?[\.\s]+(.+)",
            r"^ш[\.\s]+(.+)",
            r"^м[\.\s]+([^(?:к\.)].+)",
            r"^п(?:[сг]т)?[\.\s]+(.+)",
            r"^д(?:ер)?[\.\s]+(.+)",
            r"^д(\d+[а-яА-Я]?)",
            r"^ул[\.\s]+(.+)",
            r"^пл[\.\s]+(.+)",
            r"^вл[\.\s]+(.+)",
            r"^зд[\.\s]+(.+)",
            r"^гар[\.\s]+(.+)",
            r"^сан[\.\s]+(.+)",
            r"^наб[\.\s]+(.+)",
            r"^пер[\.\s]+(.+)",
            r"^мкр[\.\s]+(.+)",
            r"^кв(?:арта|-)л[\.\s]+(.+)",
            r"^пр(?:оез|-)д[\.\s]+(.+)",
            r"^уч\-к[\.\s]+(.+)",
            r"^пр(?:оспект|осп|-кт|-т)?[\.\s]+(.+)",
            r"^стр(?:оение)?[\.\s]+(.+)",
            r"^б\-р[\.\s]+(.+)",
            r"(.+)\s+[гс]\.?$",
            r"(.+)\s+ул\.?$",
            r"(.+)\s+пер(?:еулок)?\.?$",
            r"(.+)\s+б(?:ульва|-)р\.?$",
            r"(.+)\s+пл\.?$",
            r"(.+)\s+мкр\.?$",
            r"(.+)\s+м\.?$",
            r"(.+)\s+микрорайон\.?$",
            r"(.+)\s+д\.?$",
            r"(.+)\s+п/р\.?$",
            r"(.+)\s+пр(?:оспект|-кт|-т)?\.?$",
            r"(.+)\s+ш\.?$",
            r"(.+)\s+рп\.?$",
            r"(.+)\s+ст\.?$",
            r"(.+)\s+пр(?:оез|-)д\.?$",
            r"(.+)\s+кв\-л\.?$",
            r"(.+)\s+наб\.?$",
            r"(.+)\s+п(?:[сг]т)?\.?$",
        ]
        txt = inp.strip()
        for ptn in ptns1:
            m = re.search(ptn, txt, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return inp

    # TODO
    @classmethod
    def filter_house_text1(cls, inp):
        # цифры в конце
        ptns_digit = [
            r'(.*\d[а-яА-Я]?)[,\s]+к(?:ор|орп|орпус)?\.?\s*(\d.*)',
            r'(.*\d[а-яА-Я]?)[,\s]+стр(?:оение)?\.?\s*(\d.*)',
            r'(.*\d[а-яА-Я]?)\s*-\s*(\d)$',
        ]
        sub_digit = r'\1/\2'
        # буквы в конце
        ptns_letter = [
            r'(.*\d[а-яА-Я]?)[,\s]+к\.?\s+([а-яА-Я].*)',
            r'(.*\d[а-яА-Я]?)[,\s]+стр(?:оение)?\.?\s+([а-яА-Я].*)',
            r'(.*\d[а-яА-Я]?)[,\s]+лит(?:ера)?\.?\s+([а-яА-Я].*)',
            r'(.*\d)[\s-]+([а-яА-Я].*)',
        ]
        sub_letter = r'\1\2'

        # сборка правил
        data = [
            {"ptns": ptns_digit, "sub": sub_digit},
            {"ptns": ptns_letter, "sub": sub_letter},
        ]
        ptns_rules = []
        for row in data:
            for ptn in row["ptns"]:
                d = {"ptn": ptn, "sub": row["sub"]}
                ptns_rules.append(d)

        txt = inp
        for rule in ptns_rules:
            p = re.compile(rule["ptn"], re.IGNORECASE)
            txt = p.sub(rule["sub"], txt)
        return txt

    @classmethod
    def split1(cls, address_string):
        adr = Address()
        adr.origin_string = address_string

        sects = address_string.split(',')
        sects = list(map(cls.filter_text1, sects))

        adr.city = sects[0]
        adr.street = sects[1]
        adr.house = sects[2]
        return adr

    @classmethod
    def get_key1(cls, address):
        txt = "{},{},{}".format(address.city, address.street, address.house.replace(" ", ""))
        txt = txt.lower().replace(" ", "")
        return txt

    @classmethod
    def preprocess1(cls, address_string):
        clear_ptns1 = [
            r'\d{6,}',
            r'^Россия$',
            r'Коми\s*Респ',
            r'^РК$',
            r'^РФ$',
            r'^Респ(?:ублика?)?[\s\.]*[а-яА-Я]+$',
            r'Российская\s*Федерация',
            r'Городской\s*округ',
            r'Городское\s*поселение',
            r'сельское\s+поселен',
            r'муниципальный\s*район',
            r'^\s*(?:пгт|пос|г|п)\.?\s*(?:воргашор|Северный|комсомольский|заполярный|Сивомаскинский)\s*$',
            r'^\s*(?:воргашор|Северный|комсомольский|заполярный|Сивомаскинский)\s*(?:пгт|пос|г|п)\.?\s*$',
            r'^\s*(?:шудаяг|абезь|абезъ)\s*(?:пгт|пос|г|п)\.?\s*$',
            r'^р(?:айо|-)н\.?\s*[а-яА-Я]+$',
            r'^[а-яА-Я]+[\s\.]+р(?:айон|-н|-он)\.?$',
            r'^[а-яА-Я]+\s+р(?:айо|-)н\.?$',
            r'^[а-яА-Я]+\s+обл(?:асть)?\.?$',
            r'^обл(?:асть)?\.?\s+[а-яА-Я]+$',
        ]
        #adr = Address()
        #adr.origin_string = address_string

        # разделить по запятым
        if ',' in address_string:
            sects = address_string.split(',')
        else:
            sects = address_string.split(' ')
        # убрать пробелы и пустоты
        sects = [x.strip() for x in sects if x.strip() != '']
        sects2 = []

        # убрать лишние поля
        for sect in sects:
            check = True
            for ptn in clear_ptns1:
                if re.search(ptn, sect, re.IGNORECASE):
                    check = False
                    break
            if check:
                sects2.append(sect)
        sects = sects2

        # фильтр полей
        sects = list(map(cls.filter_text1, sects))
        sects = list(map(cls.filter_house_text1, sects))

        #adr.city = sects[0]
        #adr.street = sects[1]
        #adr.house = sects[2]

        # собрать
        jo = ','.join(sects)
        # еще фильтры
        jo = cls.filter_joined_text1(jo)
        return jo
        # ====================================


@xw.func
def address_key(address_string):
    adr = AddressManager.split1(address_string)
    return AddressManager.get_key1(adr)


@xw.func
def address_key2(address_string):
    adr = adr_preprocess1(address_string)
    adr = AddressManager.split1(adr)
    return AddressManager.get_key1(adr)


@xw.func(async_mode='threading')
def myfunction(a):
    time.sleep(5)  # long running tasks
    return a


@xw.func
def hello2(name):
    return 'Hello {0}'.format(name)


@xw.func
@xw.arg('x', pd.DataFrame)
def correl2(x):
    # x arrives as DataFrame
    return x.corr()


# get 1st group from regex
@xw.func
def regex(ptn, text):
    res = text
    m = re.search(ptn, text, re.IGNORECASE)
    if m:
        res = m.group(0)
    return res


# get n group from regex
@xw.func
def regex_g(ptn, text, group_index):
    res = text
    m = re.search(ptn, text, re.IGNORECASE)
    if m:
        res = m.group(int(group_index))
    return res


@xw.func
def regex_split_combine(ptn, text, delim=","):
    lst = re.split(ptn, text, re.IGNORECASE)
    return delim.join(lst)


@xw.func
def adr_preprocess1(address_string):
    return AddressManager.preprocess1(address_string)


@xw.func
def filter_firm_name(text):
    # clear_ptns = [
    #    '(?i)\W*ООО\W',
    #    '(?i)\W*УК\W',
    #    '(?i)\W*ТСЖ\W',
    # ]
    ptn = r"[^\da-zа-я]+"
    s = re.sub(ptn, "", text, flags=re.IGNORECASE + re.MULTILINE)
    return s


@xw.func
def filter_tv_number(text):
    s = re.sub(r"\s*\([^\)]*\)\s*$", "", text, flags=re.IGNORECASE + re.MULTILINE)
    m = re.search(r"[^\d-]0*(\d+)\s*$", s, flags=re.IGNORECASE + re.MULTILINE)
    if m:
        res = m.group(1)
    else:
        res = ""
    return res


if __name__ == "__main__":
    #s = regex_g(r"([^,]+),([^,]+)", r"Инта г, Бабушкина ул, д. 1 (МКД УК \"Юпитер\"(к))", 2)
    s = adr_preprocess1("169849, Коми Респ, Инта г, Бабушкина ул, д. 1 К")
    #s = adr_preprocess1("участковое лесничество, стр. 245/4")
    #s = re.sub(r'(.*\d) к\. (\d.*)', r'\1/\2', 'qwer 1234 к. 2 qwer')
    #s = AddressManager.filter_house_text1("40  К.  а2")
    #s = address_key("Сосногорск,Береговая,стр. 6")
    #s = filter_firm_name("ООО УК \"Рога & Копыта-123\"")
    #s = filter_tv_number("1 узел 00117443 (ВКТ7)")
    print(s)
