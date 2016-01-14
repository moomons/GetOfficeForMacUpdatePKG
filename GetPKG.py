#!/usr/bin/env python

import requests
import plistlib
import pyperclip


dict_plist = {
    'AUTOUPDATE': 'http://www.microsoft.com/mac/autoupdate/0409MSau03.xml',
    'WORD': 'http://www.microsoft.com/mac/autoupdate/0409MSWD15.xml',
    'EXCEL': 'http://www.microsoft.com/mac/autoupdate/0409XCEL15.xml',
    'POWERPNT': 'http://www.microsoft.com/mac/autoupdate/0409PPT315.xml',
    'ONENOTE': 'http://www.microsoft.com/mac/autoupdate/0409ONMC15.xml',
    'OUTLOOK': 'MISSING URL'
};


def getPKGUrl(plistURL):
    u = requests.get(plistURL)
    plist_Word = plistlib.readPlistFromString(u.content)
    return(plist_Word[0]["Location"])


allPKGURL = \
    getPKGUrl(dict_plist["WORD"]) + '\n' + \
    getPKGUrl(dict_plist["EXCEL"]) + '\n' + \
    getPKGUrl(dict_plist["POWERPNT"]) + '\n' + \
    getPKGUrl(dict_plist["ONENOTE"]) + '\n' + \
    ''

print(allPKGURL)
pyperclip.copy(allPKGURL)
