import re
from dateutil import parser


def remove_blank(text):
    return str(text).strip().replace('\r', '').replace('\t', '').replace('\n', '').replace(u'\u3000', u'').replace(u'\xa0', u'').replace(' ', '')
