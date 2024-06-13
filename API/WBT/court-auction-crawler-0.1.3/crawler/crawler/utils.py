import hashlib
from datetime import datetime
import json


def generate_id(key, sort=False, bson=False):
    message = hashlib.md5()
    if sort:
        key = {k: key[k] for k in sorted(key)}
    if bson:
        message.update(bson_dumps(key).encode('utf-8'))
    else:
        message.update(json.dumps(key).encode('utf-8'))
    return message.hexdigest()


def parse_tw_date(date, seq='/'):
    try:
        if not date:
            return None
        date = date.split(seq)
        year = sum([int(val) * (10**i)
                    for i, val in enumerate(date[0][::-1])]) + 1911
        return datetime(
            year, int(date[1]), int(date[2]))
    except (ValueError, TypeError):
        return None
