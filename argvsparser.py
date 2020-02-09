import sys
from datetime import datetime
import functools


argv = sys.argv[1:]


def get_arg_value(type):

    argv_len = len(argv)

    index = argv.index(type) if type in argv else -1

    if index < 0 or index >= argv_len - 1:
        return None
    else:
        value = argv[index+1]
        return value


def get_lable_value(type):
    return type in argv


def parse_date(value, default):
    if not value:
        return default
    else:
        return datetime.strptime(value, '%Y-%m-%d')


def parse_int(value, default):
    if not value:
        return default
    else:
        return int(value)


def parse_str(value, default):
    if not value:
        return default
    else:
        return value


def parse_boolean(value, default):
    if not value:
        return default
    else:
        return bool(value)


def parse_func(type, default, islabel=False):
    if islabel:
        exists = get_lable_value(type)
        if exists:
            return functools.partial(parse_boolean, value='True',
                                     default=default)
        else:
            return functools.partial(parse_boolean, value=default,
                                     default=default)

    valuestr = get_arg_value(type)
    getterfunc = argv_getter.get(type)
    return functools.partial(getterfunc, value=valuestr, default=default)


argv_getter = {'-d': parse_date,
               '-c': parse_int,
               '-s': parse_int,
               '-f': parse_str,
               '-f': parse_str,
               '-e': parse_boolean}
