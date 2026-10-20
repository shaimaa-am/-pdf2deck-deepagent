from json import *

def dumps(*args, **kwargs):
    return __import__('json').dumps(*args, **kwargs)

def loads(*args, **kwargs):
    return __import__('json').loads(*args, **kwargs)