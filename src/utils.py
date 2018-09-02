def get_column_letters(n, zero_indexed = True):
    if zero_indexed: 
        n += 1

    string = ""
    
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def get_column_index(n, zero_indexed = True):
    pass