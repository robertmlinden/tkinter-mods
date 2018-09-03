def get_column_letters(n, zero_indexed = True):
    if zero_indexed: 
        n += 1

    string = ""
    
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def letter_to_index(letter):
    return ord(letter.lower()) - ord('a') + 1

def get_column_index(s, go = False):
    if not go:
        return get_column_index(s, go=True) - 1
    else:
        if len(s) == 0:
            return 0

        return letter_to_index(s[-1]) + 26*get_column_index(s[:-1], go=True)