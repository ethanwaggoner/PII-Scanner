import re


# Validates if the credit card number is legitimate using Luhn's algorithm
def luhn(n):
    r = [int(ch) for ch in str(n)][::-1]
    return (sum(r[0::2]) + sum(sum(divmod(d*2, 10)) for d in r[1::2])) % 10 == 0


# Return the censored version of the PII. Censors everything except dashes and the last 4 characters
def censor(pii_list):
    pii_return = []

    for pii in pii_list:

        to_censor = pii[:-4]
        uncensored = pii[-4:]
        censored = []

        for char in to_censor:
            if char != "-":
                char = "*"
                censored.append(char)
            else:
                censored.append(char)

        censored_word = "".join(censored) + uncensored
        pii_return.append(censored_word)

    return pii_return


# Runs regex against the data looking for Social Security Numbers
def ssn(data):
    pii_list = re.findall(r"""\b\d{3}-\d{2}-\d{4}\b|\b\d{3}\s\d{2}\s\d{4}\b|\b\d{3}-\d{6}\b""", data)

    return censor(pii_list)


# Runs regex against the data looking for Credit Card Numbers
def cc(data):
    pii_list = re.findall(r"\b\d{4}\s\d{4}\s\d{4}\s\d{4}\b|\b\d{16}\b|\b\d{4}-\d{4}-\d{4}-\d{4}\b", data)
    pii_clean = []
    for pii in pii_list:
        pii = pii.replace("-", "")
        pii = pii.replace(" ", "")
        if luhn(pii):
            pii_clean.append(pii)

    return censor(pii_clean)


# Runs regex against the data looking for Maryland Drivers License IDs
def drivers_md(data):
    pii_list = re.findall(r"\b[a-zA-Z]\d{12}\b|\b[a-zA-z]-\d{3}-\d{3}-\d{3}-\d{3}\b", data)

    return censor(pii_list)


