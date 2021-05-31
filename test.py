from dicts import lifetime

string = "Стол письменный"

for key in lifetime:
    if key.lower() in string.lower():
        print(string.lower())