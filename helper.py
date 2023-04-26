def parseAddress(address):
    split_address = address.split()
    city = split_address[0]
    province = split_address[1]
    postal_code = " ".join(split_address[2:])
    return city, province, postal_code    

def trim_list(lst):
    # Find start and end indices
    start_index = lst.index('Location of Practice:') + 1
    end_index = None
    for word in lst:
        if 'This doctor has' in word:
            end_index = lst.index(word)
            break
        elif 'Area(s)' in word:
            end_index = lst.index(word)
            break

    # Get relevant elements
    if end_index != None:
        info_list = lst[start_index:end_index]
        return info_list 
    
def extract_info(lst):

    streetOne = lst[0]
    streetTwo = ''
    inc = 0
    if len(lst) == 5:
        streetTwo = lst[1]
        inc = 1

    try:
        city, province, postal_code = parseAddress(lst[1+inc])
        phone = ""
        fax = ""
        for item in lst[2+inc:]:
            if "Phone:" in item:
                phone = item.replace("Phone: ", "")
            elif "Fax:" in item:
                fax = item.replace("Fax: ", "")

        return [streetOne, streetTwo, city, province, postal_code, "Canada", phone, fax]
    except Exception as e:
        print(f"Error occured: {e}")
    else:
        return None    

