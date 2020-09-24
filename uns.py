from phpserialize import *
from collections import OrderedDict
from phpserialize import serialize, unserialize
import excel

# products = loads(loads(serialize(
#     'a:5:{s:4:"code";s:9:"IM1107005";s:5:"stock";s:1:"1";s:5:"price";s:6:"394.44";s:9:"old-price";N;s:5:"promo";N;}')))
# a = products[b'code'].decode('utf-8')
# print(a)

product_codes = []

products = excel.get_all_the_rows_from_column('AL')

for prod in products:
    if prod is not None:
        dictionary = loads(loads(serialize(prod)))
        p_code = dictionary[b'code']
        if p_code is not None:
            product_codes.append(p_code.decode('utf-8'))
        else:
            product_codes.append('n/a')
    else:
        product_codes.append('n/a')

excel.write_product_code_to_excel(product_codes, 'AM')

# key_columns = ['I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC']
# value_columns = ['J', 'L', 'N', 'P', 'R', 'T', 'V', 'X', 'Z', 'AB', 'AD']
#
# # key_columns = ['G', 'K']
# # value_columns = ['I', 'M']
#
# for i in range(len(key_columns)):
#     cell_index = excel.match_key_value(key_columns[i], value_columns[i])

# excel.match_key_value(key_columns[0], value_columns[0])
