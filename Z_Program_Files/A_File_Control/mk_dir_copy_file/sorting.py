# http://www.artnet.com/artists/lois-gross-smiley/suburban-summer-a-zKoQxl-TiSOq

import pandas as pd

path_b = 'c:/GDSRC/DATA/Result_biopsy_서정애.xlsx'

B_sorted = pd.read_excel(path_b)

# print(B_sorted)

Sorted_T = B_sorted.sort_values(['보험코드'], ascending=True)

print(Sorted_T)

Sorted_T.to_excel('c:/GDSRC/DATA/Result_biopsy_서정애_S.xlsx')