import requests
import os
from datetime import datetime

url = 'https://www.housing-stat.ch/files/Listes_cantons/Listes_NE.xlsx'
save_dir = r'U:\projets\1_MO\2022_RegBL\Listes_NE'


r = requests.get(url, allow_redirects=True)

filename = datetime.strftime(datetime.now(), '%Y%m%d_Listes_NE.xlsx')

filepath = os.path.join(save_dir, filename)
open(filepath, 'wb').write(r.content)
