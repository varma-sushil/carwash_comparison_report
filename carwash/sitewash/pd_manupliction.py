import os 

import pandas as pd

current_folder_path = os.path.dirname(os.path.abspath(__file__))
xlfile_path = os.path.join(current_folder_path,"sitewash_data.xlsx")

sites_df = pd.read_excel(xlfile_path)
print(len(sites_df))

client_name = sites_df["client_name"].to_list()

print(client_name)