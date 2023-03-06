import re
import pandas as pd


cat = pd.read_csv("./catfinder.csv", index_col=False)
# raw = pd.read_csv("./raw.csv", index_col=0, parse_dates=True)
#cutsheet = pd.read_excel("./temp.xlsx", parse_dates=True)
cutsheet_loc = pd.ExcelFile("./temp.xlsx")
      
# cat["average"] = cat["quantity"] * cat["price"]
# print(cat.head(10))
# cat_device = cat[["device"]]
# cutsheet_device = cutsheet["device"]
# move_cols = inner.reindex(columns=['location','ru','device'])
# inner2.to_excel('new_dataframe.xlsx',index=False)
# print("saving file...")
# final_df.to_excel('new_dataframe.xlsx',index=False)
# print("file saved!")
#rn_col = cat.rename(columns={"device":"a_hostname"})
# inner = pd.merge(cutsheet,rn_col, on="a_hostname", how="left")
# cols = list(inner.columns.values)
# cols.pop(cols.index('location'))
# cols.pop(cols.index('ru'))
# inner = inner[['location','ru']+cols]
# new_cat_df = inner

# rn_col2 = cat.rename(columns={"device":"z_hostname"})
# inner2  = pd.merge(new_cat_df,rn_col2, on="z_hostname", how="left")
# cols = list(inner2.columns.values)
# inner2 = inner2[['location_x', 'ru_x', 'a_hostname', 'a_interface', 'cable_type', 'z_connector_type', 'z_interface', 'z_hostname', 'ru_y','location_y']]
# final_df = inner2.rename(columns={"location_x": "location","ru_x":"ru","location_y": "location","ru_y":"ru"})
#   for x in new_cat_df["location"]:
#         new = str(x)
#         new2 =new.replace(new[:10],'')
#         print(new2)


writer = pd.ExcelWriter('./new.xlsx')
for sheet_name in cutsheet_loc.sheet_names:
    cutsheet = pd.read_excel("./temp.xlsx", sheet_name=sheet_name)
    rm_col = cat.rename(columns={"device":"a_hostname"})
    inner = pd.merge(cutsheet,rm_col, on="a_hostname", how="left")
    cols = list(inner.columns.values)
    cols.pop(cols.index('location'))
    cols.pop(cols.index('ru'))     
    inner = inner[['location','ru']+cols]
    new_cat_df = inner
    
    rn_col2 = cat.rename(columns={"device":"z_hostname"})
    inner2  = pd.merge(new_cat_df,rn_col2, on="z_hostname", how="left")
    cols = list(inner2.columns.values)
    inner2 = inner2[cols]
    final_df = inner2.rename(columns={"location_x": "location","ru_x":"ru","location_y": "location","ru_y":"ru"})
    final_df.to_excel(writer, sheet_name=sheet_name, index=False)
 
       
    
writer.save()
writer.close()    
