from pybaseball import batting_stats
import pandas as pd

data = batting_stats(2025)
print(data.head())

with pd.ExcelWriter('statcast_data.xlsx', engine='openpyxl') as writer:
    if not data.empty:
        data.to_excel(writer, sheet_name='July_4_2017', index=False)
