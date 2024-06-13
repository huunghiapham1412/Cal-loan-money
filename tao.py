import pandas as pd

# Tạo một DataFrame trống với các cột thích hợp
columns = ["Husband Name", "Wife Name", "Husband ID", "Wife ID", "Address", "Collateral Value", "Loan Amount", "Loan Duration", "Interest Rate", "Interest Amount", "Total Payment"]
df = pd.DataFrame(columns=columns)

# Lưu DataFrame này vào tệp Excel mới
df.to_excel("loan_data.xlsx", index=False, engine='openpyxl')