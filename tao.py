import pandas as pd

columns = ["Husband Name", "Wife Name", "Husband ID", "Wife ID", "Address", "Collateral Value", "Loan Amount", "Loan Duration", "Interest Rate", "Interest Amount", "Total Payment"]
df = pd.DataFrame(columns=columns)
df.to_excel("loan_data.xlsx", index=False, engine='openpyxl')