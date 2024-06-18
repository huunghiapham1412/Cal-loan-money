import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

FILE_NAME = 'loan_data.xlsx'
data_list = []

# Kiểm tra và tạo tệp Excel nếu chưa tồn tại
def check_create_excel_file():
    if not os.path.exists(FILE_NAME):
        columns = ["Husband Name", "Wife Name", "Husband ID", "Wife ID", "Address", "Collateral Value", "Loan Amount", "Loan Duration", "Interest Rate", "Interest Amount", "Total Payment"]
        df = pd.DataFrame(columns=columns)
        df.to_excel(FILE_NAME, index=False)

check_create_excel_file()

# Các hàm

def load_data():
    return pd.read_excel(FILE_NAME, engine='openpyxl')

def save_data_to_excel(data):
    data.to_excel(FILE_NAME, index=False, engine='openpyxl')
    print("Data saved to Excel:")
    print(data)

def refresh_combobox():
    combobox_values = [f"{entry['Husband Name']} - {entry['Husband ID']}" for entry in data_list] + \
                      [f"{entry['Wife Name']} - {entry['Wife ID']}" for entry in data_list]
    combobox_update['values'] = combobox_values

def is_valid_name(name):
    return name.isalpha()

def is_valid_id(id_number):
    return id_number.isdigit()

def calculate_loan(loan_amount, loan_duration):
    if loan_duration < 30:
        interest_rate = 0.09
    elif 30 <= loan_duration <= 90:
        interest_rate = 0.06
    else:
        interest_rate = 0.03

    interest_amount = loan_amount * interest_rate
    total_payment = loan_amount + interest_amount
    return interest_rate, interest_amount, total_payment

def display_loan_results(interest_rate, interest_amount, total_payment):
    entry_interest_rate.config(state='normal')
    entry_interest_amount.config(state='normal')
    entry_total_payment.config(state='normal')

    entry_interest_rate.delete(0, tk.END)
    entry_interest_rate.insert(0, f"{interest_rate * 100:.2f}%")
    entry_interest_amount.delete(0, tk.END)
    entry_interest_amount.insert(0, f"{interest_amount:.2f} VND")
    entry_total_payment.delete(0, tk.END)
    entry_total_payment.insert(0, f"{total_payment:.2f} VND")

    entry_interest_rate.config(state='readonly')
    entry_interest_amount.config(state='readonly')
    entry_total_payment.config(state='readonly')

    entry_interest_rate.grid(row=10, column=1, padx=10, pady=10)
    entry_interest_amount.grid(row=11, column=1, padx=10, pady=10)
    entry_total_payment.grid(row=12, column=1, padx=10, pady=10)

def create_loan():
    try:
        husband_name = entry_husband_name.get()
        wife_name = entry_wife_name.get()
        husband_id = entry_husband_id.get()
        wife_id = entry_wife_id.get()
        address = entry_address.get()
        collateral_value = float(entry_collateral_value.get())
        loan_amount = float(entry_loan_amount.get())
        loan_duration = int(entry_loan_duration.get())

        if not is_valid_name(husband_name):
            messagebox.showerror("Error", "Tên người chồng chỉ được chứa chữ cái.")
            return
        if not is_valid_name(wife_name):
            messagebox.showerror("Error", "Tên người vợ chỉ được chứa chữ cái.")
            return
        if not is_valid_id(husband_id):
            messagebox.showerror("Error", "Căn cước công dân người chồng chỉ được chứa số.")
            return
        if not is_valid_id(wife_id):
            messagebox.showerror("Error", "Căn cước công dân người vợ chỉ được chứa số.")
            return
        if collateral_value <= loan_amount:
            messagebox.showerror("Error", "Giá trị tài sản thế chấp phải cao hơn số tiền mượn.")
            return

        interest_rate, interest_amount, total_payment = calculate_loan(loan_amount, loan_duration)
        display_loan_results(interest_rate, interest_amount, total_payment)

        new_entry = {
            "Husband Name": husband_name,
            "Wife Name": wife_name,
            "Husband ID": husband_id,
            "Wife ID": wife_id,
            "Address": address,
            "Collateral Value": collateral_value,
            "Loan Amount": loan_amount,
            "Loan Duration": loan_duration,
            "Interest Rate": interest_rate,
            "Interest Amount": interest_amount,
            "Total Payment": total_payment
        }
        data_list.append(new_entry)
        
        # Save to Excel
        df = pd.DataFrame(data_list)
        save_data_to_excel(df)

        messagebox.showinfo("Success", "Thêm khoản vay thành công.")
        refresh_combobox()
        refresh_treeview()
    except ValueError:
        messagebox.showerror("Error", "Vui lòng nhập số hợp lệ.")

def search_loan():
    search_query = combobox_update.get()
    identifier = search_query.split(" - ")[-1]
    
    for entry in data_list:
        if entry["Husband ID"] == identifier or entry["Wife ID"] == identifier:
            entry_husband_name.delete(0, tk.END)
            entry_husband_name.insert(0, entry['Husband Name'])

            entry_wife_name.delete(0, tk.END)
            entry_wife_name.insert(0, entry['Wife Name'])

            entry_husband_id.delete(0, tk.END)
            entry_husband_id.insert(0, entry['Husband ID'])

            entry_wife_id.delete(0, tk.END)
            entry_wife_id.insert(0, entry['Wife ID'])

            entry_address.delete(0, tk.END)
            entry_address.insert(0, entry['Address'])

            entry_collateral_value.delete(0, tk.END)
            entry_collateral_value.insert(0, entry['Collateral Value'])

            entry_loan_amount.delete(0, tk.END)
            entry_loan_amount.insert(0, entry['Loan Amount'])

            entry_loan_duration.delete(0, tk.END)
            entry_loan_duration.insert(0, entry['Loan Duration'])

            display_loan_results(entry['Interest Rate'], entry['Interest Amount'], entry['Total Payment'])

            messagebox.showinfo("Success", "Tìm thấy khoản vay.")
            return
    
    messagebox.showerror("Error", "Không tìm thấy khoản vay.")

def update_loan():
    search_query = combobox_update.get()
    identifier = search_query.split(" - ")[-1]
    
    for index, entry in enumerate(data_list):
        if entry["Husband ID"] == identifier or entry["Wife ID"] == identifier:
            try:
                husband_name = entry_husband_name.get()
                wife_name = entry_wife_name.get()
                husband_id = entry_husband_id.get()
                wife_id = entry_wife_id.get()
                address = entry_address.get()
                collateral_value = float(entry_collateral_value.get())
                loan_amount = float(entry_loan_amount.get())
                loan_duration = int(entry_loan_duration.get())

                if not is_valid_name(husband_name):
                    messagebox.showerror("Error", "Tên người chồng chỉ được chứa chữ cái.")
                    return
                if not is_valid_name(wife_name):
                    messagebox.showerror("Error", "Tên người vợ chỉ được chứa chữ cái.")
                    return
                if not is_valid_id(husband_id):
                    messagebox.showerror("Error", "Căn cước công dân người chồng chỉ được chứa số.")
                    return
                if not is_valid_id(wife_id):
                    messagebox.showerror("Error", "Căn cước công dân người vợ chỉ được chứa số.")
                    return
                if collateral_value <= loan_amount:
                    messagebox.showerror("Error", "Giá trị tài sản thế chấp phải cao hơn số tiền mượn.")
                    return

                interest_rate, interest_amount, total_payment = calculate_loan(loan_amount, loan_duration)
                display_loan_results(interest_rate, interest_amount, total_payment)

                updated_entry = {
                    "Husband Name": husband_name,
                    "Wife Name": wife_name,
                    "Husband ID": husband_id,
                    "Wife ID": wife_id,
                    "Address": address,
                    "Collateral Value": collateral_value,
                    "Loan Amount": loan_amount,
                    "Loan Duration": loan_duration,
                    "Interest Rate": interest_rate,
                    "Interest Amount": interest_amount,
                    "Total Payment": total_payment
                }
                
                data_list[index] = updated_entry
                
                # Save to Excel
                df = pd.DataFrame(data_list)
                save_data_to_excel(df)

                messagebox.showinfo("Success", "Cập nhật khoản vay thành công.")
                refresh_combobox()
                refresh_treeview()
                return
            except ValueError:
                messagebox.showerror("Error", "Vui lòng nhập số hợp lệ.")
                return
    
    messagebox.showerror("Error", "Không tìm thấy khoản vay để cập nhật.")

def delete_loan():
    search_query = combobox_update.get()
    identifier = search_query.split(" - ")[-1]
    
    for index, entry in enumerate(data_list):
        if entry["Husband ID"] == identifier or entry["Wife ID"] == identifier:
            del data_list[index]
            
            # Save to Excel
            df = pd.DataFrame(data_list)
            save_data_to_excel(df)

            messagebox.showinfo("Success", "Xóa khoản vay thành công.")
            refresh_combobox()
            refresh_treeview()
            return
    
    messagebox.showerror("Error", "Không tìm thấy khoản vay để xóa.")

def refresh_treeview():
    for item in tree.get_children():
        tree.delete(item)
    for entry in data_list:
        tree.insert('', 'end', values=list(entry.values()))

# Tạo cửa sổ chính
root = tk.Tk()
root.title("Loan Calculator")
root.geometry("1350x800")

# Tạo người mượn tiền
tk.Label(root, text="Tên người chồng:").grid(row=0, column=0, padx=10, pady=10)
entry_husband_name = tk.Entry(root)
entry_husband_name.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Tên người vợ:").grid(row=1, column=0, padx=10, pady=10)
entry_wife_name = tk.Entry(root)
entry_wife_name.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Căn cước công dân người chồng:").grid(row=2, column=0, padx=10, pady=10)
entry_husband_id = tk.Entry(root)
entry_husband_id.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Căn cước công dân người vợ:").grid(row=3, column=0, padx=10, pady=10)
entry_wife_id = tk.Entry(root)
entry_wife_id.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Chỗ ở hiện tại của cả 2 vợ chồng:").grid(row=4, column=0, padx=10, pady=10)
entry_address = tk.Entry(root)
entry_address.grid(row=4, column=1, padx=10, pady=10)

# Tạo tài sản thế chấp và thời gian mượn tiền
tk.Label(root, text="Giá trị thực tế tài sản thế chấp:").grid(row=5, column=0, padx=10, pady=10)
entry_collateral_value = tk.Entry(root)
entry_collateral_value.grid(row=5, column=1, padx=10, pady=10)

tk.Label(root, text="Số tiền mượn:").grid(row=6, column=0, padx=10, pady=10)
entry_loan_amount = tk.Entry(root)
entry_loan_amount.grid(row=6, column=1, padx=10, pady=10)

tk.Label(root, text="Số thời gian mượn (ngày):").grid(row=7, column=0, padx=10, pady=10)
entry_loan_duration = tk.Entry(root)
entry_loan_duration.grid(row=7, column=1, padx=10, pady=10)

# Tạo nút tính toán
button_calculate = tk.Button(root, text="Tính toán", command=create_loan)
button_calculate.grid(row=8, column=0, columnspan=2, pady=20)

tk.Label(root, text="Lãi suất:").grid(row=10, column=0, padx=10, pady=10)
entry_interest_rate = tk.Entry(root, state='readonly')
entry_interest_rate.grid(row=10, column=1, padx=10, pady=10)
entry_interest_rate.grid_remove()

tk.Label(root, text="Tiền lãi:").grid(row=11, column=0, padx=10, pady=10)
entry_interest_amount = tk.Entry(root, state='readonly')
entry_interest_amount.grid(row=11, column=1, padx=10, pady=10)
entry_interest_amount.grid_remove()

tk.Label(root, text="Tổng số tiền phải trả:").grid(row=12, column=0, padx=10, pady=10)
entry_total_payment = tk.Entry(root, state='readonly')
entry_total_payment.grid(row=12, column=1, padx=10, pady=10)
entry_total_payment.grid_remove()

# Tạo combobox để cập nhật hoặc xóa
tk.Label(root, text="Chọn khoản vay để cập nhật hoặc xóa:").grid(row=13, column=0, padx=10, pady=10)
combobox_update = ttk.Combobox(root)
combobox_update.grid(row=13, column=1, padx=10, pady=10)

# Thêm nút "Tìm kiếm", "Cập nhật" và "Xóa"
button_search = tk.Button(root, text="Tìm kiếm", command=search_loan)
button_search.grid(row=13, column=2, padx=10, pady=10)

button_update = tk.Button(root, text="Cập nhật", command=update_loan)
button_update.grid(row=14, column=0, columnspan=2, pady=20)

button_delete = tk.Button(root, text="Xóa", command=delete_loan)
button_delete.grid(row=14, column=2, padx=10, pady=10)

# Tạo Treeview để hiển thị dữ liệu
columns = ["Husband Name", "Wife Name", "Husband ID", "Wife ID", "Address", "Collateral Value", "Loan Amount", "Loan Duration", "Interest Rate", "Interest Amount", "Total Payment"]
tree = ttk.Treeview(root, columns=columns, show='headings')
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)
tree.grid(row=15, column=0, columnspan=3, padx=10, pady=10)

# Load dữ liệu từ excel lên
data = load_data()
data_list = data.to_dict('records')


# Vòng lặp chính
refresh_combobox()
refresh_treeview()
root.mainloop()
