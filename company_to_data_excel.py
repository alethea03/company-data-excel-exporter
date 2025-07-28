import tkinter as tk
from tkinter import messagebox

root = tk.Tk()
root.title("Company Data Preview + Clipboard Copy")

fields = [
    "STF Code", "Company", "Freelancer", "English Name", "Chinese Name",
    "Address", "Phone (852)", "Industry"
]

entries = {}

# Entry form
for i, field in enumerate(fields):
    tk.Label(root, text=field).grid(row=i, column=0, sticky="w")
    ent = tk.Entry(root, width=50)
    ent.grid(row=i, column=1, padx=5, pady=2)
    entries[field] = ent

# Product section
tk.Label(root, text="Product Name").grid(row=len(fields), column=0)
tk.Label(root, text="Price").grid(row=len(fields), column=1)
tk.Label(root, text="Reference Link").grid(row=len(fields), column=2)

product_entries = []

def add_product_row():
    row = len(fields) + 1 + len(product_entries)
    pname = tk.Entry(root, width=30)
    price = tk.Entry(root, width=15)
    link = tk.Entry(root, width=40)
    
    pname.grid(row=row, column=0, padx=5, pady=1)
    price.grid(row=row, column=1, padx=5, pady=1)
    link.grid(row=row, column=2, padx=5, pady=1)
    
    product_entries.append((pname, price, link))

# Add one default product row
add_product_row()

tk.Button(root, text="Add Product", command=add_product_row).grid(row=100, column=0, pady=10)

# Text preview area
preview_box = tk.Text(root, height=10, width=120, wrap="word")
preview_box.grid(row=101, column=0, columnspan=3, padx=5, pady=10)

def generate_preview():
    company_data = [entries[f].get().strip() for f in fields]
    
    if not company_data[0]:
        messagebox.showerror("Missing Info", "STF Code is required.")
        return

    price_lines = ["Product Price List 产品价格表："]
    ref_lines = ["", "Product Reference 产品参考："]
    
    for pname, price, link in product_entries:
        name = pname.get().strip()
        p = price.get().strip()
        url = link.get().strip()
        
        if name:
            price_lines.append(f"{name} = HK${p if p else '0.00'}")
        if name and url:
            ref_lines.append(f"{name} - {url}")
    
    # Format multi-line product block
    product_ref_block = "\n".join(price_lines + ref_lines)
    
    # Escape double quotes for Excel
    product_ref_block = product_ref_block.replace('"', '""')
    
    # Prepare final row (tab-separated)
    final_row = "\t".join(company_data + [f'"{product_ref_block}"'])
    
    preview_box.delete(1.0, tk.END)
    preview_box.insert(tk.END, final_row)

def clear_text():
    for entry in entries.values():
        entry.delete(0, tk.END)
        
    for pname, price, link in product_entries:
        pname.delete(0, tk.END)
        price.delete(0, tk.END)
        link.delete (0, tk.END)
        
    preview_box.delete(0, tk.END)
    
def copy_to_clipboard():
    data = preview_box.get(1.0, tk.END).strip()
    if not data:
        messagebox.showwarning("No Preview", "Click 'Generate Preview' first.")
        return
    root.clipboard_clear()
    root.clipboard_append(data)
    messagebox.showinfo("Copied", "Data copied to clipboard! Paste it in Excel.")

# Buttons
tk.Button(root, text="Generate Preview", command=generate_preview).grid(row=102, column=0, pady=5)
tk.Button(root, text="Copy to Clipboard (Excel Ready)", command=copy_to_clipboard).grid(row=102, column=1, pady=5)
tk.Button(root, text="Clear All", command=clear_text).grid(row=102, column=2,pady=5)
 
root.mainloop()
