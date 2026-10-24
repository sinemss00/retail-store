import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl as op
import matplotlib.pyplot as plt

products = list()
sales = list()
categories = dict()
valid_codes = set()

root = tk.Tk()
root.title("Вход")
root.geometry("300x250")

tk.Label(root, text="Потребителско име:").grid(row = 0, column = 0, sticky="e", padx=10, pady=10)
e_name = tk.Entry(root, width=22)
e_name.grid(row=0, column = 1, padx=10)

tk.Label(root, text="Парола:").grid(row=1, column=0, sticky="e", padx=10, pady=10)
e_pass = tk.Entry(root, show="*", width=22)
e_pass.grid(row=1, column=1, padx=10)

login_btn = tk.Button(root, text="Вход", width=22, height=2, bg="lightgreen")
login_btn.grid(row = 2, column = 0, columnspan=2, pady=20)

login_lbl = tk.Label(root, text="", width=26)
login_lbl.grid(row=3, column = 0, columnspan = 2, pady = 10)

def login():
    user = e_name.get().strip()
    pas = e_pass.get().strip()
    if user == "admin" and pas == "1234":
        login_lbl.config(text = f"Успешен вход като {user}!", fg="green")
        load_data()
        root.destroy()
        main_menu()
    else:
        login_lbl.config(text="Моля, въведете валидни данни!", fg="red")

login_btn.config(command=login)

def load_data():
    wb_cat = op.load_workbook("categories.xlsx")
    ws_cat = wb_cat.active

    for cat_id, cat_name in ws_cat.iter_rows(min_row = 2, values_only=True):
        if cat_id:
            categories[str(cat_id)] = str(cat_name)

    wb_prod = op.load_workbook("products.xlsx")
    ws_prod = wb_prod.active

    for prod_id, cat_id, model, size, price in ws_prod.iter_rows(min_row = 2, values_only=True):
        if prod_id:
            products.append({
                'id':str(prod_id),
                'cat_id':str(cat_id),
                'model':str(model),
                'size':str(size),
                'price':float(price)
            })
            valid_codes.add(str(prod_id))
            
    wb_sales = op.load_workbook("sales.xlsx")
    ws_sales = wb_sales.active

    for prod_id, date, qty in ws_sales.iter_rows(min_row=2, values_only=True):
        if prod_id:
            sales.append({
                'id':str(prod_id),
                'date':str(date),
                'qty':int(qty)
            })

def reports_menu():
    reports_root = tk.Tk()
    reports_root.title("Справки и графики")
    reports_root.geometry("500x500")

    tk.Label(reports_root, text="Финансови справки за магазина", font="bold").grid(row=0, column=0, columnspan=2, pady=15)
    
    btn_calc = tk.Button(reports_root, text="Изчисли оборота", width=22, bg = "lightblue")
    btn_calc.grid(row=1, column=0, padx=15, pady=10)
    
    btn_chart = tk.Button(reports_root, text="Покажи графика", width=22, bg="lightblue")
    btn_chart.grid(row=1, column=1, padx=15, pady=10)
    
    listbox_rep = tk.Listbox(reports_root, width=55, height=12)
    listbox_rep.grid(row=2, column=0, columnspan=2, padx=15, pady=10)
    
    btn_back = tk.Button(reports_root, text="Назад към менюто", bg="lightgreen", width=15)
    btn_back.grid(row=3, column=0, padx=15, pady=20, sticky="w")
    
    btn_exit = tk.Button(reports_root, text="Изход", bg="red", width=15)
    btn_exit.grid(row=3, column=1, padx=15, pady=20, sticky="e")

    btn_line = tk.Button(reports_root, text = "Oborot po dati", width = 22, bg = "lightblue")
    btn_line.grid(row = 4, column = 0, columnspan = 2, padx = 15, pady = 10)

    category_sum = dict()

    def calculate_reports():
        listbox_rep.delete(0, tk.END)
        category_sum.clear()

        for cat_name in categories.values():
            category_sum[cat_name] = 0.0

        for s in sales:
            for p in products:
                if p['id'] == s['id']:
                    cat_name = categories.get(p['cat_id'], "Неизвестна")
                    category_sum[cat_name] += s['qty'] * p['price']

        listbox_rep.insert(tk.END, "ОБОРОТ ПО КАТЕГОРИИ СТОКИ:")
        for cat_name, total_sum in category_sum.items():
            listbox_rep.insert(tk.END, f" {cat_name}: {total_sum:.2f} евро")

    def show_chart():
        calculate_reports()
        
        if sum(category_sum.values()) == 0:
            messagebox.showwarning("Внимание", "Няма регистрирани продажби за показване на графика!")
            return
        
        bars = plt.bar(list(category_sum.keys()), list(category_sum.values()), color="lightblue")
        plt.bar_label(bars, labels=[f"{val:.2f} евро" for val in category_sum.values()])
        
        plt.title("Общ оборот по категории стоки (в евро)")
        plt.xlabel("Категории")
        plt.ylabel("Оборот (евро)")
        plt.tight_layout()
        plt.show()
    
    def show_line():
        calculate_reports()

        if sum(category_sum.values()) == 0:
            messagebox.showwarning("Vnimanie", "Nyama registrirani prodajbi za pokazvane!")
            return
        
        for s in sales:
            for p in products:
                if p['id'] == s['id']:
                    y = p['price'] * s['qty']
                    x = s['date']

                    plt.plot_label(f"{val:.2f} евро" for val in category_sum.values())
                    plt.title("Obsht oborot po dati")
                    plt.xlabel("Oborot")
                    plt.ylabel("Dati")
                    plt.tight_layout()
                    plt.show()

    def go_back():
        reports_root.destroy()
        main_menu()

    btn_calc.config(command=calculate_reports)
    btn_chart.config(command=show_chart)
    btn_back.config(command=go_back)
    btn_exit.config(command=reports_root.quit)
    btn_line.config(command = show_line)

    reports_root.mainloop()

def view_menu():
    view_root = tk.Tk()
    view_root.title("Наличности в магазина")
    view_root.geometry("600x500")

    columns = ('Код', 'Категория', 'Модел', 'Размер', 'Цена')
    tree = ttk.Treeview(view_root, columns=columns, show="headings", height=12)
    tree.heading('Код', text = 'Код')
    tree.heading('Категория', text = 'Категория')
    tree.heading('Модел', text = 'Модел')
    tree.heading('Размер', text = 'Размер')
    tree.heading('Цена', text = 'Цена')

    tree.column('Код', width=80, anchor="center")
    tree.column('Категория', width=120, anchor="center")
    tree.column('Модел', width=180, anchor="center")
    tree.column('Размер', width=80, anchor="center")
    tree.column('Цена', width=100, anchor="center")

    tree.grid(row=0, column=0, columnspan=2, padx=20, pady=15)
    btn_back = tk.Button(view_root, text="Назад към главно меню", width=20, height=3, bg = "lightgreen")
    btn_exit = tk.Button(view_root, text="Изход", width=20, height=3, bg = "red")
    btn_back.grid(row=1, column=0, padx=20, pady=15)
    btn_exit.grid(row=1, column=1, padx=20, pady=15)

    def fill_table():
        for p in products:
            cat_name = categories.get(p['cat_id'], "Неизвестна")
            tree.insert("", tk.END, values=(p['id'], cat_name, p['model'], p['size'], f"{p['price']:.2f} евро"))

    def go_back():
        view_root.destroy()  
        main_menu()     

    fill_table()
    btn_back.config(command=go_back)
    btn_exit.config(command=view_root.quit) 

    view_root.mainloop()

def search_menu():
    search_root = tk.Tk()
    search_root.title("Търсене и Филтриране")
    search_root.geometry("470x450")

    lbl_search_cat = tk.Label(search_root, text="Въведете код на категория (1-4):")
    e_cat_id = tk.Entry(search_root, width=10)
    btn_search_cat = tk.Button(search_root, text="Търси категория", width = 17, bg="lightblue")
    lbl_cat_result = tk.Label(search_root, text="Резултат: ")

    lbl_search_cat.grid(row=0, column=0, sticky="w", padx=15, pady=10)
    e_cat_id.grid(row=0, column=1, sticky="w", padx=5, pady=10)
    btn_search_cat.grid(row=0, column=2, sticky="w", padx=5, pady=10)
    lbl_cat_result.grid(row=1, column=0, columnspan=3, sticky="w", padx=15, pady=5)

    lbl_filter_price = tk.Label(search_root, text="Покажи стоки с цена над (евро):")
    e_price_limit = tk.Entry(search_root, width=10)
    btn_filter = tk.Button(search_root, text="Филтрирай", width = 17, bg="lightblue")
    lbl_listbox = tk.Label(search_root, text="Резултати от филтрирането:")
    listbox = tk.Listbox(search_root, width=60, height=10)

    lbl_filter_price.grid(row=2, column=0, sticky="w", padx=15, pady=10)
    e_price_limit.grid(row=2, column=1, sticky="w", padx=5, pady=10)
    btn_filter.grid(row=2, column=2, sticky="w", padx=5, pady=10)
    lbl_listbox.grid(row=3, column=0, sticky="w", padx=15, pady=5)
    listbox.grid(row=4, column=0, columnspan=3, padx=15, pady=5)
    
    btn_save_file = tk.Button(search_root, text="Запиши в текстов файл", width = 20, bg="lightblue")
    btn_back = tk.Button(search_root, text="Назад към менюто", width=20, bg="lightgreen")
    btn_exit = tk.Button(search_root, text="Изход", width=20, bg="red")

    btn_save_file.grid(row=5, column=0, columnspan=3, pady=10)
    btn_back.grid(row=6, column=0, padx=15, pady=20, sticky="w")
    btn_exit.grid(row=6, column=2, padx=15, pady=20, sticky="e")

    filtered_products = list()

    def search_category():
        search_key = e_cat_id.get().strip()
        if search_key in categories:
            lbl_cat_result.config(text=f"Резултат: Код {search_key} - {categories[search_key]}", fg="green")
        else:
            lbl_cat_result.config(text="Резултат: Няма такава категория!", fg="red")
    
    def filter_price():
        filtered_products.clear()
        listbox.delete(0, tk.END)
        
        user_input = e_price_limit.get().strip()
        if not user_input:
            messagebox.showwarning("Внимание", "Моля, въведете цена!")
            return
            
        if not user_input.replace('.', '', 1).isdigit():
            messagebox.showerror("Грешка", "Моля, въведете валидно число за цена!")
            return
            
        limit = float(user_input)
        
        for p in products:
            if p['price'] > limit:
                filtered_products.append(p)
                info_text = f"Код: {p['id']} | {p['model']} ({p['size']}) - {p['price']:.2f} евро"
                listbox.insert(tk.END, info_text)
                
        if len(filtered_products) == 0:
            listbox.insert(tk.END, "Няма намерени стоки с по-висока цена.")

    def save_to_file():
        if len(filtered_products) == 0:
            messagebox.showwarning("Внимание", "Списъкът е празен!")
            return
            
        with open("filtered_products.txt", "w", encoding="utf-8") as f:
            f.write("СПИСЪК С ФИЛТРИРАНИ СТОКИ\n")
            for p in filtered_products:
                f.write(f"Код: {p['id']} | Модел: {p['model']} | Размер: {p['size']} | Цена: {p['price']:.2f} евро\n")
                
        messagebox.showinfo("Успех", "Данните са записани във 'filtered_products.txt'!")

    def go_back():
        search_root.destroy()
        main_menu()

    btn_search_cat.config(command=search_category)
    btn_filter.config(command=filter_price)
    btn_save_file.config(command=save_to_file)
    btn_back.config(command=go_back)
    btn_exit.config(command=search_root.quit)

    search_root.mainloop()

def add_menu():
    add_root = tk.Tk()
    add_root.title("Добавяне на данни")
    add_root.geometry("400x500")

    tk.Label(add_root, text="ДОБАВЯНЕ НА НОВА СТОКА", font="bold").grid(row=0, column=0, columnspan=2, pady=10)
    
    tk.Label(add_root, text="Код стока:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
    e_new_id = tk.Entry(add_root, width=15)
    e_new_id.grid(row=1, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Код категория (1-4):").grid(row=2, column=0, sticky="w", padx=10, pady=5)
    e_new_cat = tk.Entry(add_root, width=15)
    e_new_cat.grid(row=2, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Модел/Име:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
    e_new_model = tk.Entry(add_root, width=15)
    e_new_model.grid(row=3, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Размер (S, M, L, XL):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    e_new_size = tk.Entry(add_root, width=15)
    e_new_size.grid(row=4, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Цена (евро):").grid(row=5, column=0, sticky="w", padx=10, pady=5)
    e_new_price = tk.Entry(add_root, width=15)
    e_new_price.grid(row=5, column=1, sticky="w", padx=10)

    btn_add_prod = tk.Button(add_root, text="Добави стока", bg="lightblue", width=20)
    btn_add_prod.grid(row=6, column=0, columnspan=2, pady=10)

    tk.Label(add_root, text="РЕГИСТРИРАНЕ НА ПРОДАЖБА", font="bold").grid(row=7, column=0, columnspan=2, pady=15)
    
    tk.Label(add_root, text="Код на продадена стока:").grid(row=8, column=0, sticky="w", padx=10, pady=5)
    e_sale_id = tk.Entry(add_root, width=15)
    e_sale_id.grid(row=8, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Количество:").grid(row=9, column=0, sticky="w", padx=10, pady=5)
    e_sale_qty = tk.Entry(add_root, width=15)
    e_sale_qty.grid(row=9, column=1, sticky="w", padx=10)

    tk.Label(add_root, text="Дата (ГГГГ-ММ-ДД):").grid(row=10, column=0, sticky="w", padx=10, pady=5)
    e_sale_date = tk.Entry(add_root, width=15)
    e_sale_date.grid(row=10, column=1, sticky="w", padx=10)
    e_sale_date.insert(0, "2026-05-15")

    btn_add_sale = tk.Button(add_root, text="Регистрирай продажба", bg="lightblue", width=25)
    btn_add_sale.grid(row=11, column=0, columnspan=2, pady=10)

    btn_back = tk.Button(add_root, text="Назад към менюто", bg="lightgreen", width=18)
    btn_back.grid(row=12, column=0, pady=15, padx=20, sticky="w")
    
    btn_exit = tk.Button(add_root, text="Изход", bg="red", width=18)
    btn_exit.grid(row=12, column=1, pady=15, padx=20, sticky="e")

    def add_product():
        nid = e_new_id.get().strip()
        ncat = e_new_cat.get().strip()
        nmodel = e_new_model.get().strip()
        nsize = e_new_size.get().strip()
        nprice = e_new_price.get().strip()

        if not (nid and ncat and nmodel and nsize and nprice):
            messagebox.showwarning("Грешка", "Моля, попълнете всички полета за стоката!")
            return

        if not nid.isdigit():
            messagebox.showerror("Грешка", "Кодът на стоката трябва да бъде цяло число!")
            return

        if not ncat.isdigit():
            messagebox.showerror("Грешка", "Кодът на категорията трябва да бъде цяло число!")
            return

        if nid in valid_codes:
            messagebox.showerror("Грешка", f"Продукт с код {nid} вече съществува!")
            return

        if not nprice.replace('.', '', 1).isdigit():
            messagebox.showerror("Грешка", "Цената трябва да бъде валидно число (напр. 25.50)!")
            return

        if ncat not in categories:
            answer = messagebox.askquestion("Нова категория", f"Категория с код {ncat} не съществува. Искате ли да я създадете?")
            if answer == 'yes':
                categories[ncat] = f"Група {ncat}"
                messagebox.showinfo("Успех", f"Категория {ncat} беше добавена успешно!")
            else:
                return

        products.append({
            'id': nid, 'cat_id': ncat, 'model': nmodel, 'size': nsize, 'price': float(nprice)
        })
        valid_codes.add(nid)
        messagebox.showinfo("Успех", f"Продуктът '{nmodel}' е добавен успешно!")
        
        e_new_id.delete(0, tk.END); e_new_cat.delete(0, tk.END)
        e_new_model.delete(0, tk.END); e_new_size.delete(0, tk.END); e_new_price.delete(0, tk.END)
        
    def add_sale():
        sid = e_sale_id.get().strip()
        sqty = e_sale_qty.get().strip()
        sdate = e_sale_date.get().strip()

        if not (sid and sqty and sdate):
            messagebox.showwarning("Грешка", "Попълнете всички полета за продажбата!")
            return

        if sid not in valid_codes:
            messagebox.showerror("Грешка", f"Продукт с код {sid} не съществува! Не може да се направи продажба.")
            return

        if not sqty.isdigit():
            messagebox.showerror("Грешка", "Количеството трябва да бъде цяло число!")
            return

        sales.append({
            'id': sid, 'date': sdate, 'qty': int(sqty)
        })
        messagebox.showinfo("Успех", "Продажбата е добавена успешно!")
        e_sale_id.delete(0, tk.END); e_sale_qty.delete(0, tk.END)

    def go_back():
        add_root.destroy()
        main_menu()

    btn_add_prod.config(command=add_product)
    btn_add_sale.config(command=add_sale)
    btn_back.config(command=go_back)
    btn_exit.config(command=add_root.quit)

    add_root.mainloop()

def main_menu():
    main_root = tk.Tk()
    main_root.title("Главно меню")
    main_root.geometry("500x500")

    tk.Label(main_root, text="Онлайн магазин за дрехи ЕЛЕА", font="bold").pack()
    view_btn = tk.Button(main_root, text="Преглед на наличности", width=22, height=3, bg="lightgreen")
    search_btn = tk.Button(main_root, text="Търсене и филтриране", width=22, height=3, bg="lightgreen")
    add_btn = tk.Button(main_root, text="Добавяне на нови данни", width=22, height=3, bg="lightgreen")
    reports_btn = tk.Button(main_root, text="Справки и графики", width=22, height=3, bg="lightgreen")
    exit_btn = tk.Button(main_root, text="Изход", command=main_root.quit, width=22, height=3, bg="red")

    view_btn.pack(pady=15)
    search_btn.pack(pady=15)
    add_btn.pack(pady=15)
    reports_btn.pack(pady=15)
    exit_btn.pack(pady=15)

    def open_view():
        main_root.destroy()
        view_menu()
    view_btn.config(command=open_view)

    def open_search():
        main_root.destroy()
        search_menu()
    search_btn.config(command=open_search)

    def open_add():
        main_root.destroy()
        add_menu()
    add_btn.config(command = open_add)

    def open_reports():
        main_root.destroy()
        reports_menu()
    reports_btn.config(command = open_reports)

    main_root.mainloop()

root.mainloop()