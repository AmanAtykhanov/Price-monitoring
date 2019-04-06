import sys
import excel_forma


try:
    if __name__ == "__main__":
        pe_instance = excel_forma.excel_forma(["price_last_week.xml", "price_now_week.xml", "ls_product_names.xml"], "2.xlsx")
        pe_instance.book_forming()
except Exception as e:
    print(e.args)