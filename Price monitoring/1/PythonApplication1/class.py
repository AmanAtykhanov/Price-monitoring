import sys
import excel_forma


try:
    if __name__ == "__main__":
        pe_instance = excel_forma.excel_forma(["avr_sum.xml", "min_max.xml", "ls_product_names.xml"], "1.xlsx")
        pe_instance.book_forming()

except Exception as e:
    print(e.args)