import sys
import excel_forma


try:
    if __name__ == "__main__":
        pe_instance = excel_forma.excel_forma(["ls_price_first_week.xml", 
                                               "ls_price_now_week.xml", 
                                               "ls_limit_price.xml",
                                               "ls_product_names.xml"], 
                                              "3.xlsx")
        pe_instance.book_forming()
except Exception as e:
    print(e.args)