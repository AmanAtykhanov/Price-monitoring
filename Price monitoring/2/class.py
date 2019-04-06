import sys
import XMLReader as xr
import excel_forma

try:
    if __name__ == "__main__":
        pe_instance = excel_forma.excel_forma(sys.argv[1:-1], sys.argv[-1])
        pe_instance.page2_forming()
except Exception as e:
    print(e.args)