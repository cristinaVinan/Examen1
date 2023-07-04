import xlwings as xw
import pandas as pd

def main():
    wb = xw.Book.caller()
    HOJA = wb["Hoja1"]
    QMAX = HOJA["Qmax"]
    PWF = HOJA["Pwf"]
    PR = HOJA["Pr"]
    QO = HOJA["Qo"]



if __name__ == "__main__":
    xw.Book("Caudal.xlsm").set_mock_caller()
    main()
