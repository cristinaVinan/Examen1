import xlwings as xw
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


from Caudal_oil.Model.Funcion import caudal_oil

HOJA = "Hoja1"
QMAX = "Qmax"
PWF = "Pwf"
PR = "Pr"
QO = "Qo"
def main():
    wb = xw.Book.caller()
    hoja = wb.sheets[HOJA]
    qmax = hoja[QMAX]
    pr = hoja[PR]
    pwf = hoja[PWF].value

    hoja[QO].value = caudal_oil(qmax,pwf,pr)


    # codigo para la gráfica

    fig, ax = plt.subplots()
    ax.plot(HOJA[PWF], HOJA[QO])
    hoja.pictures.add(fig, name="Pwf vs Qo", update=True, left=hoja.range("B16").left, pop=hoja.range("B16").pop)

if __name__ == "__main__":
    xw.Book("Caudal.xlsm").set_mock_caller()
    main()
