import openpyxl
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter


def plot_power(xlsx_path: str) -> None:
    book = openpyxl.load_workbook(xlsx_path)
    sheet = book.active

    # get column with power information from excel
    power_column = 0
    for col in range(1, sheet.max_column + 1):
        if sheet.cell(row=1, column=col).value == "Power(W)":
            power_column = col
            break

    # get column with time information from excel
    time_column = 0
    for col in range(1, sheet.max_column + 1):
        if sheet.cell(row=1, column=col).value == "Time":
            time_column = col
            break

    # axis for plot
    x = []
    y = []
    for row in range(2, sheet.max_row + 1):
        x.append(sheet.cell(row=row, column=time_column).value)
        y.append(sheet.cell(row=row, column=power_column).value / 1000)

    # customize plot
    plt.xlabel("Horário (h)")
    plt.ylabel("Potência (kW)")
    title = "Curva de geração " + str(x[0].date())
    plt.title(title)
    plt.plot(x, y)
    plt.grid()
    fmt = DateFormatter("%H")
    ax = plt.gca()
    ax.xaxis.set_major_formatter(fmt)

    # annotate max power
    ymax = max(y)
    xpos = y.index(ymax)
    xmax = x[xpos]
    annotate_text = "max=" + str(ymax) + "kW"
    ax.annotate(
        annotate_text,
        xy=(xmax, ymax),
        xytext=(xmax, ymax + 2),
        arrowprops=dict(arrowstyle="->", linewidth=1),
    )
    ax.set_ylim(0, ymax + 5)

    plt.show()
