from openpyxl import load_workbook

def DataCollection(START, END, TorF):
    load_WB = load_workbook("./Analysis_Preparation _Data.xlsx", data_only=True)
    load_WS = load_WB['DATA']

    usingDATA = list()
    dummyDATA = load_WS[START:END]
    for i in dummyDATA:
        for j in i:
            usingDATA.append(j.value)

    if TorF:
        return usingDATA, dummyDATA
    else:
        return usingDATA
