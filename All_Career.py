from openpyxl import load_workbook

load_wb = load_workbook("./Analysis_Preparation _Data.xlsx", data_only=True)
load_ws = load_wb['DATA']


def PROCESS(START, END):
    try:
        LIST = list()
        get_cells = load_ws[START:END]
        for row in get_cells:
            for cell in row: LIST.append(cell.value)
        return LIST
    except TypeError as e:
        print("\nError!! Error!! Enter Valid Coordinates\n")
        print(e)
        return False

def Want(START, END): return PROCESS(START, END)

def Gender(): return PROCESS('D2', 'D421')

def Age(): return PROCESS('C2', 'C421')

def Habitation(): return PROCESS('B2', 'B421')

def DECO(number):
    deco = ['='*25]
    print(deco[number])

def INPUT():
    START = input('START : ')
    END = input('END : ')
    return START, END

def HELP():
    TEXT = """
+-------------------------+
|    <HELP>               |
| 종료 : quit             |
| 도움말 : help           |
| 전체추출 : All          |
| 성별추출 : Gender       |
| 나이추출 : Age          |
| 주거지추출 : Habitation |
| START : 시작셀 입력     |
| END : 끝셀 입력         |
+-------------------------+
"""
    print(TEXT)

def Detailed_process_All():
    START, END = INPUT()
    Want_ = Want(START, END)
    if Want_:
        Want_SET = set(Want_)
        RESULT = dict()

        for i in Want_SET:
            if type(i) == str:
                RESULT[i] = Want_.count(i)
        KEY = list(RESULT.keys())
        VALUE = list(RESULT.values())
        length = 0
        for j in VALUE:
            length += j
        for k in range(len(KEY)):
            print(f"{KEY[k]} | {VALUE[k]} | {VALUE[k] * 100 / length:0.2f}%")

def Detailed_process(TYPE):
    TYPE_dict = {'Gender': 2, 'Age': 6, 'Habitation': 5}
    TYPE_ELEMENT = {'Gender': ['남성', '여성'],
                    'Age': ['19', '18', '17', '16', '15', '14'],
                    'Habitation': ['중구', '동구', '서구', '유성구', '대덕구']}
    START, END = INPUT()
    Want_ = Want(START, END)

    if TYPE == 'Gender': SORT = Gender()
    elif TYPE == 'Age':
        Age_ = Age()
        SORT = list()
        for dummy in Age_:SORT.append(str(dummy))
    elif TYPE == 'Habitation': SORT = Habitation()
    else: return

    if Want_:
        Want_SET = set(Want_)
        Want_All = list()
        for loop in range(TYPE_dict[TYPE]): Want_All.append(list())

        for h in range(len(SORT)):
            for k in range(TYPE_dict[TYPE]):
                if SORT[h] == TYPE_ELEMENT[TYPE][k]:
                    Want_All[k].append(Want_[h])
                    break
                else: continue

        RESULT_All = list()
        KEY_All = list()
        VALUE_All = list()
        LENGTH_All = list()
        for loop in range(TYPE_dict[TYPE]): RESULT_All.append(dict())

        for x in range(TYPE_dict[TYPE]):
            for i in Want_SET:
                if type(i) == str:
                    RESULT_All[x][i] = Want_All[x].count(i)
            KEY_All.append(list(RESULT_All[x].keys()))
            VALUE_All.append(list(RESULT_All[x].values()))
            length = 0
            for j in list(RESULT_All[x].values()):
                length += j
            LENGTH_All.append(length)

        for q in range(len(TYPE_ELEMENT[TYPE])):
            print()
            print(f"\t- {TYPE_ELEMENT[TYPE][q]} -")
            for k in range(len(KEY_All[q])):
                print(f"{KEY_All[q][k]} | {VALUE_All[q][k]} | {VALUE_All[q][k] * 100 / LENGTH_All[q]:0.2f}%")
            print()


if __name__ == "__main__":
    while True:
        DECO(0)
        Command = input(' : ')
        if Command == 'help': HELP()
        elif Command == 'quit': break
        elif Command == 'All': Detailed_process_All()
        elif Command == 'Gender' or Command == 'Age' or Command == 'Habitation': Detailed_process(Command)
        else: print("\nError!! Error!! Unspecified Command\n")
    input("\nPRESS ANY BUTTON IF YOU WANT TO QUIT")