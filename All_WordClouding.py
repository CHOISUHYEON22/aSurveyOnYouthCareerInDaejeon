from data_collection import DataCollection
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as mpl


if __name__ == '__main__':

    usingDATA = DataCollection("E2", "G421", False)

    # 자료 분석(빈도수)
    COUNT = Counter(usingDATA)
    word_count = list()
    words = list()

    for p, q in COUNT.most_common(21):
        dic = {'word': p, 'count': q}
        word_count.append(dic)
        words.append(dic['word'])
    length = 1160

    NumOfCount = 1
    print("=" * 40)
    for k in word_count:
        try:
            print(f"{NumOfCount}. {k['word']:>6} | {k['count']:<3} | {k['count'] * 100 / length:0.2f}%")
            NumOfCount = NumOfCount + 1
        except TypeError:
            continue
    print("=" * 40)
    print("총 단어 수 : 총 {0}개".format(length))
    print('\n\n')

    # 자료 분석(Word Clouding)
    count = 0
    usingDATAinWC = ""
    for h in usingDATA:
        try:
            if count <= 50:
                usingDATAinWC = usingDATAinWC + " " + str(h)
                count += 1
        except TypeError: continue
    WC = WordCloud(font_path="./gulim.ttc", background_color='white', max_font_size=100).generate(usingDATAinWC)
    mpl.figure(figsize=(12, 12))
    mpl.imshow(WC, interpolation='bilinear')
    mpl.axis("off")
    mpl.show()

    input("\nPRESS ANY BUTTON IF YOU WANT TO QUIT")