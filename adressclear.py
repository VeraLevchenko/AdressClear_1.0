import openpyxl

from semantic import get_street, get_kooperativ, get_kvartal, get_snt
from number import getBuild, getKorpus, getGarage


def get_clear_adress(input_adress):
    build = "None"
    korpus = "None"
    input_adress = input_adress.replace("-", '')
    input_adress = input_adress.replace(" ", '')
    input_adress = input_adress.replace("№", '')
    input_adress = input_adress.replace("Кемеровскаяобл", '')
    input_adress = input_adress.replace("обл.Кемеровская", '')
    input_adress = input_adress.replace("Кузнецкийрайон", '')
    input_adress = input_adress.replace("Орджоникидзевский", '')
    snt = get_snt(input_adress)
    # print("СНТ ", snt)
    kooperativ = get_kooperativ(input_adress)
    # print("Гаражный кооператив ", kooperativ)
    kvartal = get_kvartal(input_adress)
    # print("Квартал", kvartal)
    type_street, street = get_street(input_adress)
    # print("Улица: ", street)
    if street != "None":
        build = getBuild(input_adress, street)
        # print("Здание:", build)
        if build != "None":
            korpus = getKorpus(input_adress)
            # print("Корпус:", korpus)
    pom = getGarage(input_adress)
    # print("Гараж:", garag)
    clear_adress = dict(
        snt=snt,
        kooperativ=kooperativ,
        kvartal=kvartal,
        type_street=type_street,
        street=street,
        build=build,
        korpus=korpus,
        pom=pom
        )
    return clear_adress
if __name__ == "__main__":
    print("Running test")
    clear_adress = get_clear_adress("пр. Октябрьский, д.16")
    print(clear_adress)
    clear_adress = get_clear_adress("ул. М. Тореза, д.109")
    print(clear_adress)
    clear_adress = get_clear_adress("ул. 40 лет ВЛКСМ, д.98")
    print(clear_adress)
    clear_adress = get_clear_adress("ул. 11 Гвардейской Армии, дом 2")
    print(clear_adress)
    clear_adress = get_clear_adress("Пионерский, д.37")
    print(clear_adress)

    # wb = openpyxl.load_workbook('D:/project_Python/518_UPR_COMP/resultAdressclearUPR.xlsx')
    # sheet = wb.active
    # max_row = sheet.max_row
    #
    # for i in range(2, max_row):
    #     data = str(sheet[i][2].value) + str(sheet[i][3].value)
    #     clearData = get_clear_adress(data)
    #     sheet[i][11].value = clearData['type_street']
    #     sheet[i][12].value = clearData['street']
    #     sheet[i][13].value = clearData['build']
    #     wb.save('D:/project_Python/518_UPR_COMP/resultAdressclearUPR.xlsx')
    #     print(clearData)

