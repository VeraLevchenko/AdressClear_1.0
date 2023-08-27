def generate_list_build(street, max_number_build): # генерирует список сочетаний для поиска зданий
    liters = ["А", "Б", "В", "Г", "Д", "Е", "Е", "Ж", "З", "а", "б", "в", "г", "д", "ж", "з", ""]
    names = ["д", "д.", "дом", "здание", "зд.", "дом.", "строение", ""]
    list_build_data_liter = [f"{street},{name}{i}{liter}" for name in names for i in range(max_number_build, 0, -1) for liter in liters]
    return(list_build_data_liter, liters)

def generate_list_korpus(max_number): # генерирует список сочетаний для поиска зданий
    liters = ["А", "Б", "В", "Г", "Д", "Е", "Е", "Ж", "З", "а", "б", "в", "г", "д", "ж", "з", ""]
    names = ["К", "корпус", "корп", "к.", "блок", "корп.", "Корпус"]
    # liters = ["А", "Б", ""]
    # names = ["к.", "корпус"]
    list_build_data_liter = [f"{name}{i}{liter}" for name in names for i in range(max_number, 0, -1) for liter in liters]
    return(list_build_data_liter, liters)

def generate_list_garag(max_number): # генерирует список сочетаний для поиска помещений (гаражей)
    liters = ["А", "Б", "В", "Г", "Д", "Е", "Е", "Ж", "З", "а", "б", "в", "г", "д", "ж", "з", ""]
    names = ["помещение(гараж)", "помещение", "пом.", "гаражпод", "гаража", "под", "пом", "гараж", "номербокса(ячейки)",
             "номер", "бокс", "гар.", "г.", "гараж.", "гараж.", "кв.", "кв", "квартира", "кварт."]
    list_garag_data_liter = [f"{name}{i}{liter}" for name in names for i in range(max_number, 0, -1) for liter in liters]
    return(list_garag_data_liter, liters)

# def generate_list_type_street(): # генерирует список сочетаний для поиска зданий
#     types_street = ["Улица", "ул.", "ул", "Ул.", "Ул.", "улица"]
#     names = ["д", "д.", "дом", "здание", "зд.", "дом.", "строение", ""]
#     list_build_data_liter = [f"{street},{name}{i}{liter}" for name in names for i in range(max_number_build, 0, -1) for liter in liters]
#     return(list_build_data_liter, liters)

def get_number_build(input_adress, generate_adress ):  # возвращает найденную строку со значением улицы и здания "Кирова,16Г"
    number_build = "None"
    for number_build in generate_adress:
        index = input_adress.find(number_build)
        if index != -1:
            break
        else:
            number_build = "None"
    return number_build

def make_num_build(list_build_data_liter, liters, number_build, max_number):  # Возвращает номер здания "16Г"
    index = list_build_data_liter.index(number_build) + 1
    # -------------------------------------------------------------------------------------
    if index % len(liters) == 0:
        name_rez = index // len(liters)
    else:
        name_rez = index // len(liters) + 1
    if name_rez == 1:
        number = max_number
    else:
        if name_rez % max_number == 0:
            number = 1
        else:
            number = max_number + 1 - name_rez % max_number
    # ---------------------------------------------------------------------------------------------
    if index <= len(liters):
        liter_rez = liters[index - 1]
    else:
        if index % len(liters) == 0:
            liter_rez = liters[index % len(liters) - 1]
        else:
            liter_rez = liters[index - len(liters) * (name_rez - 1) - 1]
    build = f"{number}{liter_rez}"
    return build


def getBuild(input_adress, street):     # Главная функция - возвращает номер здания
    max_number_build = 200
    street = street.replace("-", '')
    street = street.replace(" ", '')
    list_build_data_liter, liters = generate_list_build(street, max_number_build)
    number_build = get_number_build(input_adress, list_build_data_liter)
    number = number_build
    if number_build != "None":
        number = make_num_build(list_build_data_liter, liters, number_build, max_number_build)
        # print("Адрес здания: ", number_build)
        # print("Номер: ", number)
    number = str(number).upper()
    return number

def getKorpus(input_adress):     # Главная функция - возвращает номер корпуса
    max_number_korpus = 200
    list_build_data_liter, liters = generate_list_korpus(max_number_korpus)
    number_build = get_number_build(input_adress, list_build_data_liter)
    number = number_build
    if number_build != "None":
        number = make_num_build(list_build_data_liter, liters, number_build, max_number_korpus)
        # print("Адрес здания: ", number_build)
        # print("Номер: ", number)
    return number


def getGarage(input_adress):     # Главная функция - возвращает номер корпуса
    max_number_garag = 3000
    list_build_data_liter, liters = generate_list_garag(max_number_garag)
    number_build = get_number_build(input_adress, list_build_data_liter)
    number = number_build
    if number_build != "None":
        number = make_num_build(list_build_data_liter, liters, number_build, max_number_garag)
        # print("Адрес здания: ", number_build)
        # print("Номер: ", number)
    return number
