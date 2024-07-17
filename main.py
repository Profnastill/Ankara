import numpy
import numpy as np
import xlwings as xw
import pandas as pd
from icecream import ic

beton_type = ["В3,5", "В5", "В7,5", "В10", "В12,5", "В15", "В20", "В25", "В30", "В35", "В40", "В45", "В50", "В55",
              "В60", "В70", "B80", "B90", "В100"]

beton_Rbn = [2.7, 3.5, 5.5, 7.5, 9.5, 11, 15, 18.5, 22, 25.5, 29, 32, 36, 39.5, 43, 50, 57, 64, 71]

class_arm = ["A240", "A400", "A500", "B500", ]
Rsw_arm = [170, 280, 300, 300]
Rs_arm = [210, 340, 435, 415]
Rs_n=[240,400,500,500]
data_R_arm = pd.DataFrame(index=class_arm, data=list(zip(Rsw_arm, Rs_arm,Rs_n)), columns=["Rsw", "Rs","Rsn"])
print(data_R_arm)
data_Rb = pd.DataFrame(index=beton_type, data=beton_Rbn, columns=["Rb"])
print(data_Rb)
armatura = ["А400", "A500"]


def printer(fun):
    def _pr(*args, **kwargs):
        text = "sda"
        print(f"{text}={fun}")
        return


class Usilia:
    def __init__(self, Nan, Nan_tot, Van, Van_tot):
        super().__init__()
        self.Nan = Nan  # Усилие в одном анкере
        self.Nan_tot = Nan_tot  # продольное усилие в группе анкеров
        self.Van = Van  # 7.2.1.1 Усилие горизонтальное в одном анкере
        self.Van_tot = Van_tot  # Cуммарное усилие в группе
        ic(self.Nan, self.Nan_tot, self.Van, self.Van_tot)


class Ankern_Krep(Usilia):
    def __init__(self, Nan, Nan_tot, Van, Van_tot, data: pd.DataFrame, data1: pd.DataFrame):
        super().__init__(Nan, Nan_tot, Van, Van_tot)

        """
        :param data:  Таблица с характеристиками материалов
        :param data1: Таблица усилий и расположения анкеров
        """
        self.material_st = data.loc["Класс болтов", "Параметр"]
        self.material_osn = data.loc["Основание", "Параметр"]
        self.d = data.loc["Диаметр Анкеров", "Параметр"] * 100  # мм2 диаметр анкера
        self.N_n_p = 0  # нормативное значение силы сопротивления анкера сцепления с основанием (по контакту)
        self.gamma_bt = 1.5
        self.psi_c = 0  # коэффициент прочность бетонного основания, в зависимости от класса бетона на сжатие, от типа анкера по ТП.
        """ self.gamma_Np уточнение"""
        self.gamma_Np = 1  # - коэффициент условий работы анкера, принимаемый в зависимости от типа и марки анкера по ТП  это требует уточнения
        self.gamma_Nc = 1  # коэффициент условий работы анкера, принимаемый в зависимости от типа и марки анкера по ТП  это требует уточнения
        self.Rbn = data_Rb.loc[self.material_osn, "Rb"]  # МПА
        self.k1 = {"Возможны": 7.9, "Не возможно": 11.3}[
            data.loc["Образование трещин", "Параметр"]]  # Возможно или нет образование трещин
        self.k3= {"Возможны": 2, "Не возможно": 2.8}[
            data.loc["Образование трещин", "Параметр"]]  # Возможно или нет образование трещин
        self.e_N_1 = 1  # эксцентриситет растягивающей силы относительно центра тяжести анкерной группы для соответствующего направления (см. 6.8)
        self.e_N_2 = 1  # эксцентриситет  для одиночной равен 1
        self.Es = 206000  # Модуль упругости арматуры
        self.h = (data.loc["Толщина основания", "Параметр"])
        self.As = (data.loc["Диаметр Анкеров", "Параметр"] / 2) ** 2 * np.pi  # Площадь анкера
        self.Rs = data_R_arm.loc[self.material_st, "Rs"]
        self.Rsw = data_R_arm.loc[self.material_st, "Rsw"]
        self.Rsn = data_R_arm.loc[self.material_st, "Rsn"]
        self.rasstoyani(data1, data)
        self.usl_Procn()

    def rasstoyani(self, data: pd.DataFrame, data1: pd.DataFrame):

        data.drop_duplicates(inplace=True, ignore_index=True)
        self.count_ank = data.count().size * 2
        if self.count_ank > 1:
            data.sort_values(by="x", inplace=True)
            s_x_min = data["x"].diff().dropna().min()
            # s_x_max = data["x"].diff().dropna().max()  # Получение расстояний между болтами минима и максим
            s_x_max = data["x"].max() - data["x"].min()  # Максимальное расстояние между крайними болтами
            data.sort_values(by="y", inplace=True)
            # s_y_min = data["y"].diff().dropna().min()
            # s_y_max = data["y"].diff().dropna().max()
            s_y_max = data["y"].max() - data["y"].min()  # Максимальное расстояние между крайними болтами

        print(data)
        self.c_1_x = data1.loc["c_1_x", "Параметр"]
        self.c_2_x = data1.loc["c_2_x", "Параметр"]
        self.c_1_y = data1.loc["c_1_y", "Параметр"]
        self.c_2_y = data1.loc["c_2_y", "Параметр"]

        self.h_ef = max(self.c_1_x / 1.5, self.c_2_x / 1.5, s_y_max / 3)  # эффективная глубина анкеровки
        self.s_cr_N = 3 * self.h_ef  # Максимально расстояние между болтами где нет реакции
        self.c_cr_N = 1.5 * self.h_ef  # Максимально краевое между болтами где нет реакции
        self.c = min(self.c_1_x, self.c_1_y, self.c_2_x, self.c_2_y)

        self.s_y_max = min(s_y_max, self.s_cr_N)
        self.s_x_max = min(s_x_max, self.s_cr_N)

        print(data)

    @property
    def A_c_N(self):

        s_y_max = min(self.s_y_max, self.s_cr_N)
        s_x_max = min(self.s_x_max, self.s_cr_N)
        A_c_N=0
        if self.count_ank == 1 and self.c_1_x <= self.c_cr_N:
            A_c_N = (self.c_1_x + 0.5 * self.s_cr_N) * self.s_cr_N
        elif self.count_ank == 2 and self.c_1_x <= self.c_cr_N:
            A_c_N = (self.c_1_x + s_x_max + 0.5 * self.s_cr_N) * self.s_cr_N
        elif self.count_ank >= 4 and self.c_1_x <= self.c_cr_N and self.c_1_y <= self.c_cr_N and s_x_max <= self.s_cr_N and s_y_max <= self.s_cr_N:
            A_c_N = (self.c_1_x + s_x_max + 0.5 * self.s_cr_N) * (self.c_1_y + s_y_max + 0.5 * self.s_cr_N)
        ic(A_c_N)
        return A_c_N

    @property
    def N_ult_s(self):
        """
        По прочности стали анкера на растяжение
        :return:
        """
        if self.material_st in ["A400", "A500"]:
            gamm_Ns = 1.25
        else:
            gamm_Ns = 1
        N_n_s = self.As * self.Rs
        return N_n_s / gamm_Ns

    @property
    def N_ul_p(self):

        N_ult_p = self.N_n_p * self.psi_c / self.gamma_bt / self.gamma_Np  # формула 7.6
        return N_ult_p

    @property
    def N0_n_c(self):
        N0_n_c = self.k1 * pow(self.Rbn, 0.5) * pow(self.h_ef, 1.5)
        return N0_n_c

    @property
    def A_c_N_0(self):
        A_c_N_0 = self.s_cr_N * self.s_cr_N
        return A_c_N_0

    @property
    def psi_s_N(self):
        psi_s_N = 0.7 + 0.3 * self.c / self.c_cr_N
        return psi_s_N

    @property
    def psi_re_N(self):
        psi_re_N = 0.5 * self.h_ef
        return psi_re_N

    @property
    def psi_ec_N(self):

        si_ec_N = (1 / (1 + 2 * self.e_N_1 / self.s_cr_N) * (1 / 1 + 2 * self.e_N_2 / self.s_cr_N))
        if si_ec_N < 1:
            si_ec_N
        else:
            si_ec_N = 1
        return si_ec_N

    @property
    def N_ult_c(self):  #
        N_ult_c = self.N0_n_c / self.gamma_bt / self.gamma_Nc * self.A_c_N / self.A_c_N_0 * self.psi_s_N * self.psi_re_N * self.psi_ec_N  # формула 7.9
        return N_ult_c

    @property
    def N_nc_sp(self):  #
        gamma_Nc = 1
        N_ult_c = self.N0_n_c / self.gamma_bt / gamma_Nc * self.A_c_N / self.A_c_N_0 * self.psi_s_N * self.psi_re_N * self.psi_ec_N
        return N_ult_c

    @property
    # @printer

    def N_ult_sp(self):
        self.gamma_N_sp = 1  # Для стальных анкеров
        self.h_min = self.h_ef + 30  # Для стальных анкеров
        psi_h_sp = pow(self.h / self.h_min, 2 / 3)
        if psi_h_sp <= pow(2 * self.h_ef / self.h_min, 2 / 3):
            None
        else:
            psi_h_sp = pow(2 * self.h_ef / self.h_min, 2 / 3)
        if self.h / self.h_ef >= 2.0:  # Только для стальных анкеров
            c_cr_sp = 1 * self.h_ef
        elif 1.3 < self.h / self.h_ef < 2:
            c_cr_sp = 4.6 * self.h_ef - 1.8 * self.h
        elif self.h / self.h_ef <= 1.3:
            c_cr_sp = 2.3 * self.h_ef
        self.s_cr_sp = 2 * c_cr_sp

        a=self.s_cr_N# Временное присвоение
        b=self.c_cr_N

        self.s_cr_N = self.s_cr_sp  #Тут происходит перезапить коэффициентов для вычесления  Nult_sp в Asn Asn0
        self.c_cr_N = c_cr_sp #Тут происходит перезапить коэффициентов для вычесления  Nult_sp у учета в Asn Asn0
        N_ult_sp = self.N_nc_sp / self.gamma_N_sp * psi_h_sp
        self.s_cr_N=a# Тут возвращаем прежние значения
        self.c_cr_N=b
        return N_ult_sp

    @property
    def V_ult_s(self):
        gamma_vs=1
        V_n_s=0.5*self.Rsn**np.pi*(self.d**2)/4
        V_ult_s=V_n_s/gamma_vs
        return  V_ult_s

    @property
    def V_ult_cp(self):
        gamma_V_cp=1
        k=2# Для стальных анкеров из арматуры

        V_ult_cp=k*self.N_ult_c/gamma_V_cp
        return  V_ult_cp

    def V_ult_c(self):

        a=0.1*(l_f/c1)**0.5
        V_0_nc=self.k3*(self.d)**a*(l_f)*betta*pow(self.Rbn,0.5)*c_1**1.5
        gamma_bt=1.5
        V_ult_c=V_0_nc/gamma_bt/gamma_vc *A_c_V/A0_c_V* psi_s_V*psi_h_V*psi_a_V*psi_ec_V*psi_re_V#7.4.3

    def usl_Procn(self):
        # Nan = self.Rs * self.As
        proverka_usl = []
        if self.Nan <= self.N_ult_s:  # по стали анкера 7.1.1.1
            text = "Проверка по стали анкера п.7.1.1.1 выполняется"
            print(text)
        else:
            text = "Проверка по стали анкера п.7.1.1.1 НЕ выполняется!!"

        proverka_usl.append(text)

        if self.Nan <= self.N_ul_p:  # Расчет по прочности при нарушении сцепления анкера с основанием 7.1.2.2
            text = "Проверка по прочности при нарушении сцепления анкера с основанием 7.1.2.2 выполняется"
            print(text)
        else:
            text = "Проверка по прочности при нарушении сцепления анкера с основанием 7.1.2.2 выполняется НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Nan_tot <= self.N_ul_p:  # Расчет по прочности при нарушении сцепления анкера с основанием 7.1.5.3
            text = "Проверка по прочности при нарушении сцепления анкера с основанием 7.1.5.3 выполняется"
            print(text)
        else:
            text = "Проверка по прочности при нарушении сцепления анкера с основанием 7.1.5.3 НЕ выполняется!!"

        proverka_usl.append(text)
        if self.Nan <= self.N_ult_c:  # по выкалыванию основания  7.1.3
            text = "Проверка по выкалыванию основания  7.1.3выполняется"
            print(text)
        else:
            text = "Проверка по выкалыванию основания  7.1.3 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Nan_tot <= self.N_ult_c:  # по выкалыванию основания  7.1.3.2
            text = "Проверка по выкалыванию основания  7.1.3.2 выполняется"
            print(text)
        else:
            text = "Проверка по выкалыванию основания  7.1.3.2 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Nan <= self.N_ult_sp:  # Раскалывание основания 7.1.4.1
            text = "Проверка раскалывание основания 7.1.4.1 выполняется"
            print(text)
        else:
            text = "Проверка раскалывание основания 7.1.4.1 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Nan_tot <= self.N_ult_sp:  # Раскалывание основания 7.1.4.2
            text = "Проверка раскалывание основания  для группы анкеров 7.1.4.2 выполняется"
            print(text)
        else:
            text = "Проверка раскалывание основания  для группы анкеров 7.1.4.2 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Van <= self.V_ult_s:  # 7.2.1.1 Разрушение по стали анкера
            text = "Проверка разрушение по стали анкера по поперечной силе 7.1.4.2 выполняется"
            print(text)
        else:
            text = "Проверка разрушение по стали анкера по поперечной силе 7.1.4.2 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Van <= self.V_ult_cp:  # Разрущение от выкалывания основания
            text = "Проверка разрушение от выкалывания основания по поперечной силе 7.1.4.2 выполняется"
            print(text)
        else:
            text = "Проверка разрушение от выкалывания основания по поперечной силе  7.1.4.2 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Van_tot <= self.V_ult_cp:  # Разрушение от выкалывания основания группы 7.2.2.1
            text = "Проверка  разрушение от выкалывания основания группы 7.2.2.1 выполняется"
            print(text)
        else:
            text = "Проверка  разрушение от выкалывания основания группы 7.2.2.1 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Van <= self.V_ult_c:  # Разрушение от откалывания края основания 7.2.3.1
            text = "Проверка  разрушение от откалывания края основания 7.2.3.1 выполняется"
            print(text)
        else:
            text = "Проверка  разрушение от откалывания края основания 7.2.3.1 НЕ выполняется!!"
        proverka_usl.append(text)

        if self.Van_tot <= self.V_ult_c:  # Разрушение от откалывания края основания группы 7.2.3.2
            text = "Проверка  разрушение от откалывания края основания группы 7.2.3.2 выполняется"
            print(text)
        else:
            text = "Проверка  разрушение от откалывания края основания группы 7.2.3.2 НЕ выполняется!!"
        proverka_usl.append(text)


if __name__ == '__main__':
    book = xw.books
    # sheet = book.active.sheets
    wb = xw.Book(r'Ankers_table.xlsx')
    sheet = wb.sheets["Лист1"]
    data_ank2: pd.DataFrame
    data_ank: pd.DataFrame
    data_ank = sheet.range("A1").options(pd.DataFrame, expand='table', index_col=True).value
    data_ank = data_ank.reset_index()
    data_ank2 = sheet.range("G1").options(pd.DataFrame, expand='table', index_col=True).value
    print(data_ank)
    print(data_ank2)
    Nan = data_ank["N"].abs()
    Nan = Nan.max()
    ic(Nan)
    Nan_tot = data_ank["N"].sum()  # продольное усилие в группе анкеров
    Van = data_ank["V"].abs()  # 7.2.1.1 Усилие горизонтальное в одном анкере
    Van = Van.max()
    Van_tot = data_ank["V"].sum()  # Cуммарное усилие в группе

    ic(Nan, Nan_tot, Van, Van_tot)
    print(data_ank2.loc["Образование трещин", "Параметр"])
    # ank = Ankern_Krep(Nan=Nan,Nan_tot=Nan_tot,Van=Van,Van_tot=Van_tot,data=data_ank2, data1=data_ank)
    ank = Ankern_Krep(Nan, Nan_tot, Van, Van_tot, data_ank2, data_ank)
