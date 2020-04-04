import sys
from tkinter import *
from tkinter import messagebox
from docxtpl import DocxTemplate
from test import *
from PyQt5 import QtCore, QtGui, QtWidgets
from test import Ui_MainWindow
from tkinter import ttk


def okno2():
    if v.get() == '':
        messagebox.showerror("Ошибка", "Не выбран пользователь")
    else:
        tk.destroy()

        class MyWin(QtWidgets.QMainWindow):

            def __init__(self, parent=None):
                QtWidgets.QWidget.__init__(self, parent)
                self.ui = Ui_MainWindow()
                self.ui.setupUi(self)
                self.setWindowTitle('Мастер ДАП ver. 1.0.04.20 / ' + v.get())
                self.ui.pushButton_go.clicked.connect(self._generate)
                self.ui.pushButton_exit.clicked.connect(self.close_window)

            def close_window():
                self.destroy()
        #кпонка сформировать

            def _generate(self):
                if v.get() == "М.В. Дубровина":
                    doloar = "Уполномоченный"
                    doloar_vp = "уполномоченного"
                    fiooar_vp = "М.В. Дубровину"
                if v.get() == "Н.В. Задорожная":
                    doloar = "Старший уполномоченный"
                    doloar_vp = "старшего уполномоченного"
                    fiooar_vp = "Н.В. Задорожную"
                if v.get() == "А.П. Зайцев":
                    doloar = "Уполномоченный по особо важным делам"
                    doloar_vp = "уполномоченного по особо важным делам"
                    fiooar_vp = "А.П. Зайцева"
                if v.get() == "А.Г. Левина":
                    doloar = "Старший уполномоченный"
                    doloar_vp = "старшего уполномоченного"
                    fiooar_vp = "А.Г. Левину"
                if v.get() == "О.А. Петрушева":
                    doloar = "Старший уполномоченный по особо важным делам"
                    doloar_vp = "старшего уполномоченного по особо важным делам"
                    fiooar_vp = "О.А. Петрушеву"
                if v.get() == "Е.Н. Сажина":
                    doloar = "Уполномоченный"
                    doloar_vp = "уполномоченного"
                    fiooar_vp = "Е.Н. Сажину"
                if v.get() == "Э.В. Степанов":
                    doloar = "Уполномоченный по особо важным делам"
                    doloar_vp = "уполномоченного по особо важным делам"
                    fiooar_vp = "Э.В. Степанова"
                if v.get() == "Е.В. Трунина":
                    doloar = "Старший уполномоченный"
                    doloar_vp = "старшего уполномоченного"
                    fiooar_vp = "Е.В. Трунину"
                if v.get() == "А.С. Хлыбов":
                    doloar = "Заместитель начальника"
                    doloar_vp = "заместителя начальника"
                    fiooar_vp = "А.С. Хлыбова"
                if v.get() == "Е.В. Шумилина":
                    doloar = "Уполномоченный"
                    doloar_vp = "уполномоченного"
                    fiooar_vp = "Е.В. Шумилину"
                if v.get() == "А.В. Яльтиков":
                    doloar = "Старший уполномоченный по особо важным делам"
                    doloar_vp = "старшего уполномоченного по особо важным делам"
                    fiooar_vp = "А.В. Ялтикова"
                number = self.ui.lineEdit_number.text()
                chkoap = self.ui.lineEdit_chkoap.text()
                stkoap = self.ui.lineEdit_stkoap.text()
                dd1 = self.ui.lineEdit_dd1.text()
                dd2 = self.ui.lineEdit_dd2.text()
                dd3 = self.ui.lineEdit_dd3.text()
                dd4 = self.ui.lineEdit_dd4.text()
                dd5 = self.ui.lineEdit_dd5.text()
                month1 = self.ui.comboBox_month1.currentText()
                month2 = self.ui.comboBox_month2.currentText()
                month3 = self.ui.comboBox_month3.currentText()
                month4 = self.ui.comboBox_month4.currentText()
                month5 = self.ui.comboBox_month5.currentText()
                ul = self.ui.lineEdit_ul.text()
                zp_ip = self.ui.lineEdit_zp_ip.text()
                zp_rp = self.ui.lineEdit_zp_rp.text()
                zp_dp = self.ui.lineEdit_zp_dp.text()
                ogrn = self.ui.lineEdit_ogrn.text()
                inn = self.ui.lineEdit_inn.text()
                kpp = self.ui.lineEdit_kpp.text()
                ad_index = self.ui.lineEdit_ad_index.text()
                subrf = self.ui.lineEdit_subrf.text()
                naspunkt = self.ui.lineEdit_naspunkt.text()
                ulitsadom = self.ui.lineEdit_ulitsadom.text()
                dateul = self.ui.lineEdit_dateul.text()
                fio1 = self.ui.comboBox_perar.currentText()
                fio2 = self.ui.comboBox_zapros.currentText()
                fio3 = self.ui.comboBox_prinyal.currentText()
                fio4 = self.ui.comboBox_otlozh.currentText()
                fio5 = self.ui.comboBox_rassm.currentText()
                if fio1 == "В.М. Мамонтов":
                    znt1 = "Заместитель"
                    znt1_1 = "заместителя"
                    fio1_1 = "В.М. Мамонтова"
                else:
                    znt1 = "И.о. заместителя"
                    znt1_1 = "и.о. заместителя"
                    fio1_1 = "Н.П. Алексеевой"
                if fio2 == "В.М. Мамонтов":
                    znt2 = "Заместитель"
                else:
                    znt2 = "И.о. заместителя"
                if fio3 == "В.М. Мамонтов":
                    znt3 = "Заместитель"
                    znt3_1 = "Заместителю"
                    fio3_1 = "В.М. Мамонтову"
                else:
                    znt3 = "И.о. заместителя"
                    znt3_1 = znt3
                    fio3_1 = "Н.П. Алексеевой"
                if fio4 == "В.М. Мамонтов":
                    znt4 = "Заместитель"
                else:
                    znt4 = "И.о. заместителя"
                if fio5 == "В.М. Мамонтов":
                    znt5 = "Заместитель"
                else:
                    znt5 = "И.о. заместителя"

        #условие если выбран 1 файл ГОТОВ
                if self.ui.checkBox_01.isChecked():
                    doc1 = DocxTemplate("templates/1.docx")
                    context1 = {
                                'dd2': dd2, 'doloar': doloar, 'doloar_vp': doloar_vp,
                                'fiooar': v.get(), 'fiooar_vp': fiooar_vp,
                                'month2': month2, 'number': number,
                                'ul': ul, 'orgn': ogrn, 'inn': inn, 'kpp': kpp,
                                'ad_index': ad_index, 'subrf': subrf,
                                'naspunkt': naspunkt, 'ulitsadom': ulitsadom,
                                'dateul': dateul, 'chkoap': chkoap, 'stkoap': stkoap,
                                'znt1': znt1, 'fio1': fio1
                               }
                    doc1.render(context1)
                    doc1.save(number + "__ 2.1 РЕШЕНИЕ о передаче дела для проведения АР.docx")

        #условие если выбран 2 файл ГОТОВ
                if self.ui.checkBox_02.isChecked():
                    doc2 = DocxTemplate("templates/2.docx")
                    context2 = {
                                'dd2': dd2, 'month2': month2, 'number': number,
                                'ul': ul, 'orgn': ogrn, 'inn': inn, 'kpp': kpp,
                                'ad_index': ad_index, 'subrf': subrf,
                                'naspunkt': naspunkt, 'ulitsadom': ulitsadom,
                                'dateul': dateul, 'znt1_1': znt1_1, 'fio1_1': fio1_1
                               }
                    doc2.render(context2)
                    doc2.save(number + "__ 2.2 Определение о принятии дела к своему производству.docx")

        #условие если выбран 8 файл
                if self.ui.checkBox_08.isChecked():
                    doc8 = DocxTemplate("templates/8.docx")
                    context8 = {'number': number}
                    doc8.render(context8)
                    doc8.save(number + "__ 2.8 запрос в ГИБДД.docx")

        #условие если выбран 9 файл
                if self.ui.checkBox_09.isChecked():
                    doc9 = DocxTemplate("templates/9.docx")
                    context9 = {'number': number}
                    doc9.render(context9)
                    doc9.save(number + "__ 2.9 запрос в Росреестр.docx")

        #условие если выбран 10 файл
                if self.ui.checkBox_10.isChecked():
                    doc10 = DocxTemplate("templates/10.docx")
                    context10 = {
                                'number': number
                                }
                    doc10.render(context10)
                    doc10.save(number + "__ 3.1 рапорт.docx")

        #условие если выбран 11 файл
                if self.ui.checkBox_11.isChecked():
                    doc11 = DocxTemplate("templates/11.docx")
                    context11 = {
                                'number': number, 'dd4': dd4,
                                'znt3': znt3, 'fio3': fio3
                                }
                    doc11.render(context11)
                    doc11.save(number + "__ 3.2 справка об издержках.docx")

                messagebox.showinfo('Сформировано', 'Успешно сформировано по делу об АП \n №10418000-' + number+'/2020')
                readfile()
        if __name__ == "__main__":
            app = QtWidgets.QApplication(sys.argv)
            myapp = MyWin()
            myapp.show()
            sys.exit(app.exec_())

tk = Tk()
tk.title("Мастер ДАП ver. 1.0.04.20")
w = tk.winfo_screenwidth()
h = tk.winfo_screenheight()
w = w // 2
h = h // 2
w = w - 200
h = h - 100
tk.geometry('400x200+{}+{}'.format(w, h))
label = Label(text='Войти как:')
label.pack()
btn = Button(tk, text="Далее", height=2, width=20, command=okno2)
btn.pack(side=BOTTOM)
v = StringVar()
combobox = ttk.Combobox(tk, textvariable=v, values=[
                                    "М.В. Дубровина", "Н.В. Задорожная",
                                    "А.П. Зайцев", "А.Г. Левина",
                                    "О.А. Петрушева", "Е.Н. Сажина",
                                    "Э.В. Степанов", "Е.В. Трунина",
                                    "А.С. Хлыбов", "Е.В. Шумилина",
                                    "А.В. Яльтиков"], height = 11)
combobox['state'] = 'readonly'
combobox.pack()

tk.mainloop()