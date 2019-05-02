from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5 import QtGui, QtCore
from PyQt5 import QtWidgets
import sys
import sqlite3
import datetime
from PyQt5.uic import loadUiType
from xlrd import *
from xlsxwriter import *
from hashutils import make_pw_hash, check_pw_hash

ui,_ = loadUiType('ui.ui')


class MainApp(QMainWindow , ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.statusBar().setStyleSheet("color: white")
        self.Handle_buttons()
        self.Handle_UI_Changes()
        self.show_activity_combobox()
        self.show_team_combobox()
        self.show_lob_combobox()
        self.show_task_combobox()
        self.show_type_combobox()
        self.show_user_combobox()

    def Handle_UI_Changes(self):
        self.tabWidget.tabBar().setVisible(False)

    def Handle_buttons(self):
        self.pushButton.clicked.connect(self.Open_daily)
        self.pushButton_2.clicked.connect(self.Open_tasks)
        self.pushButton_3.clicked.connect(self.Open_users)
        self.pushButton_6.clicked.connect(self.add_new_task)
        self.pushButton_8.clicked.connect(self.search_task)
        self.pushButton_7.clicked.connect(self.edit_tasks)
        self.pushButton_9.clicked.connect(self.delete_tasks)
        self.pushButton_10.clicked.connect(self.add_new_user)
        self.pushButton_11.clicked.connect(self.login)
        self.pushButton_5.clicked.connect(self.addRow)
        self.pushButton_5.clicked.connect(self.show_lcd)
        self.pushButton_4.clicked.connect(self.delete)
        self.pushButton_13.clicked.connect(self.insert_data)
        self.pushButton_15.clicked.connect(self.change_tab)
        self.pushButton_14.clicked.connect(self.register)
        self.pushButton_16.clicked.connect(self.profile)
        self.pushButton_17.clicked.connect(self.logoff)
        self.pushButton_18.clicked.connect(self.home)
        self.pushButton_12.clicked.connect(self.edit_user)
        self.pushButton_19.clicked.connect(self.retrieve)
        self.pushButton_20.clicked.connect(self.admin_extract)
        self.radioButton_2.toggled.connect(self.show_user_combobox)
        self.pushButton_22.clicked.connect(self.search_user)
        self.pushButton_21.clicked.connect(self.edit_type)

    def make_pw_hash(password):
        return hashlib.sha256(str.encode(password)).hexdigest()

    def check_pw_hash(password, hash):
        if make_pw_hash(password)== hash:
            return True
        return False




#############Functions#########################
    # Insert Data to SQLite

    def insert_data(self):
        Team = [self.tableWidget.item(row, 2).text() for row in range(self.tableWidget.rowCount())]
        LOB = [self.tableWidget.item(row, 1).text() for row in range(self.tableWidget.rowCount())]
        Task = [self.tableWidget.item(row, 4).text() for row in range(self.tableWidget.rowCount())]
        Activity = [self.tableWidget.item(row, 3).text() for row in range(self.tableWidget.rowCount())]
        Date = [self.tableWidget.item(row, 0).text() for row in range(self.tableWidget.rowCount())]
        Time = [self.tableWidget.item(row, 5).text() for row in range(self.tableWidget.rowCount())]
        comments = [self.tableWidget.item(row, 6).text() for row in range(self.tableWidget.rowCount())]
        user = [self.tableWidget.item(row, 7).text() for row in range(self.tableWidget.rowCount())]
        cols = list(zip(Team ,LOB ,Task ,Activity ,Time ,user ,Date ,comments))
        data = zip(Date, Time)
        dates = []
        for d in data:
            dt = datetime.datetime.strptime("{}, {}".format(*d), "%Y-%m-%d, %H:%M:%S")
            dates.append(dt)
        ########## Count unique dates ############
        u_dates = []
        for uniques in Date:
            if uniques not in u_dates:
                u_dates.append(uniques)
        ###########################################
        totals = {}
        for d in dates:
            if d.date() not in totals: totals[d.date()] = d.hour+d.minute/60.0
            else: totals[d.date()] += d.hour+d.minute/60.0
        val = ', '.join(map(str, cols))
        for date, time in totals.items():
            s = sum(totals.values())/len(u_dates)
            print(s)

        if s>=8.0:
            con = sqlite3.connect('dut.db')
            cur = con.cursor()
            cur.execute("""INSERT INTO day(team, LOB, task, activity, time, user, date_date, comments)
                        VALUES {}""".format(val))
            con.commit()
            self.statusBar().showMessage("Data Submitted Successfully")
            self.tableWidget.clearContents()
            self.tableWidget.setRowCount(0)
            self.textEdit_3.clear()
            self.lcdNumber.display(0)
        elif time<8.0 or s<8.0:
                warning = QMessageBox.warning(self, 'Login hours', "Please submit an average of 8 hours per day",
                                                    QMessageBox.Ok)
        else:
            self.statusBar().showMessage("Error in submission")


    def delete(self):
        self.tableWidget.removeRow(self.tableWidget.currentRow())
        self.show_lcd()

    def addRow(self):
        # Retrieve text from QLineEdit
        date = self.dateEdit.date()
        d = str(date.toString('yyyy-MM-dd'))
        task = self.comboBox.currentText()
        time = self.timeEdit.time()
        activity = self.comboBox_10.currentText()
        LOB = self.comboBox_11.currentText()
        team = self.comboBox_9.currentText()
        comments = self.textEdit_3.toPlainText()
        user = self.label_24.text()
        t = str(time.toString('HH:mm:ss'))
        # Create a empty row at bottom of table
        numRows = self.tableWidget.rowCount()
        self.tableWidget.insertRow(numRows)
        # Add text to the row
        self.tableWidget.setItem(numRows, 0, QtWidgets.QTableWidgetItem(d))
        self.tableWidget.setItem(numRows, 1, QtWidgets.QTableWidgetItem(LOB))
        self.tableWidget.setItem(numRows, 2, QtWidgets.QTableWidgetItem(team))
        self.tableWidget.setItem(numRows, 3, QtWidgets.QTableWidgetItem(activity))
        self.tableWidget.setItem(numRows, 4, QtWidgets.QTableWidgetItem(task))
        self.tableWidget.setItem(numRows, 5, QtWidgets.QTableWidgetItem(t))
        self.tableWidget.setItem(numRows, 6, QtWidgets.QTableWidgetItem(comments))
        self.tableWidget.setItem(numRows, 7, QtWidgets.QTableWidgetItem(user))


    def show_lcd(self):

        Time = [self.tableWidget.item(row, 5).text() for row in range(self.tableWidget.rowCount())]
        total=0
        for i in Time:
                h, m, s = map(int, i.split(":"))
                total += 3600*h + 60*m + s
                d="%02d:%02d:%02d" % (total / 3600, total / 60 % 60, total % 60)
        self.lcdNumber.display(str(d))


    # Extract Date to excel
    def admin_extract(self):
        dialog = QFileDialog()
        from_date = str(self.dateEdit_2.date().toString('yyyy-MM-dd'))
        to_date = str(self.dateEdit_3.date().toString('yyyy-MM-dd'))
        user = self.comboBox_12.currentText()
        dialog.setDefaultSuffix('xlsx')
        if self.radioButton.isChecked():
            self.comboBox_12.clear()
            self.db = sqlite3.connect('dut.db')
            self.cur = self.db.cursor()
            self.cur.execute('''  SELECT * FROM day WHERE date_date >= ? AND date_date <= ? ''', (from_date, to_date))
            data = self.cur.fetchall()
        else:
            self.cur.execute('''  SELECT * FROM day WHERE date_date >= ? AND date_date <= ? AND user =? ''', (from_date, to_date, user, ))
            data = self.cur.fetchall()
        fileName,_ = QFileDialog.getSaveFileName(self, "Extract Data", "", "Excel Files (*.xlsx)")
        wb = Workbook(fileName)
        sheet1 = wb.add_worksheet()
        sheet1.write(0, 0, 'ID')
        sheet1.write(0, 1, 'team')
        sheet1.write(0, 2, 'LOB')
        sheet1.write(0, 3, 'task')
        sheet1.write(0, 4, 'activity')
        sheet1.write(0, 5, 'time')
        sheet1.write(0, 6, 'description')
        sheet1.write(0, 7, 'responsible')
        sheet1.write(0, 8, 'user')
        sheet1.write(0, 9, 'date')
        sheet1.write(0, 10, 'comments')

        row_no = 1
        for row in data:
            col_no = 0
            for item in row:
                sheet1.write(row_no, col_no , str(item))
                col_no +=1
            row_no +=1

        wb.close()
        self.statusBar().showMessage('Report extracted')

    def show_user_combobox(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT user_name FROM users ''')
        data = self.cur.fetchall()
        self.comboBox_12.clear()
        for user in data:
            self.comboBox_12.addItem(user[0])

    def show_type_combobox(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT usertype FROM usertype ''')
        data = self.cur.fetchall()
        self.comboBox_2.clear()
        for user in data:
            self.comboBox_2.addItem(user[0])


    def edit_type(self):
        user = self.lineEdit_20.text()
        type1 = self.comboBox_2.currentText()
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''UPDATE users SET type=? WHERE user_name =? ''', (type1, user, ))
        self.db.commit()
        self.statusBar().showMessage('User updated')

    def search_user(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        user = self.lineEdit.text()
        sql = '''SELECT user_name, type FROM users WHERE user_name =?'''
        ex = self.cur.execute(sql, [(user)])
        data = ex.fetchone()
        if len(data) > 0:
            self.lineEdit_20.setText(data[0])
            self.lineEdit_21.setText(data[1])
        else:
            self.statusBar().showMessage('No user found')


    def daily_add(self): #37
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.db.commit()
        date_date = self.dateEdit.date()
        task = self.comboBox.currentText()
        time = self.timeEdit.time()
        activity = self.comboBox_10.currentText()
        LOB = self.comboBox_11.currentText()
        team = self.comboBox_9.currentText()
        user = self.label_24.text()

        self.cur.execute('''
            INSERT INTO day(date_date, team, LOB, task, activity, time, user)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''' ,(str(date_date.toString('dd-MM-yyyy')), team, LOB, task, activity, str(time.toString('HH:mm')),user))
        self.db.commit()
        self.statusBar().showMessage('Daily task added')
        self.show_day()
        self.show_hours()

    def show_day(self):
        today_date = self.dateEdit.date()
        datestr = str(today_date.toString('dd-MM-yyyy'))
        user = self.label_24.text()
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute(''' SELECT id, date_date, LOB, team, activity, task, time,
          description, responsible FROM day WHERE date_date=? AND user=?''', (str(datestr), user))
        data = self.cur.fetchall()
        if data:
            self.tableWidget.setRowCount(0)
            self.tableWidget.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)

    def retrieve(self):
        from_date = str(self.dateEdit_2.date().toString('yyyy-MM-dd'))
        to_date = str(self.dateEdit_3.date().toString('yyyy-MM-dd'))
        user = self.label_24.text()
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        qry = ('''SELECT * FROM day WHERE date_date BETWEEN "{}" AND "{}" AND user="{}"''').format(str(from_date), str(to_date), user)
        print(qry)
        self.cur.execute(qry)
        data = self.cur.fetchall()
        print(data)
        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
        else:
            self.statusBar().showMessage('No Data')

    def show_hours(self): #38
        today_date = self.dateEdit.date()
        datestr = str(today_date.toString('dd-MM-yyyy'))
        user=self.label_24.text()
        self.db = sqlite3.connect('dut.db')
        self.db.commit()
        self.cur = self.db.cursor()
        sql = self.cur.execute('''SELECT sum(time) FROM day
        WHERE user =? AND date_date=?''', (user, str(datestr)))
        self.db.commit()
        exe = sql.fetchone()
        self.lcdNumber.display(str(exe))



    def del_day(self):
        try:
            row = (self.tableWidget.item(self.tableWidget.currentIndex().row(), 0).text())
            self.db = sqlite3.connect('dut.db')
            self.db.commit()
            self.cur = self.db.cursor()
            self.cur.execute('''DELETE FROM day WHERE id=?''', (row,))
            self.db.commit()
            self.statusBar().showMessage('Task deleted')
            self.show_hours()
            self.show_day()
        except:
            self.statusBar().showMessage('Please choose a transaction to delete')

    def submit_task(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        warning = QMessageBox.warning(self, 'Update', "Are you sure to submit the details?",
                                       QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes:
            self.db.commit()
            self.statusBar().showMessage('Task details submitted successfuly')




    def Open_daily(self):
        if self.label_24.text() != '':
            self.tabWidget.setCurrentIndex(0)
            self.statusBar().showMessage('Update task details')
        else:
            self.statusBar().showMessage('Login to update task details')
            self.tabWidget.setCurrentIndex(2)

    def Open_tasks(self):
        if self.label_24.text() != '':
            self.tabWidget.setCurrentIndex(1)
            self.statusBar().showMessage('Manage tasks')
        else:
            self.statusBar().showMessage('Login to manage tasks')
            self.tabWidget.setCurrentIndex(2)

    def Open_users(self):
            if self.label_24.text()!='':
                self.tabWidget.setCurrentIndex(3)
                self.statusBar().showMessage('User home')
            else:
                self.tabWidget.setCurrentIndex(2)
                self.statusBar().showMessage('Please Login or Register!')




#######Tasks###############

    def add_new_task(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()

        task = self.lineEdit_2.text()
        description = self.textEdit.toPlainText()
        responsible = self.lineEdit_3.text()
        activity = self.comboBox_5.currentText()
        LOB = self.comboBox_3.currentText()
        team = self.comboBox_4.currentText()

        self.cur.execute('''
            INSERT INTO task(task,description,responsible,activity,LOB,team)
            VALUES (?, ?, ?, ?, ?, ?)
        ''' ,(task , description , responsible , activity , LOB, team))
        self.db.commit()
        self.statusBar().showMessage('New Task added')
        self.show_lob_combobox()
        self.show_team_combobox()
        self.show_activity_combobox()
        self.show_task_combobox()


    def search_task(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        task = self.lineEdit_5.text()
        sql = '''SELECT *, COUNT(task) FROM task WHERE task =?'''
        ex = self.cur.execute(sql, [(task)])
        data = ex.fetchone()
        if len(data) > 0:
            self.lineEdit_6.setText(data[0])
            self.textEdit_2.setPlainText(data[1])
            self.lineEdit_4.setText(data[2])
            self.comboBox_8.setCurrentIndex(data[5])
            self.comboBox_7.setCurrentIndex(data[4])
            self.comboBox_6.setCurrentIndex(data[3])
        else:
            self.statusBar().showMessage('No Task found')


    def edit_tasks(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        task = self.lineEdit_5.text()
        description = self.textEdit_2.toPlainText()
        responsible = self.lineEdit_4.text()
        activity = self.comboBox_6.currentIndex()
        LOB = self.comboBox_7.currentIndex()
        team = self.comboBox_8.currentIndex()
        search_task = self.lineEdit_5.text()
        self.cur.execute('''
            UPDATE task SET task=?, description=?, responsible=?, activity=?, LOB=?, team=? WHERE task=?
        ''' , (task,description,responsible,activity,LOB,team,search_task))
        self.db.commit()
        self.statusBar().showMessage('Task updated')


    def delete_tasks(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        task = self.lineEdit_5.text()
        warning = QMessageBox.warning(self, 'Delete Task', "Are you sure to delete the task?" , QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            sql = '''DELETE FROM task WHERE task=?'''
            self.cur.execute(sql, [(task)])
            self.db.commit()
            self.statusBar().showMessage('Task Removed')


#######User Registration###############

    def add_new_user(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        first = self.lineEdit_18.text()
        last = self.lineEdit_17.text()
        user = self.lineEdit_7.text()
        email = self.lineEdit_9.text()
        password = self.lineEdit_8.text()
        password2 = self.lineEdit_10.text()
        if password == password2:
            self.cur.execute('''INSERT INTO users(user_name, user_email, user_password,first,last)
                                VALUES (? , ? , ? , ? ,? )''', (user, email, make_pw_hash(password), first, last))
            self.db.commit()
            self.statusBar().showMessage('Registration Successful!')
            msg = QMessageBox.warning(self, 'Registration Status',
                                      "Registration Successfull! Use your credentials to login",
                                      QMessageBox.Ok)
            if msg == QMessageBox.Ok:
                self.tabWidget.setCurrentIndex(2)
                self.lineEdit_18.clear()
                self.lineEdit_17.clear()
                self.lineEdit_7.clear()
                self.lineEdit_9.clear()
                self.lineEdit_8.clear()
                self.lineEdit_10.clear()
        else:
            self.label_27.setText('Password do not match')


    def login(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        username = self.lineEdit_11.text()
        password = self.lineEdit_12.text()
        sql = '''SELECT * FROM users WHERE user_name =? '''
        self.cur.execute(sql,([username]))
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and check_pw_hash(password, row[3]):
                if row[6] == 'General':
                    self.statusBar().showMessage('Valid Username & Password')
                    self.tab_7.setEnabled(True)
                    self.tab.setEnabled(True)# enable day screen
                    self.tab_10.setEnabled(True)
                    self.lineEdit_13.setText(row[4])
                    self.lineEdit_14.setText(row[2])
                    self.lineEdit_16.setText(row[3])
                    self.lineEdit_19.setText(row[5])
                    self.label_24.setText(username)
                    self.tabWidget.setCurrentIndex(3)
                    self.tabWidget_2.setVisible(False)
                    self.pushButton_2.clicked.connect(self.restrict_tab)
                elif row[6] == 'Superuser':
                    self.statusBar().showMessage('Valid Username & Password')
                    self.tab_7.setEnabled(True)
                    self.tab.setEnabled(True)# enable day screen
                    self.tab_3.setEnabled(True)
                    self.lineEdit_13.setText(row[4])
                    self.lineEdit_14.setText(row[2])
                    self.lineEdit_16.setText(row[3])
                    self.lineEdit_19.setText(row[5])
                    self.label_24.setText(username)
                    self.tabWidget.setCurrentIndex(3)
                    self.tabWidget_2.setVisible(True)
                    self.pushButton_2.clicked.connect(self.Open_tasks)
            else:
                self.statusBar().showMessage('Incorrect Username or Password')

    def restrict_tab(self):
        self.tabWidget.setCurrentIndex(6)

    def change_tab(self):
        self.tabWidget.setCurrentIndex(0)

    def register(self):
        self.tabWidget.setCurrentIndex(4)


    def profile(self):
        self.tabWidget.setCurrentIndex(5)

    def home(self):
        self.tabWidget.setCurrentIndex(3)


    def edit_user(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()

        first = self.lineEdit_13.text()
        last = self.lineEdit_19.text()
        email = self.lineEdit_14.text()
        password = self.lineEdit_16.text()
        password2 = self.lineEdit_15.text()

        if password == password2:
            self.cur.execute('''
            UPDATE users SET first=?, last=?, user_email=?, user_password=?''', (first, last, email, password, ))
            self.db.commit()
            self.statusBar().showMessage('Profile Updated')

        else:
            self.label_27.setText('Password do not match')

    def logoff(self):
        self.label_24.clear()
        self.tab.setEnabled(False)
        self.tabWidget.setCurrentIndex(2)
        self.lineEdit_11.clear()
        self.lineEdit_12.clear()
        self.tab_3.setEnabled(False)
        self.tab_7.setEnabled(False)
        self.statusBar().showMessage('User logged off')






#######settings###############
    def add_activity(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        category_name = self.lineEdit_17.text()
        self.cur.execute('''
            INSERT INTO category (category_name) VALUES (%s)
        ''' , (category_name,))
        self.db.commit()
        self.statusBar().showMessage('New Category Added')
        self.lineEdit_17.setText('')
        self.show_category()
        self.show_activity_combobox()


    def add_team(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        team_name = self.lineEdit_18.text()
        self.cur.execute('''
            INSERT INTO team (team_name) VALUES (%s)
        ''', (team_name,))
        self.db.commit()
        self.lineEdit_18.setText('')
        self.statusBar().showMessage('New Team Added')
        self.show_team()
        self.show_team_combobox()

    def add_lob(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        lob_name = self.lineEdit_19.text()
        self.cur.execute('''
                INSERT INTO lob (lob_name) VALUES (%s)
            ''', (lob_name,))
        self.db.commit()
        self.lineEdit_19.setText('')
        self.statusBar().showMessage('New LOB Added')
        self.show_lob()
        self.show_lob_combobox()




    def show_activity_combobox(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT activity_name FROM category ''')
        data = self.cur.fetchall()
        self.comboBox_5.clear()
        self.comboBox_6.clear()
        self.comboBox_10.clear()
        for category in data:
            self.comboBox_5.addItem(category[0])
            self.comboBox_6.addItem(category[0])
            self.comboBox_10.addItem(category[0])


    def show_team_combobox(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT team_name FROM team ''')
        data = self.cur.fetchall()
        self.comboBox_8.clear()
        self.comboBox_4.clear()
        self.comboBox_9.clear()
        for team in data:
            self.comboBox_8.addItem(team[0])
            self.comboBox_4.addItem(team[0])
            self.comboBox_9.addItem(team[0])
    def show_lob_combobox(self):
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT lob_name FROM lob ''')
        data = self.cur.fetchall()
        self.comboBox_7.clear()
        self.comboBox_3.clear()
        self.comboBox_3.clear()
        for lob in data:
            self.comboBox_7.addItem(lob[0])
            self.comboBox_3.addItem(lob[0])
            self.comboBox_11.addItem(lob[0])

    def show_task_combobox(self):
        activity = self.comboBox_10.currentText()
        team = self.comboBox_9.currentText()
        self.db = sqlite3.connect('dut.db')
        self.cur = self.db.cursor()
        self.cur.execute('''  SELECT task FROM task WHERE team=? and activity=?''', (team, activity, ))
        data = self.cur.fetchall()
        self.comboBox.clear()
        for task in data:
            self.comboBox.addItem(task[0])
        self.comboBox_9.currentTextChanged.connect(self.show_task_combobox)
        self.comboBox_10.currentTextChanged.connect(self.show_task_combobox)


    def export_data(self):
        try:
            dialog = QFileDialog()
            dialog.setDefaultSuffix('xlsx')
            self.db = sqlite3.connect('dut.db')
            self.cur = self.db.cursor()
            self.cur.execute('''  SELECT * FROM day''')
            data = self.cur.fetchall()
            fileName,_ = QFileDialog.getSaveFileName(self, "Extract Data", "", "Excel Files (*.xlsx)")
            wb = Workbook(fileName)
            sheet1 = wb.add_worksheet()
            sheet1.write(0, 0, 'ID')
            sheet1.write(0, 1, 'team')
            sheet1.write(0, 2, 'LOB')
            sheet1.write(0, 3, 'task')
            sheet1.write(0, 4, 'activity')
            sheet1.write(0, 5, 'time')
            sheet1.write(0, 6, 'description')
            sheet1.write(0, 7, 'responsible')
            sheet1.write(0, 8, 'user')
            sheet1.write(0, 9, 'date')

            row_no = 1
            for row in data:
                col_no = 0
                for item in row:
                    sheet1.write(row_no, col_no , str(item))
                    col_no +=1
                row_no +=1

            wb.close()
            self.statusBar().showMessage('Report extracted')
        except:
            self.statusBar().showMessage('Extract cancelled')




def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))
    window = MainApp()
    window.show()
    app.exec_()


if __name__ =='__main__':
    main()
