import tkinter as tk
import tkinter.font as tkFont
from datetime import date as Date
import threading
from operator import itemgetter
import xlsxwriter
import os
import sys


def makeDate(strDate):
    month, day, year = strDate.split('/')
    return Date(int(year), int(month), int(day))


def makeStr(dateObj):
    year, month, day = dateObj.year, dateObj.month, dateObj.day
    return str(month) + '/' + str(day) + '/' + str(year)


class MainMenu(tk.Frame):

    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.grid_propagate(0)
        self.payWorkerButton = tk.Button(master=self, text='Pay Workers', command=parent.drawPayWorkerMenu)
        self.manageWorkersButton = tk.Button(master=self, text='Manage Workers', command=parent.drawManageWorkersMenu)
        self.manageDistsButton = tk.Button(master=self, text='Manage Distributions and Companies',
                                           command=parent.drawManageCompaniesMenu)
        self.masterXLSheetButton = tk.Button(master=self, text='Create Master Spreadsheet',
                                             command=parent.drawMasterSpreadsheetMenu)
        self.exitButton = tk.Button(master=self, text='Exit', command=parent.killProgram)
        self.lbl1 = tk.Label(master=self, text='')
        self.lbl2 = tk.Label(master=self, text='')

        rowMinSize = (kwargs['height'] - 60) / 2.35
        columnSize = (kwargs['width'] - 80)
        self.grid_rowconfigure(0, minsize=rowMinSize)
        self.grid_rowconfigure(5, minsize=rowMinSize)
        self.grid_columnconfigure(1, minsize=columnSize)
        self.lbl1.grid(row=0)
        self.lbl2.grid(row=4)

        self.payWorkerButton.grid(row=1, column=1)
        self.manageWorkersButton.grid(row=2, column=1)
        self.manageDistsButton.grid(row=3, column=1)
        self.masterXLSheetButton.grid(row=4, column=1)
        self.exitButton.grid(row=6, column=0)

    def draw(self):
        self.grid()

    def undraw(self):
        self.grid_forget()


class MasterSpreadsheetMenu(tk.Frame):

    def __init__(self, parent, workerList, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.workerList = workerList
        self.companyList = companyList

        self.lblMain = tk.Label(master=self, text='Leave dates blank to include all data')
        self.lblStartDate = tk.Label(master=self, text='Start Date:')
        self.startDateText = tk.StringVar()
        self.entStartDate = tk.Entry(master=self, width=15, textvariable=self.startDateText)
        self.lblEndDate = tk.Label(master=self, text='End Date:')
        self.endDateText = tk.StringVar()
        self.entEndDate = tk.Entry(master=self, width=15, textvariable=self.endDateText)
        self.btnMakeSpreadsheet = tk.Button(master=self, text='Create Spreadsheet', command=self.makeSpreadsheet)
        self.btnBack = tk.Button(master=self, text='Back', command=parent.drawMainMenu)

        self.grid_rowconfigure(0, minsize=150)
        self.grid_columnconfigure(0, minsize=200)

        self.lblMain.grid(row=1, column=1, columnspan=5)
        self.lblStartDate.grid(row=2, column=1)
        self.entStartDate.grid(row=2, column=2)
        self.lblEndDate.grid(row=2, column=4)
        self.entEndDate.grid(row=2, column=5)
        self.btnMakeSpreadsheet.grid(row=3, column=3)
        self.btnBack.grid(row=5, column=1)

    def draw(self):
        self.grid()

    def undraw(self):
        self.grid_forget()

    def makeSpreadsheet(self):
        startDate = self.startDateText.get()
        endDate = self.endDateText.get()
        if startDate != '':
            startDate = makeDate(startDate)
        else:
            startDate = None
        if endDate != '':
            endDate = makeDate(endDate)
        else:
            endDate = None

        filename = ''
        hoursData = self.workerList.getAllHoursData()
        # data: [id, date, workerName, [company, hours], paidBy]
        exportHoursData = []
        if startDate:
            if endDate:
                filename = f'MasterSpreadsheet_{startDate}_{endDate}.xlsx'
                for row in hoursData:
                    if startDate <= row[1] <= endDate:
                        exportHoursData.append(row)
            else:
                filename = f'MasterSpreadsheet_{startDate}_NOW.xlsx'
                for row in hoursData:
                    if startDate <= row[1]:
                        exportHoursData.append(row)
        else:
            if endDate:
                filename = f'MasterSpreadsheet_BEGINNING_{endDate}.xlsx'
                for row in hoursData:
                    if row[1] <= endDate:
                        exportHoursData.append(row)
            else:
                filename = f'MasterSpreadsheet_ALL_DATA.xlsx'
                exportHoursData = hoursData

        desktopPath = ''
        if sys.platform == 'darwin' or sys.platform == 'linux':
            desktopPath = os.path.expanduser('~/Desktop')
        elif sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                desktopPath = os.path.join(os.environ['HOMEPATH'], 'Desktop')
            except:
                desktopPath = os.path.join(os.environ['USERPROFILE'], 'Desktop')

        # create the spreadsheet
        workbook = xlsxwriter.Workbook(os.path.join(desktopPath, filename))
        worksheet = workbook.add_worksheet()
        exportHoursData.sort(key=itemgetter(1), reverse=True)

        # set up the top row
        worksheet.write(0, 0, 'Date')
        worksheet.write(0, 1, 'Worker')
        worksheet.write(0, 2, 'Pay Rate')
        worksheet.write(0, 3, 'Data Type')
        worksheet.write(0, 4, 'Total')
        worksheet.write(0, 5, 'Miscalculation')
        companyColumn = 6
        companyStartColumn = companyColumn
        companies = self.companyList.getAllCompanies()
        companyColumns = {}
        for company in companies:
            worksheet.write(0, companyColumn, company.name)
            companyColumns[company.name] = companyColumn
            companyColumn += 1

        numToLetter = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E', 5: 'F', 6: 'G', 7: 'H', 8: 'I', 9: 'J', 10: 'K', 11: 'L',
                       12: 'M', 13: 'N', 14: 'O', 15: 'P', 16: 'Q', 17: 'R', 18: 'S', 19: 'T', 20: 'U', 21: 'V',
                       22: 'W', 23: 'X', 24: 'Y', 25: 'Z'}

        worksheet.write(1, 5, 'Company Totals:')
        companyTotals = {}
        for company in companies:
            companyTotals[company.name] = 0

        rowNum = 2
        for row in exportHoursData:
            worker = self.workerList[row[2]]
            worksheet.write(rowNum, 0, makeStr(row[1]))
            worksheet.write(rowNum, 1, row[2])
            worksheet.write(rowNum, 2, worker.getPayRate(row[1]))
            worksheet.write(rowNum, 3, 'Hours:')
            worksheet.write(rowNum+1, 3, 'Dollars:')
            worksheet.write(rowNum, 4,
                            f'=SUM({numToLetter[companyStartColumn]}{rowNum+1}:' +
                            f'{numToLetter[companyColumn-1]}{rowNum+1})')
            worksheet.write(rowNum+1, 4,
                            f'=SUM({numToLetter[companyStartColumn]}{rowNum+2}:' +
                            f'{numToLetter[companyColumn-1]}{rowNum+2})')
            miscalculation = worker.getMiscalculation(row[1])
            if miscalculation:
                worksheet.write(rowNum+1, 5, miscalculation)
            for tup in row[3]:
                column = companyColumns[tup[0]]
                worksheet.write(rowNum, column, tup[1])
                worksheet.write(rowNum + 1, column, f'=C{rowNum+1}*{numToLetter[column]}{rowNum+1}')
                companyTotals[tup[0]] += tup[1] * worker.getPayRate(row[1])
            rowNum += 2
        for company in [company.name for company in companies]:
            worksheet.write(1, companyColumns[company], companyTotals[company])
        worksheet.write(0, companyColumn + 1, 'Miscalculation is negative for overpayment, positive for underpayment')
        worksheet.write(1, companyColumn + 1, 'These totals do not reflect rent deduction')

        workbook.close()


class PayWorkerMenu(tk.Frame):

    def __init__(self, parent, workerList, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.grid_propagate(0)
        self.parent = parent
        self.workerList = workerList
        self.companyList = companyList
        self.config(padx=50, pady=50)

        self.frame1 = tk.Frame(master=self)
        self.lblWorker = tk.Label(master=self.frame1, text='Worker:')
        self.lblSelectedWorker = tk.Label(master=self.frame1, text='')
        self.sbxWorker = tk.Listbox(master=self.frame1, height=10, width=30)
        self.workerScroller = tk.Scrollbar(master=self.frame1)
        self.sbxWorker.configure(yscrollcommand=self.workerScroller.set)
        self.workerScroller.configure(command=self.sbxWorker.yview)
        self.sbxWorker.bind('<<ListboxSelect>>', self.selectWorker)

        self.lblDate = tk.Label(master=self.frame1, text='Date:')
        self.dateText = tk.StringVar()
        self.entDate = tk.Entry(master=self.frame1, width=8, textvariable=self.dateText)

        self.companyFrame = tk.Frame(master=self)
        # tuples of (index, label, entry, textVar)
        self.companyWidgets = []
        index = 0
        for company in companyList.getNames():
            lbl = tk.Label(master=self.companyFrame, text=str(company))
            textVar = tk.StringVar()
            ent = tk.Entry(master=self.companyFrame, width=5, textvariable=textVar)
            self.companyWidgets.append((index, lbl, ent, textVar))
            index += 1
        self.calcBtn = tk.Button(master=self.companyFrame, text='Calculate: $', command=self.calcTotal)
        self.lblTotal = tk.Label(master=self.companyFrame, text='0')
        self.btnClear = tk.Button(master=self.companyFrame, text='Clear', command=self.clearCompanies)
        self.lblPaidBy = tk.Label(master=self.companyFrame, text='Paid By:')
        self.paidByText = tk.StringVar()
        self.entPaidBy = tk.Entry(master=self.companyFrame, textvariable=self.paidByText)

        self.historyFrame = tk.Frame(master=self)
        self.sbxHistory = tk.Listbox(master=self.historyFrame, height=15, width=40)
        self.historyScroller = tk.Scrollbar(master=self.historyFrame)
        self.sbxHistory.configure(yscrollcommand=self.historyScroller.set)
        self.historyScroller.configure(command=self.sbxHistory.yview)
        self.backBtn = tk.Button(master=self.historyFrame, text='Back', command=parent.drawMainMenu)

        self.buttonFrame = tk.Frame(master=self)
        self.btnAdd = tk.Button(master=self.buttonFrame, text='Add Payment', command=self.addPayment)
        self.btnEdit = tk.Button(master=self.buttonFrame, text='Edit Selected', command=self.editSelected)
        self.btnSearch = tk.Button(master=self.buttonFrame, text='Search', command=self.search)
        self.btnUpdate = tk.Button(master=self.buttonFrame, text='Update Selected', command=self.updateSelected)
        self.btnDelete = tk.Button(master=self.buttonFrame, text='Delete Selected', command=self.tryDelete)
        # self.btnXlSheet = tk.Button(master=self.buttonFrame, text='Create Spreadsheet')
        # TODO add XL function

        self.frame1.grid(row=0, column=0, sticky='NSEW')
        self.lblWorker.grid(row=0, column=0)
        self.lblSelectedWorker.grid(row=1, column=0)
        self.sbxWorker.grid(row=0, column=1, rowspan=3)
        self.workerScroller.grid(row=0, column=2, rowspan=3)
        self.lblDate.grid(row=3, column=0, pady=3)
        self.entDate.grid(row=3, column=1, pady=3, sticky='W')

        companyFrameSpan = len(self.companyWidgets) // 3
        if companyFrameSpan < 1:
            companyFrameSpan = 1
        self.companyFrame.grid(row=0, column=1, columnspan=companyFrameSpan, sticky='NSEW')
        rowStart = 0
        columnStart = 0
        for tup in self.companyWidgets:
            rowOffset = (tup[0] % 3) * 2
            columnOffset = tup[0] // 3
            tup[1].grid(row=rowStart + rowOffset, column=columnStart + columnOffset)
            tup[2].grid(row=rowStart + rowOffset + 1, column=columnStart + columnOffset)
        self.calcBtn.grid(row=6, column=0)
        self.lblTotal.grid(row=6, column=1)
        self.btnClear.grid(row=6, column=2)
        self.lblPaidBy.grid(row=6, column=3)
        self.entPaidBy.grid(row=6, column=4)

        self.historyFrame.grid(row=1, column=0, columnspan=companyFrameSpan, sticky='NSEW')
        self.sbxHistory.grid(row=0, column=0, padx=3)
        self.historyScroller.grid(row=0, column=1)
        self.backBtn.grid(row=2, column=0, sticky='SW')

        self.buttonFrame.grid(row=1, column=companyFrameSpan, sticky='NSEW')
        self.btnAdd.grid(row=0, column=0)
        self.btnEdit.grid(row=0, column=1)
        self.btnSearch.grid(row=1, column=0)
        self.btnUpdate.grid(row=1, column=1)
        self.btnDelete.grid(row=2, column=0)
        # self.btnXlSheet.grid(row=2, column=1)

        self.deleting = False

    def draw(self):
        self.grid(sticky='NSEW')
        self.addHistory()
        self.addWorkers()
        self.dateText.set(makeStr(Date.today()))
        self.paidByText.set('Dave')

    def undraw(self):
        self.grid_forget()

    def selectWorker(self, event):
        index = self.sbxWorker.curselection()[0]
        selected = self.sbxWorker.get(index)
        self.lblSelectedWorker.config(text=selected)

    def addHistory(self):
        self.sbxHistory.delete(0, tk.END)
        data = self.workerList.getAllHoursData()
        if data is None:
            return
        printData = []
        for row in data:
            printData.append([row[0], makeStr(row[1]), row[2], row[3], row[4]])
        self.sbxHistory.insert(tk.END, *printData)

    def addWorkers(self):
        self.sbxWorker.delete(0, tk.END)
        self.sbxWorker.insert(tk.END, *self.workerList.getNames())

    def calcTotal(self):
        workerName = self.sbxWorker.get(tk.ACTIVE)
        if workerName not in self.workerList.getNames():
            return
        worker = self.workerList[workerName]
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        payRate = worker.getPayRate(date)
        if payRate is None:
            self.lblTotal.config(text='No Pay Data for that Date')
            return
        total = 0
        for tup in self.companyWidgets:
            txt = tup[3].get()
            if txt == '':
                continue
            try:
                txt = float(txt)
            except ValueError:
                continue
            total += txt * payRate
        self.lblTotal.config(text=str(total))

    def search(self):
        workerName = self.sbxWorker.get(tk.ACTIVE)
        if workerName not in self.workerList.getNames():
            return
        worker = self.workerList[workerName]
        data = worker.getHoursData()
        printData = []
        for row in data:
            printData.append([row[0], makeStr(row[1]), row[2], row[3], row[4]])
        self.sbxHistory.delete(0, tk.END)
        self.sbxHistory.insert(tk.END, *printData)

    def addPayment(self):
        workerName = self.sbxWorker.get(tk.ACTIVE)
        if workerName not in self.workerList.getNames():
            return
        worker = self.workerList[workerName]
        try:
            date = makeDate(self.dateText.get())
        except:
            print('Date Invalid')
            return
        hoursData = []
        for tup in self.companyWidgets:
            name = tup[1].cget('text')
            if tup[3].get() == '':
                continue
            try:
                hours = float(tup[3].get())
            except ValueError:
                print('invalid hour entry')
                continue
            hoursData.append([name, hours])
        if len(hoursData) == 0:
            return
        paidBy = self.paidByText.get()
        worker.addHours(date, hoursData, paidBy)
        self.addHistory()

    def editSelected(self):
        selected = self.sbxHistory.get(tk.ACTIVE)
        for i in range(self.sbxWorker.size()):
            if self.sbxWorker.get(i) == selected[2]:
                self.sbxWorker.activate(i)
                self.sbxWorker.see(i)
                self.lblSelectedWorker.config(text=selected[2])
        self.selectedWorkerName = selected[2]
        self.selectedID = selected[0]
        self.dateText.set(selected[1])
        self.paidByText.set(selected[4])
        if type(selected[3][0]) is tuple:
            for hoursTup in selected[3]:
                for tup in self.companyWidgets:
                    if hoursTup[0] == tup[1].cget('text'):
                        tup[3].set(str(hoursTup[1]))
        else:
            for tup in self.companyWidgets:
                if selected[3][0] == tup[1].cget('text'):
                    tup[3].set(str(selected[3][1]))

    def updateSelected(self):
        workerName = self.sbxWorker.get(tk.ACTIVE)
        worker = self.workerList[workerName]
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        hoursData = []
        for tup in self.companyWidgets:
            name = tup[1].cget('text')
            if tup[3].get() == '':
                continue
            try:
                hours = float(tup[3].get())
            except ValueError:
                print('invalid hour entry')
                continue
            hoursData.append([name, hours])
        paidBy = self.paidByText.get()
        if workerName != self.selectedWorkerName:
            oldWorker = self.workerList[self.selectedWorkerName]
            oldWorker.deleteHours(self.selectedID)
            worker.addHours(date, hoursData, paidBy)
        else:
            worker.updateHours(self.selectedID, date, hoursData, paidBy)
        self.addHistory()

    def tryDelete(self):
        if self.deleting:
            self.deleteSelected()
            return
        timer = threading.Timer(5, self.cancelDelete)
        timer.start()
        self.deleting = True

    def cancelDelete(self):
        self.deleting = False

    def deleteSelected(self):
        selected = self.sbxHistory.get(tk.ACTIVE)
        ident = selected[0]
        worker = self.workerList[selected[2]]
        worker.deleteHours(ident)
        self.addHistory()

    def clearCompanies(self):
        for tup in self.companyWidgets:
            tup[3].set('')
        self.dateText.set(makeStr(Date.today()))


class ManageWorkerMenu(tk.Frame):

    def __init__(self, parent, workerList, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.workerList = workerList
        self.companyList = companyList
        self.config(padx=25, pady=25)

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawMainMenu)

        self.sbxWorker = tk.Listbox(master=self, width=50, height=30)
        self.workerScroller = tk.Scrollbar(master=self)
        self.sbxWorker.configure(yscrollcommand=self.workerScroller.set)
        self.workerScroller.configure(command=self.sbxWorker.yview)

        self.btnEdit = tk.Button(master=self, text='Edit Selected', command=self.editSelected)
        self.btnLoans = tk.Button(master=self, text='Manage Selected Loans', command=self.manageSelectedLoans)
        self.btnAddWorker = tk.Button(master=self, text='Add Worker', command=parent.drawAddWorkerMenu)
        self.btnDeactivate = tk.Button(master=self, text='Deactivate Selected Worker', command=self.deactivateSelected)
        self.btnReactivate = tk.Button(master=self, text='Reactivate Selected Worker', command=self.reactivateSelected)

        self.sbxDeactivated = tk.Listbox(master=self, width=40, height=20)
        self.deactivatedScroller = tk.Scrollbar(master=self)
        self.sbxDeactivated.configure(yscrollcommand=self.deactivatedScroller.set)
        self.deactivatedScroller.configure(command=self.sbxDeactivated.yview)

        self.backBtn.grid(row=6, column=0)
        self.sbxWorker.grid(row=0, column=0, rowspan=6, columnspan=2)
        self.workerScroller.grid(row=0, column=2, rowspan=6)

        self.btnEdit.grid(row=0, column=3)
        self.btnLoans.grid(row=1, column=3)
        self.btnAddWorker.grid(row=2, column=3)
        self.btnDeactivate.grid(row=3, column=3)
        self.btnReactivate.grid(row=4, column=3)

        self.sbxDeactivated.grid(row=5, column=3, rowspan=2)
        self.deactivatedScroller.grid(row=5, column=4, rowspan=2)

    def draw(self):
        self.grid()
        self.updateScrollboxes()

    def undraw(self):
        self.grid_forget()

    def updateScrollboxes(self):
        self.sbxDeactivated.delete(0, tk.END)
        self.sbxWorker.delete(0, tk.END)
        self.sbxWorker.insert(tk.END, *self.workerList.getNames())
        self.sbxDeactivated.insert(tk.END, *self.workerList.getInactiveNames())

    def editSelected(self):
        worker = self.workerList[self.sbxWorker.get(tk.ACTIVE)]
        if worker is None:
            return
        self.parent.drawEditWorkerMenu(worker)

    def manageSelectedLoans(self):
        worker = self.workerList[self.sbxWorker.get(tk.ACTIVE)]
        if worker is None:
            return
        self.parent.drawLoanMenu(worker)

    def deactivateSelected(self):
        worker = self.workerList[self.sbxWorker.get(tk.ACTIVE)]
        worker.deactivate()
        self.updateScrollboxes()

    def reactivateSelected(self):
        worker = self.workerList[self.sbxDeactivated.get(tk.ACTIVE)]
        worker.activate()
        self.updateScrollboxes()


class EditWorkerMenu(tk.Frame):

    def __init__(self, parent, worker, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.worker = worker
        self.companyList = companyList
        self.config(padx=50, pady=40)
        self.grid()

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageWorkersMenu)
        self.btnConfirm = tk.Button(master=self, text='Confirm Changes', command=self.confirmChanges)

        self.lblName = tk.Label(master=self, text='Name:')
        self.nameText = tk.StringVar()
        self.entName = tk.Entry(master=self, textvariable=self.nameText)
        self.nameText.set(str(worker))
        self.lblPay = tk.Label(master=self, text='Pay Rate:')
        self.payText = tk.StringVar()
        self.entPay = tk.Entry(master=self, textvariable=self.payText)
        self.lblDate = tk.Label(master=self, text='Effective:')
        self.dateText = tk.StringVar()
        self.entDate = tk.Entry(master=self, textvariable=self.dateText)

        self.sbxPayRate = tk.Listbox(master=self, width=25, height=12)

        self.btnAdd = tk.Button(master=self, text='Add Pay Rate', command=self.addPayRate)
        self.btnEdit = tk.Button(master=self, text='Edit Selected', command=self.editSelected)
        self.btnUpdate = tk.Button(master=self, text='Update', command=self.updateSelected)
        self.btnDelete = tk.Button(master=self, text='Delete', command=self.tryDelete)

        self.lblRent = tk.Label(master=self, text='Rent Paid to:')
        self.sbxCompanies = tk.Listbox(master=self, width=20, height=10)
        self.sbxCompanies.bind('<<ListboxSelect>>', self.selectCompany)
        self.lblRentPaid = tk.Label(master=self, text='')

        self.sbxRents = tk.Listbox(master=self, width=20, height=10)
        self.lblStartDate = tk.Label(master=self, text='Start Date:')
        self.startDateText = tk.StringVar()
        self.entStartDate = tk.Entry(master=self, width=10, textvariable=self.startDateText)
        self.lblEndDate = tk.Label(master=self, text='End Date:')
        self.endDateText = tk.StringVar()
        self.entEndDate = tk.Entry(master=self, width=10, textvariable=self.endDateText)

        self.btnAddRent = tk.Button(master=self, text='Add', command=self.addRent)
        self.btnEditRent = tk.Button(master=self, text='Edit', command=self.editRent)
        self.btnUpdateRent = tk.Button(master=self, text='Update', command=self.updateRent)
        self.btnDeleteRent = tk.Button(master=self, text='Delete', command=self.tryDeleteRent)

        self.sbxCompanies.insert(tk.END, 'None')
        self.sbxCompanies.insert(tk.END, *companyList.getNames())

        self.lblName.grid(row=0, column=0)
        self.entName.grid(row=0, column=1)
        self.lblPay.grid(row=0, column=2)
        self.entPay.grid(row=0, column=3, columnspan=2)
        self.lblDate.grid(row=1, column=2)
        self.entDate.grid(row=1, column=3, columnspan=2)

        self.sbxPayRate.grid(row=2, column=0, rowspan=2, columnspan=2)

        self.btnAdd.grid(row=2, column=3)
        self.btnEdit.grid(row=2, column=4)
        self.btnUpdate.grid(row=3, column=3)
        self.btnDelete.grid(row=3, column=4)

        self.sbxRents.grid(row=5, rowspan=3, column=0)
        self.lblRent.grid(row=4, column=1, columnspan=2)
        self.sbxCompanies.grid(row=5, column=1, rowspan=3, columnspan=2)
        self.lblRentPaid.grid(row=4, column=3, sticky='W')

        self.lblStartDate.grid(row=4, column=4)
        self.entStartDate.grid(row=5, column=4, sticky='N')
        self.lblEndDate.grid(row=4, column=5)
        self.entEndDate.grid(row=5, column=5, sticky='N')

        self.btnAddRent.grid(row=6, column=4)
        self.btnEditRent.grid(row=6, column=5)
        self.btnUpdateRent.grid(row=7, column=4)
        self.btnDeleteRent.grid(row=7, column=5)

        self.backBtn.grid(row=8, column=0, sticky='SW')
        self.btnConfirm.grid(row=8, column=5)

        self.fillScreen()
        self.deleting = False
        self.deletingRent = False

    def fillScreen(self):
        self.sbxPayRate.delete(0, tk.END)
        data = self.worker.getAllPayRates()
        printData = []
        for row in data:
            printData.append([row[0], makeStr(row[1]), row[2]])
        self.sbxPayRate.insert(tk.END, *printData)
        rentData = self.worker.getRentOverview()
        printRentData = []
        for row in rentData:
            if row[1] is not None:
                startDate = makeStr(row[1])
            else:
                startDate = None
            if row[2] is not None:
                endDate = makeStr(row[2])
            else:
                endDate = None
            printRentData.append([row[0], startDate, endDate, row[3]])
        self.sbxRents.delete(0, tk.END)
        self.sbxRents.insert(tk.END, *printRentData)
        mostRecent = printRentData[0]
        self.startDateText.set(mostRecent[1])
        self.endDateText.set(mostRecent[2])
        self.lblRentPaid.config(text=str(mostRecent[3]))

    def undraw(self):
        self.destroy()

    def confirmChanges(self):
        newName = self.nameText.get()
        if newName != self.worker.name:
            self.worker.updateName(self.nameText.get())
        self.parent.drawManageWorkersMenu()

    def addPayRate(self):
        try:
            date = makeDate(self.dateText.get())
        except TypeError:
            return
        try:
            rate = float(self.payText.get())
        except ValueError:
            return
        self.worker.addPayRate(date, rate)
        self.fillScreen()

    def editSelected(self):
        selected = self.sbxPayRate.get(tk.ACTIVE)
        self.selectedID = int(selected[0])
        self.dateText.set(selected[1])
        self.payText.set(str(selected[2]))

    def updateSelected(self):
        try:
            date = makeDate(self.dateText.get())
        except TypeError:
            return
        try:
            rate = float(self.payText.get())
        except ValueError:
            return
        self.worker.editPayRate(self.selectedID, date, rate)
        self.fillScreen()

    def tryDelete(self):
        if self.deleting:
            self.deleteSelected()
            return
        timer = threading.Timer(5, self.cancelDelete)
        timer.start()
        self.deleting = True

    def cancelDelete(self):
        self.deleting = False

    def deleteSelected(self):
        selected = self.sbxPayRate.get(tk.ACTIVE)
        ident = int(selected[0])
        self.worker.deletePayData(ident)
        self.fillScreen()

    def selectCompany(self, event):
        index = self.sbxCompanies.curselection()[0]
        selected = self.sbxCompanies.get(index)
        self.lblRentPaid.config(text=selected)

    def addRent(self):
        startDate = self.startDateText.get()
        endDate = self.endDateText.get()
        company = self.lblRentPaid.cget('text')
        if startDate == '' or startDate == 'None':
            startDate = None
        else:
            try:
                startDate = makeDate(startDate)
            except:
                return
        if endDate == '' or endDate == 'None':
            endDate = None
        else:
            try:
                endDate = makeDate(endDate)
            except:
                return
        if company == 'None':
            self.worker.updateRent(self.selectedRentID, None, None, None)
            self.fillScreen()
            return
        else:
            if startDate is None:
                return
        self.worker.addRent(startDate, endDate, company)
        self.fillScreen()

    def editRent(self):
        selected = self.sbxRents.get(tk.ACTIVE)
        self.selectedRentID = selected[0]
        self.lblRentPaid.config(text=selected[3])
        self.startDateText.set(str(selected[1]))
        self.endDateText.set(str(selected[2]))

    def updateRent(self):
        startDate = self.startDateText.get()
        endDate = self.endDateText.get()
        company = self.lblRentPaid.cget('text')
        if startDate == '' or startDate == 'None':
            startDate = None
        else:
            try:
                startDate = makeDate(startDate)
            except:
                return
        if endDate == '' or endDate == 'None':
            endDate = None
        else:
            try:
                endDate = makeDate(endDate)
            except:
                return
        if company == 'None':
            self.worker.updateRent(self.selectedRentID, None, None, None)
            self.fillScreen()
            return
        else:
            if startDate is None:
                return
        self.worker.updateRent(self.selectedRentID, startDate, endDate, company)
        self.fillScreen()

    def tryDeleteRent(self):
        if self.deletingRent:
            self.deleteSelectedRent()
            return
        timer = threading.Timer(5, self.cancelDeleteRent)
        timer.start()
        self.deletingRent = True

    def cancelDeleteRent(self):
        self.deletingRent = False

    def deleteSelectedRent(self):
        selected = self.sbxRents.get(tk.ACTIVE)
        ident = int(selected[0])
        self.worker.deleteRentData(ident)
        self.fillScreen()


class ManageLoanMenu(tk.Frame):

    def __init__(self, parent, worker, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.worker = worker
        self.args = args
        self.kwargs = kwargs
        self.grid()
        self.config(padx=75, pady=75)

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageWorkersMenu)
        self.lblName = tk.Label(master=self, text=str(worker), font=tkFont.Font(size=20))
        self.lblDebt = tk.Label(master=self, text='Debt: $', font=tkFont.Font(size=20))

        self.subjText = tk.StringVar()
        self.entSubject = tk.Entry(master=self, textvariable=self.subjText)
        self.amtText = tk.StringVar()
        self.entAmount = tk.Entry(master=self, textvariable=self.amtText)
        self.dateText = tk.StringVar()
        self.entDate = tk.Entry(master=self, textvariable=self.dateText)
        self.memoText = tk.StringVar()
        self.entMemo = tk.Entry(master=self, textvariable=self.memoText)

        self.sbxHistory = tk.Listbox(master=self, width=30, height=10)
        self.historyScroller = tk.Scrollbar(master=self)
        self.sbxHistory.configure(yscrollcommand=self.historyScroller.set)
        self.historyScroller.configure(command=self.sbxHistory.yview)

        self.sbxDebts = tk.Listbox(master=self, width=15, height=10)
        self.debtsScroller = tk.Scrollbar(master=self)
        self.sbxHistory.configure(yscrollcommand=self.debtsScroller.set)
        self.debtsScroller.configure(command=self.sbxDebts.yview)
        self.sbxDebts.bind('<<ListboxSelect>>', self.selectDebt)

        self.btnAddLoan = tk.Button(master=self, text='Add Loan', command=self.addLoan)
        self.btnAddPayment = tk.Button(master=self, text='Add Payment', command=self.addPayment)
        self.btnEditSingle = tk.Button(master=self, text='Edit Selected (single)', command=self.editSelected)
        self.btnUpdate = tk.Button(master=self, text='Update', command=self.updateSelected)

        self.sbxRecurring = tk.Listbox(master=self, width=30, height=5)
        self.recurringScroller = tk.Scrollbar(master=self)
        self.sbxRecurring.configure(yscrollcommand=self.recurringScroller.set)
        self.recurringScroller.configure(command=self.sbxRecurring.yview)

        self.btnSearch = tk.Button(master=self, text='Search History', command=self.searchHistory)
        self.btnSpreadsheet = tk.Button(master=self, text='Create Spreadsheet', command=self.createSpreadsheet)
        self.btnAddRecurring = tk.Button(master=self, text='Add Recurring', command=self.addRecurring)
        self.btnEditRecurring = tk.Button(master=self, text='Edit Selected (recurring)', command=self.editRecurring)

        self.lblName.grid(row=0, column=0, columnspan=2)
        self.lblDebt.grid(row=0, column=3, columnspan=3)

        tk.Label(master=self, text='Subject:').grid(row=1, column=0)
        self.entSubject.grid(row=1, column=1)
        tk.Label(master=self, text='Amount: $').grid(row=1, column=3, columnspan=2)
        self.entAmount.grid(row=1, column=5)
        tk.Label(master=self, text='Date:').grid(row=2, column=0)
        self.entDate.grid(row=2, column=1)
        tk.Label(master=self, text='Memo:').grid(row=2, column=3, columnspan=2)
        self.entMemo.grid(row=2, column=5)

        self.sbxHistory.grid(row=3, column=0, rowspan=4, columnspan=2, pady=10)
        self.historyScroller.grid(row=3, column=2, rowspan=4)
        self.sbxDebts.grid(row=3, column=3, rowspan=4)
        self.debtsScroller.grid(row=3, column=4, rowspan=4)

        self.btnAddLoan.grid(row=3, column=5)
        self.btnAddPayment.grid(row=4, column=5)
        self.btnEditSingle.grid(row=5, column=5)
        self.btnUpdate.grid(row=6, column=5)

        self.sbxRecurring.grid(row=7, column=0, rowspan=2, columnspan=2)
        self.recurringScroller.grid(row=7, column=2, rowspan=2)

        self.btnAddRecurring.grid(row=7, column=3)
        self.btnEditRecurring.grid(row=8, column=3)
        self.btnSearch.grid(row=7, column=5)
        self.btnSpreadsheet.grid(row=8, column=5)
        self.backBtn.grid(row=9, column=0, pady=10)

        self.fillBoxes()
        self.dateText.set(makeStr(Date.today()))

    def undraw(self):
        self.destroy()

    def tempUndraw(self):
        self.grid_forget()

    def draw(self):
        self.subMenu.undraw()
        self.grid()
        self.fillBoxes()

    def fillBoxes(self):
        data = self.worker.getLoanData()
        printData = []
        for row in data:
            printData.append([row[0], makeStr(row[1]), row[2], row[3], row[4]])
        self.sbxHistory.delete(0, tk.END)
        self.sbxHistory.insert(tk.END, *printData)
        debts = {}
        for row in data:
            if row[3] in debts:
                debts[row[3]] = round(debts[row[3]] + row[2], ndigits=2)
            else:
                debts[row[3]] = row[2]
        debtsList = []
        total = 0
        for key, amt in debts.items():
            if amt != 0:
                debtsList.append([key, amt])
                total += amt
        debtsList.sort(key=itemgetter(1))
        self.sbxDebts.delete(0, tk.END)
        self.sbxDebts.insert(tk.END, *debtsList)
        self.lblDebt.config(text=str(round(total, ndigits=2)))
        recurring = self.worker.getRecurring()
        drawnRecurring = []
        for row in recurring:
            if row[2] is None:
                drawnRecurring.append([row[0], makeStr(row[1]), row[2], row[3], row[4], row[5]])
                continue
            elif (Date.today() - row[2]).days > 62 and Date.today() > row[2]:
                continue
            elif row[4] == 0:
                continue
            else:
                drawnRecurring.append([row[0], makeStr(row[1]), makeStr(row[2]), row[3], row[4], row[5]])
        self.sbxRecurring.delete(0, tk.END)
        self.sbxRecurring.insert(tk.END, *drawnRecurring)

    def addRecurring(self):
        self.tempUndraw()
        self.subMenu = AddRecurringMenu(self, self.parent, self.worker, *self.args, **self.kwargs)

    def editRecurring(self):
        recurringLoan = self.sbxRecurring.get(tk.ACTIVE)
        if recurringLoan is None or recurringLoan == '':
            return
        self.tempUndraw()
        self.subMenu = EditRecurringMenu(self, self.parent, self.worker, recurringLoan, *self.args, **self.kwargs)

    def addLoan(self):
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            amount = float(self.amtText.get())
        except ValueError:
            return
        if amount > 0:
            amount *= -1
        subject = self.subjText.get()
        if subject == '':
            return
        memo = self.memoText.get()
        self.worker.addLoan(date, amount, subject, memo)
        self.fillBoxes()

    def addPayment(self):
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            amount = float(self.amtText.get())
        except ValueError:
            return
        if amount < 0:
            amount *= -1
        subject = self.subjText.get()
        if subject == '':
            return
        memo = self.memoText.get()
        self.worker.addPayment(date, amount, subject, memo)
        self.fillBoxes()

    def editSelected(self):
        selected = self.sbxHistory.get(tk.ACTIVE)
        if selected is None or selected == '':
            return
        self.selectedID = int(selected[0])
        self.dateText.set(selected[1])
        self.amtText.set(str(selected[2]))
        self.subjText.set(selected[3])
        self.memoText.set(selected[4])

    def updateSelected(self):
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            amt = float(self.amtText.get())
        except ValueError:
            return
        subject = self.subjText.get()
        memo = self.memoText.get()
        self.worker.editLoan(self.selectedID, date, amt, subject, memo)
        self.fillBoxes()

    def searchHistory(self):
        subject = self.subjText.get()
        if subject == '':
            self.fillBoxes()
        else:
            data = self.worker.searchLoanData(subject)
            for row in data:
                row[1] = makeStr(row[1])
            self.sbxHistory.delete(0, tk.END)
            self.sbxHistory.insert(tk.END, *data)

    def selectDebt(self, event):
        index = self.sbxDebts.curselection()[0]
        selected = self.sbxDebts.get(index)
        self.subjText.set(selected[0])

    def createSpreadsheet(self):
        desktopPath = ''
        if sys.platform == 'darwin' or sys.platform == 'linux':
            desktopPath = os.path.expanduser('~/Desktop')
        elif sys.platform == 'win32' or sys.platform == 'cygwin':
            try:
                desktopPath = os.path.join(os.environ['HOMEPATH'], 'Desktop')
            except:
                desktopPath = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        workbook = xlsxwriter.Workbook(os.path.join(desktopPath, self.worker.name + '.xlsx'))
        worksheet = workbook.add_worksheet()
        data = self.worker.getLoanData()
        data.sort(key=itemgetter(1))
        for row in data:
            row[1] = makeStr(row[1])
        worksheet.write(0, 0, 'Date')
        worksheet.write(0, 1, 'Item')
        worksheet.write(0, 2, 'Amount')
        worksheet.write(0, 3, 'Total')
        worksheet.write(0, 4, 'Note')
        rowNum = 1
        for row in data:
            worksheet.write(rowNum, 0, row[1])
            worksheet.write(rowNum, 1, row[3])
            worksheet.write(rowNum, 2, row[2])
            worksheet.write(rowNum, 3, '=SUM(C1:C' + str(rowNum) + ')')
            worksheet.write(rowNum, 4, row[4])
            rowNum += 1
        workbook.close()


class AddWorkerMenu(tk.Frame):

    def __init__(self, parent, workerList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.workerList = workerList
        self.config(padx=150, pady=150)
        self.grid()

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageWorkersMenu)
        self.btnConfirm = tk.Button(master=self, text='Confirm', command=self.confirm)

        self.lblName = tk.Label(master=self, text='Name:')
        self.nameText = tk.StringVar()
        self.entName = tk.Entry(master=self, textvariable=self.nameText)
        self.lblPay = tk.Label(master=self, text='Pay Rate:')
        self.payText = tk.StringVar()
        self.entPay = tk.Entry(master=self, textvariable=self.payText)
        self.lblDate = tk.Label(master=self, text='Effective:')
        self.dateText = tk.StringVar()
        self.entDate = tk.Entry(master=self, textvariable=self.dateText)

        self.lblName.grid(row=0, column=0)
        self.entName.grid(row=0, column=1)
        self.lblPay.grid(row=0, column=2)
        self.entPay.grid(row=0, column=3)
        self.lblDate.grid(row=1, column=2)
        self.entDate.grid(row=1, column=3)

        self.backBtn.grid(row=4, column=0)
        self.btnConfirm.grid(row=4, column=3)

    def undraw(self):
        self.destroy()

    def confirm(self):
        name = self.nameText.get()
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            rate = float(self.payText.get())
        except ValueError:
            return
        worker = self.workerList.addWorker(name)
        worker.addPayRate(date, rate)
        self.parent.drawManageWorkersMenu()


class EditRecurringMenu(tk.Frame):

    def __init__(self, loanMenu, parent, worker, recurringLoan, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.loanMenu = loanMenu
        self.parent = parent
        self.worker = worker
        self.recurringID = recurringLoan[0]
        self.grid(padx=125, pady=200)

        self.backBtn = tk.Button(master=self, text='Back', command=self.back)
        self.btnConfirm = tk.Button(master=self, text='Confirm', command=self.confirm)

        self.lblMemo = tk.Label(master=self, text='Memo:')
        self.memoText = tk.StringVar()
        self.entMemo = tk.Entry(master=self, textvariable=self.memoText)
        self.lblAmt = tk.Label(master=self, text='Amount:')
        self.amtText = tk.StringVar()
        self.entAmt = tk.Entry(master=self, textvariable=self.amtText)
        self.lblHelp = tk.Label(master=self, text='Negative for Loan, Positive for Payment')
        self.lblStartDate = tk.Label(master=self, text='Start Date:')
        self.startDateText = tk.StringVar()
        self.entStartDate = tk.Entry(master=self, textvariable=self.startDateText)
        self.lblEndDate = tk.Label(master=self, text='End Date:')
        self.endDateText = tk.StringVar()
        self.entEndDate = tk.Entry(master=self, textvariable=self.endDateText)

        self.sbxFreq = tk.Listbox(master=self, width=30, height=4)
        self.freqScroller = tk.Scrollbar(master=self)
        self.sbxFreq.configure(yscrollcommand=self.freqScroller.set)
        self.freqScroller.configure(command=self.sbxFreq.yview)
        self.sbxFreq.insert(tk.END, *['Daily', 'Weekly', 'Monthly', 'Yearly'])

        self.lblMemo.grid(row=1, column=0)
        self.entMemo.grid(row=1, column=1)
        self.lblAmt.grid(row=1, column=2)
        self.entAmt.grid(row=1, column=3)
        self.lblHelp.grid(row=0, column=2, columnspan=2)
        self.lblStartDate.grid(row=2, column=0)
        self.entStartDate.grid(row=2, column=1)
        self.lblEndDate.grid(row=2, column=2)
        self.entEndDate.grid(row=2, column=3)

        self.sbxFreq.grid(row=3, column=0, columnspan=2)
        self.freqScroller.grid(row=3, column=2)

        self.backBtn.grid(row=4, column=0)
        self.btnConfirm.grid(row=4, column=3)

        self.startDateText.set(recurringLoan[1])
        if recurringLoan[2] is not None:
            self.endDateText.set(recurringLoan[2])
        self.memoText.set(recurringLoan[3])
        self.amtText.set(str(recurringLoan[4]))
        for i in range(self.sbxFreq.size()):
            if self.sbxFreq.get(i) == recurringLoan[5]:
                self.sbxFreq.activate(i)
                break

    def back(self):
        self.loanMenu.draw()

    def undraw(self):
        self.destroy()

    def confirm(self):
        try:
            startDate = makeDate(self.startDateText.get())
        except:
            return
        endDate = self.endDateText.get()
        if endDate == '' or endDate == 'None':
            endDate = None
        else:
            try:
                endDate = makeDate(endDate)
            except:
                return
        try:
            amount = float(self.amtText.get())
        except ValueError:
            return
        memo = self.memoText.get()
        freq = self.sbxFreq.get(tk.ACTIVE)
        self.worker.updateRecurring(self.recurringID, startDate, endDate, memo, amount, freq)
        self.back()


class AddRecurringMenu(tk.Frame):

    def __init__(self, loanMenu, parent, worker, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.loanMenu = loanMenu
        self.parent = parent
        self.worker = worker
        self.grid(padx=125, pady=200)

        self.backBtn = tk.Button(master=self, text='Back', command=self.back)
        self.btnConfirm = tk.Button(master=self, text='Confirm', command=self.confirm)

        self.lblMemo = tk.Label(master=self, text='Memo:')
        self.memoText = tk.StringVar()
        self.entMemo = tk.Entry(master=self, textvariable=self.memoText)
        self.lblAmt = tk.Label(master=self, text='Amount:')
        self.amtText = tk.StringVar()
        self.entAmt = tk.Entry(master=self, textvariable=self.amtText)
        self.lblHelp = tk.Label(master=self, text='Negative for Loan, Positive for Payment')
        self.lblStartDate = tk.Label(master=self, text='Start Date:')
        self.startDateText = tk.StringVar()
        self.entStartDate = tk.Entry(master=self, textvariable=self.startDateText)
        self.lblEndDate = tk.Label(master=self, text='End Date:')
        self.endDateText = tk.StringVar()
        self.entEndDate = tk.Entry(master=self, textvariable=self.endDateText)

        self.sbxFreq = tk.Listbox(master=self, width=30, height=4)
        self.freqScroller = tk.Scrollbar(master=self)
        self.sbxFreq.configure(yscrollcommand=self.freqScroller.set)
        self.freqScroller.configure(command=self.sbxFreq.yview)
        self.sbxFreq.insert(tk.END, *['Daily', 'Weekly', 'Monthly', 'Yearly'])

        self.lblMemo.grid(row=1, column=0)
        self.entMemo.grid(row=1, column=1)
        self.lblAmt.grid(row=1, column=2)
        self.entAmt.grid(row=1, column=3)
        self.lblHelp.grid(row=0, column=2, columnspan=2)
        self.lblStartDate.grid(row=2, column=0)
        self.entStartDate.grid(row=2, column=1)
        self.lblEndDate.grid(row=2, column=2)
        self.entEndDate.grid(row=2, column=3)

        self.sbxFreq.grid(row=3, column=0, columnspan=2)
        self.freqScroller.grid(row=3, column=2)

        self.backBtn.grid(row=4, column=0)
        self.btnConfirm.grid(row=4, column=3)

        self.startDateText.set(makeStr(Date.today()))

    def undraw(self):
        self.destroy()

    def back(self):
        self.destroy()
        self.loanMenu.draw()

    def confirm(self):
        try:
            startDate = makeDate(self.startDateText.get())
        except:
            return
        endDate = self.endDateText.get()
        if endDate == '' or endDate == 'None':
            endDate = None
        else:
            try:
                endDate = makeDate(endDate)
            except:
                return
        try:
            amount = float(self.amtText.get())
        except ValueError:
            return
        memo = self.memoText.get()
        freq = self.sbxFreq.get(tk.ACTIVE)
        self.worker.addRecurring(startDate, endDate, memo, amount, freq)
        self.back()


class ManageCompaniesMenu(tk.Frame):

    def __init__(self, parent, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.config(padx=175, pady=80)
        self.companyList = companyList

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawMainMenu)

        self.sbxCompanies = tk.Listbox(master=self, width=30, height=20)
        self.compannyScroller = tk.Scrollbar(master=self)
        self.sbxCompanies.configure(yscrollcommand=self.compannyScroller.set)
        self.compannyScroller.configure(command=self.sbxCompanies.yview)

        self.btnEdit = tk.Button(master=self, text='Edit Selected Settings', command=self.editSelected)
        self.btnView = tk.Button(master=self, text='View Details/Manage', command=self.viewSelected)
        self.btnAdd = tk.Button(master=self, text='Add Company', command=parent.drawAddCompanyMenu)
        self.btnDeactivate = tk.Button(master=self, text='Deactivate Selected', command=self.deactivateSelected)
        self.btnReactivate = tk.Button(master=self, text='Reactivate Selected', command=self.reactivateSelected)

        self.sbxDeactivated = tk.Listbox(master=self, width=20, height=5)
        self.deactivatedScroller = tk.Scrollbar(master=self)
        self.sbxDeactivated.configure(yscrollcommand=self.deactivatedScroller.set)
        self.deactivatedScroller.configure(command=self.sbxDeactivated.yview)

        self.sbxCompanies.grid(row=0, column=0, rowspan=5)
        self.compannyScroller.grid(row=0, column=1, rowspan=5)

        self.btnView.grid(row=0, column=2)
        self.btnEdit.grid(row=1, column=2)
        self.btnAdd.grid(row=2, column=2)
        self.btnDeactivate.grid(row=3, column=2)
        self.btnReactivate.grid(row=4, column=2)

        self.sbxDeactivated.grid(row=5, column=1, rowspan=2, columnspan=2)
        self.deactivatedScroller.grid(row=5, column=3, rowspan=2)
        self.backBtn.grid(row=6, column=0, sticky='SW')


    def draw(self):
        self.grid()
        self.fillBoxes()

    def undraw(self):
        self.grid_forget()

    def fillBoxes(self):
        self.sbxCompanies.delete(0, tk.END)
        self.sbxDeactivated.delete(0, tk.END)
        self.sbxCompanies.insert(tk.END, *self.companyList.getActiveNames())
        self.sbxDeactivated.insert(tk.END, *self.companyList.getInactiveNames())

    def editSelected(self):
        company = self.companyList[self.sbxCompanies.get(tk.ACTIVE)]
        if company is None:
            return
        self.parent.drawEditCompanyMenu(company)

    def viewSelected(self):
        company = self.companyList[self.sbxCompanies.get(tk.ACTIVE)]
        if company is None:
            return
        self.parent.drawViewCompanyMenu(company)

    def deactivateSelected(self):
        company = self.companyList[self.sbxCompanies.get(tk.ACTIVE)]
        if company is None:
            return
        company.deactivate()
        self.fillBoxes()

    def reactivateSelected(self):
        company = self.companyList[self.sbxDeactivated.get(tk.ACTIVE)]
        if company is None:
            return
        company.activate()
        self.fillBoxes()


class AddCompanyMenu(tk.Frame):

    def __init__(self, parent, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.companyList = companyList
        self.config(padx=70, pady=150)
        self.grid()

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageCompaniesMenu)
        self.confirmBtn = tk.Button(master=self, text='Confirm', command=self.confirm)

        self.lblName = tk.Label(master=self, text='Name:')
        self.nameText = tk.StringVar()
        self.entName = tk.Entry(master=self, textvariable=self.nameText)
        self.lblClearing = tk.Label(master=self, text='Clears Through:')
        self.lblClearsThrough = tk.Label(master=self, text='None')

        self.sbxClearsThrough = tk.Listbox(master=self, width=40, height=15)
        self.clearingScroller = tk.Scrollbar(master=self)
        self.sbxClearsThrough.configure(yscrollcommand=self.clearingScroller.set)
        self.clearingScroller.configure(command=self.sbxClearsThrough.yview)
        self.sbxClearsThrough.bind('<<ListboxSelect>>', self.selectClearing)
        self.sbxClearsThrough.insert(tk.END, 'None')
        self.sbxClearsThrough.insert(tk.END, *companyList.getNames())

        self.lblName.grid(row=0, column=0)
        self.entName.grid(row=0, column=1)
        self.sbxClearsThrough.grid(row=0, column=2, rowspan=3, padx=15)
        self.clearingScroller.grid(row=0, column=3, rowspan=3)
        self.lblClearsThrough.grid(row=1, column=4)
        self.lblClearing.grid(row=1, column=1, sticky='E')
        self.backBtn.grid(row=3, column=0)
        self.confirmBtn.grid(row=3, column=2, sticky='SE')

    def undraw(self):
        self.destroy()

    def confirm(self):
        name = self.nameText.get()
        clearsThrough = self.lblClearsThrough.cget('text')
        if clearsThrough == 'None':
            clearsThrough = None
        company = self.companyList.addCompany(name)
        company.setClearsThrough(clearsThrough)
        self.parent.drawManageCompaniesMenu()

    def selectClearing(self, event):
        index = self.sbxClearsThrough.curselection()[0]
        selected = self.sbxClearsThrough.get(index)
        self.lblClearsThrough.config(text=selected)


class EditCompanyMenu(tk.Frame):

    def __init__(self, parent, company, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.company = company
        self.companyList = companyList
        self.config(padx=70, pady=150)
        self.grid()

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageCompaniesMenu)
        self.confirmBtn = tk.Button(master=self, text='Confirm', command=self.confirm)

        self.lblName = tk.Label(master=self, text='Name:')
        self.nameText = tk.StringVar()
        self.entName = tk.Entry(master=self, textvariable=self.nameText)
        self.lblClearing = tk.Label(master=self, text='Clears Through:')
        self.lblClearsThrough = tk.Label(master=self, text='')

        self.sbxClearsThrough = tk.Listbox(master=self, width=40, height=15)
        self.clearingScroller = tk.Scrollbar(master=self)
        self.sbxClearsThrough.configure(yscrollcommand=self.clearingScroller.set)
        self.clearingScroller.configure(command=self.sbxClearsThrough.yview)
        self.sbxClearsThrough.insert(tk.END, 'None')
        self.sbxClearsThrough.insert(tk.END, *[name for name in companyList.getNames() if name != company.name])
        self.sbxClearsThrough.bind('<<ListboxSelect>>', self.selectClearing)

        self.lblName.grid(row=0, column=0)
        self.entName.grid(row=0, column=1)
        self.sbxClearsThrough.grid(row=0, column=2, rowspan=3, padx=15)
        self.clearingScroller.grid(row=0, column=3, rowspan=3)
        self.lblClearsThrough.grid(row=1, column=4)
        self.lblClearing.grid(row=1, column=1, sticky='E')
        self.backBtn.grid(row=3, column=0)
        self.confirmBtn.grid(row=3, column=2, sticky='SE')

        if company.clearsThrough is None:
            self.lblClearsThrough.config(text='None')
        else:
            self.lblClearsThrough.config(text=company.clearsThrough)
        self.nameText.set(company.name)
        for i in range(self.sbxClearsThrough.size()):
            if self.sbxClearsThrough.get(i) == company.clearsThrough:
                self.sbxClearsThrough.activate(i)
                break

    def undraw(self):
        self.destroy()

    def confirm(self):
        name = self.nameText.get()
        clearsThrough = self.lblClearsThrough.cget('text')
        if clearsThrough == 'None':
            clearsThrough = None
        if name != self.company.name:
            self.company.updateName(name)
        self.company.setClearsThrough(clearsThrough)
        self.parent.drawManageCompaniesMenu()

    def selectClearing(self, event):
        index = self.sbxClearsThrough.curselection()[0]
        selected = self.sbxClearsThrough.get(index)
        self.lblClearsThrough.config(text=selected)


class ViewCompanyMenu(tk.Frame):

    def __init__(self, parent, company, workerList, companyList, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.company = company
        self.companyList = companyList
        self.workerList = workerList
        self.config(padx=20, pady=80)

        self.backBtn = tk.Button(master=self, text='Back', command=parent.drawManageCompaniesMenu)

        self.lblName = tk.Label(master=self, text=company.name)
        self.lblOwedBy = tk.Label(master=self, text='Is Owed By')
        self.lblOwes = tk.Label(master=self, text='Owes')

        self.sbxHistory = tk.Listbox(master=self, width=30, height=20)
        self.historyScroller = tk.Scrollbar(master=self)
        self.sbxHistory.configure(yscrollcommand=self.historyScroller.set)
        self.historyScroller.configure(command=self.sbxHistory.yview)

        self.sbxOwedBy = tk.Listbox(master=self, width=30, height=10)
        self.owedByScroller = tk.Scrollbar(master=self)
        self.sbxOwedBy.configure(yscrollcommand=self.owedByScroller.set)
        self.owedByScroller.configure(command=self.sbxOwedBy.yview)
        self.sbxOwedBy.bind('<<ListboxSelect>>', self.selectOwedBy)

        self.sbxOwes = tk.Listbox(master=self, width=30, height=10)
        self.owesScroller = tk.Scrollbar(master=self)
        self.sbxOwes.configure(yscrollcommand=self.owesScroller.set)
        self.owesScroller.configure(command=self.sbxOwes.yview)
        self.sbxOwes.bind('<<ListboxSelect>>', self.selectOwes)

        self.distFrame = tk.Frame(master=self)
        self.btnTo = tk.Button(master=self.distFrame, text='To', command=self.setTo)
        self.btnFrom = tk.Button(master=self.distFrame, text='From', command=self.setFrom)
        self.lblDistName = tk.Label(master=self.distFrame, text='Name Will Be Here')
        self.lblAmt = tk.Label(master=self.distFrame, text='Amount:')
        self.amtText = tk.StringVar()
        self.entAmt = tk.Entry(master=self.distFrame, width=15, textvariable=self.amtText)
        self.lblDate = tk.Label(master=self.distFrame, text='Date')
        self.dateText = tk.StringVar()
        self.entDate = tk.Entry(master=self.distFrame, width=15, textvariable=self.dateText)
        self.dateText.set(makeStr(Date.today()))
        self.lblToFrom = tk.Label(master=self.distFrame, text='')

        self.btnEditDist = tk.Button(master=self.distFrame, text='Edit Distribution', command=self.editDistribution)
        self.btnAdd = tk.Button(master=self.distFrame, text='Add Payment', command=self.addPayment)
        # self.btnSpreadsheet = tk.Button(master=self.distFrame, text='Create Spreadsheet')
        self.btnUpdate = tk.Button(master=self.distFrame, text='Update Record', command=self.updateDistribution)
        self.btnDelete = tk.Button(master=self.distFrame, text='Delete', command=self.tryDelete)

        self.lblName.grid(row=0, column=0)
        self.lblOwedBy.grid(row=0, column=2)
        self.lblOwes.grid(row=0, column=4)
        self.backBtn.grid(row=3, column=0, sticky='SW')

        self.sbxHistory.grid(row=1, column=0, rowspan=2)
        self.historyScroller.grid(row=1, column=1, rowspan=2)
        self.sbxOwedBy.grid(row=1, column=2)
        self.owedByScroller.grid(row=1, column=3)
        self.sbxOwes.grid(row=1, column=4)
        self.owesScroller.grid(row=1, column=5)
        self.distFrame.grid(row=2, column=2, rowspan=2, columnspan=4)

        self.btnTo.grid(row=0, column=0)
        self.btnFrom.grid(row=0, column=1)
        self.lblToFrom.grid(row=0, column=2)
        self.lblDistName.grid(row=0, column=3, columnspan=2)
        self.lblAmt.grid(row=1, column=0)
        self.entAmt.grid(row=1, column=1)
        self.lblDate.grid(row=1, column=2)
        self.entDate.grid(row=1, column=3)
        self.btnEditDist.grid(row=2, column=0, columnspan=2)
        self.btnAdd.grid(row=2, column=2, columnspan=2)
        self.btnUpdate.grid(row=3, column=0, columnspan=2)
        # self.btnSpreadsheet.grid(row=3, column=2, columnspan=2)
        self.btnDelete.grid(row=3, column=2, columnspan=2)

        self.deleting = False

    def draw(self):
        self.grid()
        self.fillBoxes()

    def undraw(self):
        self.destroy()

    def fillBoxes(self):
        hoursData = self.company.getHoursData()
        distData = self.company.getDistributionData()
        for row in distData:
            row[1] = makeStr(row[1])
        self.sbxHistory.delete(0, tk.END)
        self.sbxHistory.insert(tk.END, *distData)
        self.sbxOwes.delete(0, tk.END)
        self.sbxOwedBy.delete(0, tk.END)
        clearsThroughMe = self.companyList.getClearsThrough(self.company.name)
        someoneClearsThroughMe = len(clearsThroughMe) > 0
        iClearThroughSomeone = self.company.clearsThrough is not None
        clearsThroughMeData = []
        for company in clearsThroughMe:
            for row in company.getHoursData():
                clearsThroughMeData.append(row)
        # myPayouts[paidBy] = amt
        myPayouts = {}
        # clearPayouts[companyName] = {paidBy: amt, }
        clearPayouts = {}
        for row in hoursData:
            if row[5] in myPayouts:
                myPayouts[row[5]] += row[4]
            else:
                myPayouts[row[5]] = row[4]
        if someoneClearsThroughMe:
            for row in clearsThroughMeData:
                # if we already have data for the company
                if row[0] in clearPayouts:
                    if row[5] in clearPayouts[row[0]]:
                        clearPayouts[row[0]][row[5]] += row[4]
                    else:
                        clearPayouts[row[0]][row[5]] = row[4]
                else:
                    clearPayouts[row[0]] = {row[5]: row[4]}
        # debts[company] positive for owes me, negative for I owe
        debts = {}
        if someoneClearsThroughMe:
            for company in clearsThroughMe:
                debts[company.name] = 0
                if company.name in clearPayouts:
                    for paidBy, amt in clearPayouts[company.name].items():
                        debts[company.name] += amt
                        if paidBy in debts:
                            debts[paidBy] -= amt
                        else:
                            debts[paidBy] = amt * -1
        if iClearThroughSomeone:
            myTotalPayout = 0
            for amt in myPayouts.values():
                myTotalPayout += amt
            debts[self.company.clearsThrough] = myTotalPayout * -1
        else:
            for paidBy, amt, in myPayouts.items():
                if paidBy in debts:
                    debts[paidBy] -= amt
                else:
                    debts[paidBy] = (amt * -1)
        for worker in self.workerList.getWorkers():
            rentData = worker.getRentsOwed(self.company.name)
            if rentData is None:
                continue
            else:
                for paidBy, amt in rentData.items():
                    if paidBy in debts:
                        debts[paidBy] -= amt
                    else:
                        debts[paidBy] = (amt * -1)
        for row in distData:
            if row[2] == self.company.name:
                if row[3] in debts:
                    debts[row[3]] += row[4]
                else:
                    debts[row[3]] = row[4]
            elif row[3] == self.company.name:
                if row[2] in debts:
                    debts[row[2]] -= row[4]
                else:
                    debts[row[2]] = row[4] * -1
        for name, amt in debts.items():
            if amt > 0:
                self.sbxOwedBy.insert(tk.END, [name, '$' + str(abs(amt))])
            elif amt < 0:
                self.sbxOwes.insert(tk.END, [name, '$' + str(amt)])

    def setTo(self):
        self.toCompany = True
        self.lblToFrom.config(text='To')

    def setFrom(self):
        self.toCompany = False
        self.lblToFrom.config(text='From')

    def addPayment(self):
        if self.toCompany:
            recipient = self.lblDistName.cget('text')
            donor = self.company.name
        else:
            donor = self.lblDistName.cget('text')
            recipient = self.company.name
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            amt = float(self.amtText.get())
        except ValueError:
            return
        self.company.addDistribution(date, donor, recipient, amt)
        self.fillBoxes()

    def editDistribution(self):
        selected = self.sbxHistory.get(tk.ACTIVE)
        self.selectedID = selected[0]
        self.dateText.set(selected[1])
        if selected[2] == self.company.name:
            self.setTo()
            self.lblDistName.config(text=selected[3])
        else:
            self.setFrom()
            self.lblDistName.config(text=selected[2])
        self.amtText.set(str(selected[4]))

    def updateDistribution(self):
        try:
            date = makeDate(self.dateText.get())
        except:
            return
        try:
            amt = float(self.amtText.get())
        except ValueError:
            return
        if self.toCompany:
            recipient = self.lblDistName.cget('text')
            donor = self.company.name
        else:
            donor = self.lblDistName.cget('text')
            recipient = self.company.name
        self.company.editDistribution(self.selectedID, date, donor, recipient, amt)
        self.fillBoxes()

    def tryDelete(self):
        if self.deleting:
            self.deleteSelected()
            return
        timer = threading.Timer(5, self.cancelDelete)
        timer.start()
        self.deleting = True

    def cancelDelete(self):
        self.deleting = False

    def deleteSelected(self):
        selected = self.sbxHistory.get(tk.ACTIVE)
        ident = selected[0]
        self.company.deleteDistribution(ident)
        self.fillBoxes()

    def selectOwedBy(self, event):
        index = self.sbxOwedBy.curselection()[0]
        selected = self.sbxOwedBy.get(index)
        self.lblDistName.config(text=selected[0])
        self.setFrom()

    def selectOwes(self, event):
        index = self.sbxOwes.curselection()[0]
        selected = self.sbxOwes.get(index)
        self.lblDistName.config(text=selected[0])
        self.setTo()


class MainApplication(tk.Frame):

    def __init__(self, root, workerList, companyList, *args, **kwargs):
        tk.Frame.__init__(self, root, *args, **kwargs)
        self.root = root
        self.grid(sticky='NSEW')
        self.workerList = workerList
        self.companyList = companyList
        self.args = args
        self.kwargs = kwargs

        self.grid_propagate(0)

        self.mainMenu = MainMenu(self, *args, **kwargs)
        self.manageCompaniesMenu = ManageCompaniesMenu(self, companyList, *args, **kwargs)

        self.current = self.mainMenu
        self.mainMenu.draw()

    def killProgram(self):
        self.root.destroy()

    def drawMainMenu(self):
        self.current.undraw()
        self.mainMenu.draw()
        self.current = self.mainMenu

    def drawPayWorkerMenu(self):
        self.current.undraw()
        self.current = PayWorkerMenu(self, self.workerList, self.companyList, *self.args, **self.kwargs)
        self.current.draw()

    def drawManageWorkersMenu(self):
        self.current.undraw()
        self.current = ManageWorkerMenu(self, self.workerList, self.companyList, *self.args, **self.kwargs)
        self.current.draw()

    def drawMasterSpreadsheetMenu(self):
        self.current.undraw()
        self.current = MasterSpreadsheetMenu(self, self.workerList, self.companyList, *self.args, **self.kwargs)
        self.current.draw()

    def drawEditWorkerMenu(self, worker):
        self.current.undraw()
        self.current = EditWorkerMenu(self, worker, self.companyList, *self.args, **self.kwargs)

    def drawLoanMenu(self, worker):
        self.current.undraw()
        self.current = ManageLoanMenu(self, worker, *self.args, **self.kwargs)

    def drawAddWorkerMenu(self):
        self.current.undraw()
        self.current = AddWorkerMenu(self, self.workerList, *self.args, **self.kwargs)

    def drawManageCompaniesMenu(self):
        self.current.undraw()
        self.current = ManageCompaniesMenu(self, self.companyList, *self.args, **self.kwargs)
        self.current.draw()

    def drawAddCompanyMenu(self):
        self.current.undraw()
        self.current = AddCompanyMenu(self, self.companyList, *self.args, **self.kwargs)

    def drawEditCompanyMenu(self, company):
        self.current.undraw()
        self.current = EditCompanyMenu(self, company, self.companyList, *self.args, **self.kwargs)

    def drawViewCompanyMenu(self, company):
        self.current.undraw()
        self.current = ViewCompanyMenu(self, company, self.workerList, self.companyList, *self.args, **self.kwargs)
        self.current.draw()
