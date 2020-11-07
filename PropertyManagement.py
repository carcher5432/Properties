from Database import DbConnection, DbTable
from GuiDesign import *
import json
from datetime import date as Date
from datetime import timedelta
from operator import itemgetter
import os


WINDOW_WIDTH = 800
WINDOW_HEIGHT = 600

window = tk.Tk()
window.geometry(str(WINDOW_WIDTH) + 'x' + str(WINDOW_HEIGHT))
window.resizable(0, 0)
window.title('Property Management')

# for pyInstaller
"""
def resource_path(relative):
    # print(os.environ)
    application_path = os.path.abspath(".")
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the pyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = sys._MEIPASS
    # print(application_path)
    return os.path.join(application_path, relative)
"""


# for running as is
def resource_path(relative):
    path = os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data'), relative)
    return path


DB = DbConnection(resource_path('properties.db'))
""" Drop all tables
DB.execute('DROP TABLE hours')
DB.execute('DROP TABLE loans')
DB.execute('DROP TABLE distributions')
DB.commit()
"""
# hoursDB company and hours: 'comp1/comp2/comp3', 'hrs1/hrs2/hrs3'
hoursDB = DbTable(DB, 'hours', autoCommit=True, date='text', name='text', company='text', hours='text', paidBy='text')
loansDB = DbTable(DB, 'loans', autoCommit=True, date='text', name='text', amount='real', subject='text', memo='text')
distDB = DbTable(DB, 'distributions', autoCommit=True, date='text', donor='text', recipient='text', amount='real')


def makeDate(strDate):
    month, day, year = strDate.split('/')
    return Date(int(year), int(month), int(day))


def makeStr(dateObj):
    year, month, day = dateObj.year, dateObj.month, dateObj.day
    return str(month) + '/' + str(day) + '/' + str(year)


def hoursDataToString(hoursData):
    companies = hoursData[0][0]
    hours = str(hoursData[0][1])
    for tup in hoursData[1:]:
        companies += '/' + tup[0]
        hours += '/' + str(tup[1])
    return companies, hours


def hoursDataFromString(companyString, hoursString):
    companies = companyString.split('/')
    hours = [float(x) for x in hoursString.split('/')]
    data = []
    for i in range(len(companies)):
        data.append([companies[i], float(hours[i])])
    return data


def dateIncrement(date, frequency):
    if frequency == 'Daily':
        return date + timedelta(days=1)
    if frequency == 'Weekly':
        return date + timedelta(days=7)
    if frequency == 'Monthly':
        if date.month == 12:
            return Date(date.year + 1, 1, date.day)
        else:
            return Date(date.year, date.month + 1, date.day)
    if frequency == 'Yearly':
        return Date(date.year + 1, date.month, date.day)


class Worker:

    def __init__(self, d, workerList):
        self.name = d['name']
        self.active = d['active']
        self.workerList = workerList
        # payRates: [int(id), dateObj, float(rate)]
        self.payRates = d['pay rates']
        self.payID = 0
        for row in self.payRates:
            if row[0] >= self.payID:
                self.payID = row[0] + 1
            row[1] = makeDate(row[1])
        # recurring: [int(id), dateObj Start, dateObj End or None, Memo, $Amt, Freq, lastUpdate]
        try:
            self.recurring = d['recurring']
            self.recurringID = 0
            for row in self.recurring:
                if row[0] >= self.recurringID:
                    self.recurringID = row[0] + 1
                row[1] = makeDate(row[1])
                if row[2] is not None:
                    row[2] = makeDate(row[2])
                row[6] = makeDate(row[6])
        except KeyError:
            self.recurringID = 0
            self.recurring = []
        # rentPaidTo: [int(id), dateObj Start, dateObj End or None, company]
        try:
            self.rentID = 0
            self.rentPaidTo = d['rent paid to']
            for row in self.rentPaidTo:
                if row[0] >= self.rentID:
                    self.rentID = row[0] + 1
                if row[1] is not None:
                    row[1] = makeDate(row[1])
                if row[2] is not None:
                    row[2] = makeDate(row[2])
        except KeyError:
            self.rentID = 0
            self.rentPaidTo = []
        self.payRates.sort(key=itemgetter(1), reverse=True)
        self.recurring.sort(key=itemgetter(1), reverse=True)
        if len(self.rentPaidTo) == 0:
            self.rentPaidTo.append([0, None, None, None])
            self.rentID = 1
        for recurringLoan in self.recurring:
            self.processRecurring(recurringLoan)

    def __str__(self):
        return self.name

    def makeDict(self):
        payRates = []
        recurring = []
        rents = []
        for row in self.payRates:
            payRates.append([row[0], makeStr(row[1]), row[2]])
        for row in self.recurring:
            if row[2] is None:
                endDate = None
            else:
                endDate = makeStr(row[2])
            recurring.append([row[0], makeStr(row[1]), endDate, row[3], row[4], row[5], makeStr(row[6])])
        for row in self.rentPaidTo:
            if row[1] is None:
                startDate = None
            else:
                startDate = makeStr(row[1])
            if row[2] is None:
                endDate = None
            else:
                endDate = makeStr(row[2])
            rents.append([row[0], startDate, endDate, row[3]])
        dic = {
            'name': self.name,
            'active': self.active,
            'pay rates': payRates,
            'recurring': recurring,
            'rent paid to': rents
        }
        return dic

    def save(self):
        self.workerList.save()

    def getPayRate(self, date):
        # pay rates organized most recent to earliest, first one found earlier is correct
        for tup in self.payRates:
            if tup[1] <= date:
                return tup[2]
        print('No pay data for ' + self.name + ' on ' + makeStr(date))

    def getLoanData(self):
        rows = loansDB.search(name=self.name)
        data = []
        for row in rows:
            date = makeDate(row[1])
            amt = float(row[3])
            data.append([row[0], date, amt, row[4], row[5]])
        data.sort(key=itemgetter(1), reverse=True)
        # data: [id, date, amt, subject, memo]
        return data

    def getHoursData(self):
        rows = hoursDB.search(name=self.name)
        data = []
        for row in rows:
            date = makeDate(row[1])
            hoursData = hoursDataFromString(row[3], row[4])
            data.append([row[0], date, row[2], hoursData, row[5]])
        # data: [id, date, workerName, [company, hours], paidBy]
        data.sort(key=itemgetter(1), reverse=True)
        return data

    def updateName(self, newName):
        loansDB.changeValue(newName, 'name', self.name)
        hoursDB.changeValue(newName, 'name', self.name)
        self.name = newName
        self.workerList.save()

    def addHours(self, date, hoursData, paidBy):
        date = makeStr(date)
        companies, hours = hoursDataToString(hoursData)
        hoursDB.insert(date, self.name, companies, hours, paidBy)

    def updateHours(self, ident, date, hoursData, paidBy):
        date = makeStr(date)
        companies, hours = hoursDataToString(hoursData)
        hoursDB.update(ident, date=date, name=self.name, company=companies, hours=hours, paidBy=paidBy)

    @staticmethod
    def deleteHours(ident):
        hoursDB.delete(ident)

    def deactivate(self):
        self.active = False
        self.workerList.save()

    def activate(self):
        self.active = True
        self.workerList.save()

    def addPayRate(self, date, payRate):
        self.payRates.append([self.payID, date, payRate])
        self.payID += 1
        self.payRates.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def getAllPayRates(self):
        return self.payRates

    def editPayRate(self, ident, date, rate):
        for row in self.payRates:
            if row[0] == ident:
                row[1] = date
                row[2] = rate
        self.payRates.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def deletePayData(self, ident):
        self.payRates = [tup for tup in self.payRates if tup[0] != ident]
        self.payRates.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def getRecurring(self):
        return self.recurring

    def addLoan(self, date, amount, subject, memo):
        date = makeStr(date)
        loansDB.insert(date, self.name, amount, subject, memo)

    def addPayment(self, date, amount, subject, memo):
        date = makeStr(date)
        loansDB.insert(date, self.name, amount, subject, memo)

    @staticmethod
    def editLoan(ident, date, amount, subject, memo):
        date = makeStr(date)
        loansDB.update(ident, date=date, amount=amount, subject=subject, memo=memo)

    @staticmethod
    def deleteLoan(ident):
        loansDB.delete(ident)

    def searchLoanData(self, subject):
        rows = loansDB.search(name=self.name)
        data = []
        for row in rows:
            date = makeDate(row[1])
            amt = float(row[3])
            if row[4] == subject:
                data.append([row[0], date, amt, row[4], row[5]])
        data.sort(key=itemgetter(1), reverse=True)
        return data

    def updateRecurring(self, ident, startDate, endDate, memo, amount, freq):
        for i in range(len(self.recurring)):
            if self.recurring[i][0] == ident:
                recurringIndex = i
                oldStartDate, oldEndDate, oldMemo, oldAmount, oldFreq, lastUpdate = self.recurring[i][1:]
                self.recurring[i] = [ident, startDate, endDate, memo, amount, freq, lastUpdate]
        loanRows = loansDB.search(name=self.name)
        recurringLoanData = []
        for row in loanRows:
            if row[3] == oldAmount and row[4] == oldMemo and row[5] == 'Recurring':
                date = makeDate(row[1])
                recurringLoanData.append([row[0], date, oldAmount, oldMemo, 'Recurring'])
        if len(recurringLoanData) == 0:
            self.workerList.save()
            return
        if freq == oldFreq and startDate == oldStartDate:
            for row in recurringLoanData:
                row[2] = amount
                row[3] = memo
                self.editLoan(row[0], row[1], row[2], row[3], row[4])
        else:
            for row in recurringLoanData:
                self.deleteLoan(row[0])
            self.recurring[recurringIndex][6] = None
            self.processRecurring(self.recurring[recurringIndex])
        self.workerList.save()

    def addRecurring(self, startDate, endDate, memo, amount, freq):
        recurringLoan = [self.recurringID, startDate, endDate, memo, amount, freq, None]
        self.recurringID += 1
        self.recurring.append(recurringLoan)
        self.processRecurring(recurringLoan)

    def processRecurring(self, recurringLoan):
        startDate, endDate, memo, amount, freq, lastUpdate = recurringLoan[1:]
        today = Date.today()
        if lastUpdate is None:
            self.addLoan(startDate, amount, memo, 'Recurring')
            currDate = startDate
            lastUpdate = currDate
        else:
            currDate = lastUpdate
        currDate = dateIncrement(currDate, freq)
        while currDate < today:
            if endDate is not None:
                if endDate <= currDate:
                    break
            self.addLoan(currDate, amount, memo, 'Recurring')
            lastUpdate = currDate
            currDate = dateIncrement(currDate, freq)
        recurringLoan[-1] = lastUpdate
        self.workerList.save()

    def getRentOverview(self):
        return sorted([row for row in self.rentPaidTo], key=itemgetter(1), reverse=True)

    def addRent(self, startDate, endDate, company):
        self.rentPaidTo.append([self.rentID, startDate, endDate, company])
        self.rentID += 1
        self.rentPaidTo.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def updateRent(self, ident, startDate, endDate, company):
        for i in range(len(self.rentPaidTo)):
            if self.rentPaidTo[i][0] == ident:
                self.rentPaidTo[i] = [ident, startDate, endDate, company]
        self.rentPaidTo.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def deleteRentData(self, ident):
        self.rentPaidTo = [tup for tup in self.rentPaidTo if tup[0] != ident]
        self.rentPaidTo.sort(key=itemgetter(1), reverse=True)
        self.workerList.save()

    def changeRentCompanyName(self, oldName, newName):
        for row in self.rentPaidTo:
            if row[3] == oldName:
                row[3] = newName
        self.workerList.save()

    def getRentsOwed(self, companyName):
        rentPeriods = []
        for row in self.rentPaidTo:
            if row[3] == companyName:
                rentPeriods.append(row)
        if len(rentPeriods) == 0:
            return
        paidBy = {}
        for row in self.getHoursData():
            if row[1] in paidBy:
                continue
            else:
                paidBy[row[1]] = row[4]
        allRents = [row for row in self.getLoanData() if row[2] < 0 and (row[3] == 'Rent' or row[3] == 'rent')]
        # rentTotals[paidBy] = -amt
        rentTotals = {}
        for rentRow in rentPeriods:
            if rentRow[2] is None:
                for row in allRents:
                    if rentRow[1] <= row[1]:
                        if row[1] in paidBy:
                            if paidBy[row[1]] in rentTotals:
                                rentTotals[paidBy[row[1]]] += row[2]
                            else:
                                rentTotals[paidBy[row[1]]] = row[2]
                        else:
                            if 'Dave' in rentTotals:
                                rentTotals['Dave'] += row[2]
                            else:
                                rentTotals['Dave'] = row[2]
            else:
                for row in allRents:
                    if rentRow[1] <= row[1] < rentRow[2]:
                        if row[1] in paidBy:
                            if paidBy[row[1]] in rentTotals:
                                rentTotals[paidBy[row[1]]] += row[2]
                            else:
                                rentTotals[paidBy[row[1]]] = row[2]
                        else:
                            if 'Dave' in rentTotals:
                                rentTotals['Dave'] += row[2]
                            else:
                                rentTotals['Dave'] = row[2]
        if len(rentTotals) == 0:
            return
        return rentTotals

    def getMiscalculation(self, date):
        loanData = self.getLoanData()
        miscalculations = ['Overpayment', 'Underpayment', 'Miscalculation']
        for row in loanData:
            if row[1] == date and row[3] in miscalculations:
                return row[2]
        return None


class WorkerList:

    def __init__(self, workersFileName):
        self.workersFileName = workersFileName
        self.workerList = []
        with open(resource_path(workersFileName)) as file:
            d = json.load(file)
        for dic in d.values():
            self.workerList.append(Worker(dic, self))

    def __getitem__(self, name):
        if type(name) != str:
            raise TypeError
        for worker in self.workerList:
            if worker.name == name:
                return worker
        raise KeyError(f'No worker with the name {name}')

    def save(self):
        with open(resource_path(self.workersFileName), 'w') as file:
            d = {}
            for worker in self.workerList:
                d[worker.name] = worker.makeDict()
            json.dump(d, file, indent=4)

    def addWorker(self, worker):
        if type(worker) == str:
            d = {
                'name': worker,
                'active': True,
                'pay rates': []
            }
            workerObj = Worker(d, self)
            self.workerList.append(workerObj)
            self.save()
            return workerObj
        else:
            self.workerList.append(worker)
            self.save()

    def getAllHoursData(self):
        rows = []
        for worker in self.workerList:
            data = worker.getHoursData()
            if data is None:
                continue
            for row in data:
                rows.append(row)
        rows.sort(key=itemgetter(1), reverse=True)
        return rows

    def getNames(self, inactive=False):
        if inactive:
            names = [worker.name for worker in self.workerList]
        else:
            names = [worker.name for worker in self.workerList if worker.active]
        return names

    def getInactiveNames(self):
        names = [worker.name for worker in self.workerList if not worker.active]
        return names

    def changeRentCompanyName(self, oldName, newName):
        for worker in self.workerList:
            worker.changeRentCompanyName(oldName, newName)

    def getWorkers(self):
        return self.workerList


class Company:

    def __init__(self, d, companyList):
        self.name = d['name']
        self.active = d['active']
        self.clearsThrough = d['clears through']
        self.companyList = companyList

    def makeDict(self):
        d = {
            'name': self.name,
            'active': self.active,
            'clears through': self.clearsThrough
        }
        return d

    def deactivate(self):
        self.active = False
        self.companyList.save()

    def activate(self):
        self.active = True
        self.companyList.save()

    def setClearsThrough(self, clearsThrough):
        self.clearsThrough = clearsThrough
        self.companyList.save()

    def getHoursData(self):
        allData = hoursDB.getAll()
        # data: [companyName, date, worker, hours, amount, paidBy]
        data = []
        for row in allData:
            if self.name in row[3]:
                pos = row[3].split('/').index(self.name)
                hrs = float(row[4].split('/')[pos])
                worker = workList[row[2]]
                date = makeDate(row[1])
                rate = worker.getPayRate(date)
                if rate is None:
                    rate = 0
                data.append([self.name, date, row[2], hrs, hrs * rate, row[5]])
        data.sort(key=itemgetter(1))
        return data

    def getDistributionData(self):
        # data: [id, date, donor, recipient, amt]
        rows = distDB.search(donor=self.name, recipient=self.name)
        data = []
        for row in rows:
            date = makeDate(row[1])
            amt = float(row[4])
            data.append([row[0], date, row[2], row[3], amt])
        data.sort(key=itemgetter(1))
        return data

    def addDistribution(self, date, donor, recipient, amt):
        date = makeStr(date)
        if donor is None:
            donor = self.name
        if recipient is None:
            recipient = self.name
        distDB.insert(date, donor, recipient, amt)

    @staticmethod
    def editDistribution(ident, date, donor, recipient, amt):
        date = makeStr(date)
        distDB.update(ident, date=date, donor=donor, recipient=recipient, amount=amt)

    @staticmethod
    def deleteDistribution(ident):
        distDB.delete(ident)

    def updateName(self, newName):
        distDB.changeValue(newName, 'donor', self.name)
        distDB.changeValue(newName, 'recipient', self.name)
        hoursDB.replaceFrom(newName, 'company', self.name)
        workList.changeRentCompanyName(self.name, newName)
        self.name = newName
        self.companyList.save()


class CompanyList:

    def __init__(self, companyFile):
        self.fileName = companyFile
        self.companyList = []
        with open(resource_path(companyFile)) as file:
            d = json.load(file)
        for dic in d.values():
            self.companyList.append(Company(dic, self))

    def save(self):
        d = {}
        for company in self.companyList:
            d[company.name] = company.makeDict()
        with open(resource_path(self.fileName), 'w') as file:
            json.dump(d, file, indent=4)

    def addCompany(self, company):
        if type(company) is str:
            d = {
                'name': company,
                'active': True,
                'clears through': None
            }
            companyObj = Company(d, self)
            self.companyList.append(companyObj)
            return companyObj
        else:
            self.companyList.append(company)
        self.save()

    def __getitem__(self, companyName):
        for company in self.companyList:
            if company.name == companyName:
                return company

    def getNames(self):
        names = []
        for company in self.companyList:
            names.append(company.name)
        return names

    def getAllCompanies(self):
        return self.companyList

    def getActiveNames(self):
        return [company.name for company in self.companyList if company.active]

    def getInactiveNames(self):
        return [company.name for company in self.companyList if not company.active]

    def getClearsThrough(self, companyName):
        clearing = []
        for company in self.companyList:
            if company.clearsThrough == companyName:
                clearing.append(company)
        return clearing


compList = CompanyList('companies.json')
workList = WorkerList('workers.json')

mainApp = MainApplication(window, workList, compList, width=WINDOW_WIDTH, height=WINDOW_HEIGHT)
window.mainloop()
