import os
import pandas as pd
import numpy as np
import string
import itertools

#set global parameters
path = '/Users/xieminghua/PycharmProjects/AutomatedReport'
in_file = 'test.xlsx'
sheet = 'Master '
status = ['Active']
month_start_col = 'CQ'
yr_start_col = 'AI'
days_in_month = 31
num_of_working_days_current_month = 22
offices = ['Head Office (London)',
           'London office',
           'Mancherster office',
           'Edinburgh Office',
           'US London']
report_items = ['Out of office work',
                'work in China',
                'work in US',
                'Late',
                'Working from home',
                '1/2 working from home',
                'Sick leave',
                '1/2 sick leave',
                'Personal reason',
                'Bank Holiday',
                'Weekend',
                'Unpaid leave',
                'Annual leave',
                'Absent',
                'TOIL',
                'Unpaid leave',
                'no log in',
                '1/2 Annual leave']

#import excel file
class GenerateReport:
    def __init__(self, path, in_file, sheet, month_start_col, days_in_month, yr_start_col, status, offices, report_items):
        self.path = path
        self.in_file = in_file
        self.sheet = sheet
        self.month_start_col = month_start_col
        self.days_in_month = days_in_month
        self.yr_start_col = yr_start_col
        self.status = status
        self.offices = offices
        self.report_items = report_items

    def inputFile(self):
        global start_col
        global year_start_col
        global df2
        global df1
        def excel_cols():
            n = 1
            while True:
                yield from (''.join(group) for group in itertools.product(string.ascii_uppercase, repeat=n))
                n += 1
        col_letter = list(itertools.islice(excel_cols(), 16383))
        col_index = list(range(0, 16384))
        key = {}
        for i in col_letter:
            for k in col_index:
                key[i] = k
                col_index.remove(k)
                break

        start_col = key[self.month_start_col]
        year_start_col = key[self.yr_start_col]
        os.chdir(self.path)
        end_col = start_col + self.days_in_month
        df1 = pd.read_excel(self.in_file, sheet_name=self.sheet)

        col_range = list(range(start_col, end_col))
        col_range.insert(0, 0)
        col_range.insert(1, 1)
        col_range.insert(2, 2)
        col_range.insert(3, 3)
        # all rows in the report month
        df2 = df1.iloc[:, col_range]

    def activeUsersonly(self):
        # filter for all active users in the report month
        global df3
        df3 = df2[df2.Status.isin(self.status)]
        df3 = df3.astype(np.object)

    def currentMonth(self):
        # generate report for each office
        global df_all_current_month
        global df4
        global d
        global df_all_current_month
        d = {}
        for k in range(len(self.offices)):
            df_temp = df3[df3['Office'] == self.offices[k]]
            df_temp = df_temp.astype(np.object)
            dict = {}
            col_list = []
            staff_list = []
            count_list = [[] for i in range(len(self.report_items))]
            counter = 1

            for i in self.report_items:
                count = counter - 1
                col_list.append(i)
                for j in range(len(df_temp)):
                    df4 = df_temp.iloc[[j]]
                    staff_name = df4.iloc[0, 0]
                    num = df4[df4 == i].count().sum()
                    count_list[count].append(num)
                    print('{} took {} days of {} in March'.format(staff_name, num, i))
                counter += 1
                dict[i] = count_list[count]
            d[k] = pd.DataFrame.from_dict(dict)
            staff_list = df_temp['Name'].values.tolist()
            staff_list_series = pd.Series(staff_list)
            d[k].insert(loc=0, column='Name', value=staff_list_series)
            d[k].insert(loc=1, column='Office', value=offices[k])
            days_due = df3['Days due'].values.tolist()
            days_due_series = pd.Series(days_due)
            d[k].insert(loc=2, column='Days due', value=days_due_series)

        # merge all dataframes into one
        df_all_current_month = pd.DataFrame()
        for i in range(len(d)):
            df_all_current_month = df_all_current_month.append(d[i])
            print('Appending dataframe{}'.format(i))

        # calculate total annual leave, work from home, sick leave
        df_all_current_month['Total annual leave taken in current month'] = df_all_current_month[
                                                                                '1/2 Annual leave'] * 0.5 + \
                                                                            df_all_current_month['Annual leave']
        df_all_current_month['Total working from home'] = df_all_current_month['1/2 working from home'] * 0.5 + \
                                                          df_all_current_month['Working from home']
        df_all_current_month['Total sick leave'] = df_all_current_month['1/2 sick leave'] * 0.5 + df_all_current_month[
            'Sick leave']

    def yeartoMonth(self):
        global df_all_current_month
        # calculate the total annual leave taken to month
        year_end_col = start_col - 1
        ytm_col_range = list(range(year_start_col, year_end_col))
        ytm_col_range.insert(0, 0)
        ytm_col_range.insert(1, 1)
        ytm_col_range.insert(2, 2)
        ytm_col_range.insert(3, 3)

        df2_ytm = df1.iloc[:, ytm_col_range]

        # filter for all active users in the report month
        df3_ytm = df2_ytm[df2_ytm.Status.isin(self.status)]
        df3_ytm = df3_ytm.astype(np.object)

        d_ytm = {}
        for k in range(len(self.offices)):
            df_temp_ytm = df3_ytm[df3_ytm['Office'] == self.offices[k]]
            df_temp_ytm = df_temp_ytm.astype(np.object)
            dict = {}
            col_list = []
            staff_list = []
            count_list = [[] for i in range(len(self.report_items))]
            counter = 1

            for i in self.report_items:
                count = counter - 1
                col_list.append(i)
                for j in range(len(df_temp_ytm)):
                    df4_ytm = df_temp_ytm.iloc[[j]]
                    staff_name = df4_ytm.iloc[0, 0]
                    num = df4_ytm[df4_ytm == i].count().sum()
                    count_list[count].append(num)
                    print('{} took {} days of {} in March'.format(staff_name, num, i))
                counter += 1
                dict[i] = count_list[count]
            d_ytm[k] = pd.DataFrame.from_dict(dict)
            staff_list_ytm = df_temp_ytm['Name'].values.tolist()
            staff_list_series_ytm = pd.Series(staff_list_ytm)
            d_ytm[k].insert(loc=0, column='Name', value=staff_list_series_ytm)
            d_ytm[k].insert(loc=1, column='Office', value=offices[k])
            days_due_ytm = df3_ytm['Days due'].values.tolist()
            days_due_series_ytm = pd.Series(days_due_ytm)
            d_ytm[k].insert(loc=2, column='Days due', value=days_due_series_ytm)

        # merge all dataframes into one
        df_all_ytm = pd.DataFrame()
        for i in range(len(d_ytm)):
            df_all_ytm = df_all_ytm.append(d_ytm[i])
            print('Appending ytm dataframe {}'.format(i))

        df_all_ytm['Total annual leave taken to previous month'] = df_all_ytm['1/2 Annual leave'] * 0.5 + df_all_ytm[
            'Annual leave']

        annual_leave_ytm = df_all_ytm['Total annual leave taken to previous month'].values.tolist()
        annual_leave_ytm_series = pd.Series(annual_leave_ytm)

        df_all_current_month.insert(loc=3,
                                    column='Annual leave taken until previous month end',
                                    value=annual_leave_ytm_series)

        df_all_current_month['Total annual leave remaining'] = df_all_current_month['Days due'] - df_all_current_month[
            'Annual leave taken until previous month end'] - df_all_current_month[
                                                                   'Total annual leave taken in current month']
        df_all_current_month['% of working from home'] = df_all_current_month[
                                                             'Total working from home'] / num_of_working_days_current_month * 100
        df_all_current_month['% of sick leave'] = df_all_current_month[
                                                      'Total sick leave'] / num_of_working_days_current_month * 100
        df_all_current_month['% of late'] = df_all_current_month['Late'] / num_of_working_days_current_month * 100
        df_all_current_month['% of no log in'] = df_all_current_month[
                                                     'no log in'] / num_of_working_days_current_month * 100

        # list(df_all_current_month.columns.values)

        df_all_current_month = df_all_current_month[['Name',
                                                     'Office',
                                                     'Days due',
                                                     'Total annual leave remaining',
                                                     'Annual leave taken until previous month end',
                                                     'Total annual leave taken in current month',
                                                     'Out of office work', 'work in China', 'work in US',
                                                     'Late', '% of late', 'Working from home', '1/2 working from home',
                                                     'Total working from home',
                                                     '% of working from home', 'Sick leave', '1/2 sick leave',
                                                     'Total sick leave', '% of sick leave', 'Personal reason',
                                                     'Bank Holiday',
                                                     'Weekend', 'Unpaid leave', 'Absent', 'TOIL', 'no log in',
                                                     '% of no log in',
                                                     '1/2 Annual leave', 'Annual leave']]

        print(df_all_current_month.to_string())

    def exportResults(self):
        df_all_current_month.to_excel('Results.xlsx', sheet_name='Current Month')
        print('Report generation complete!')

#export dataframe as a test
#df_all_current_month
#df_all_current_month.to_excel('Results.xlsx', sheet_name='Head Office')

#export all dataframes as separate files
#for i in range(len(d)):
#    d[i].to_excel('{}_results.xlsx'.format(offices[i]), sheet_name='Results')

#df_all_current_month.columns.tolist()

initiate = GenerateReport(path, in_file, sheet, month_start_col, days_in_month, yr_start_col, status, offices, report_items)

initiate.inputFile()
initiate.activeUsersonly()
initiate.currentMonth()
initiate.yeartoMonth()
initiate.exportResults()
