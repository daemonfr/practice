import os
import pandas as pd
import numpy as np

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
        key = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12,
         'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24,
         'Z': 25, 'AA': 26, 'AB': 27, 'AC': 28, 'AD': 29, 'AE': 30, 'AF': 31, 'AG': 32, 'AH': 33, 'AI': 34, 'AJ': 35,
         'AK': 36, 'AL': 37, 'AM': 38, 'AN': 39, 'AO': 40, 'AP': 41, 'AQ': 42, 'AR': 43, 'AS': 44, 'AT': 45, 'AU': 46,
         'AV': 47, 'AW': 48, 'AX': 49, 'AY': 50, 'AZ': 51, 'BA': 52, 'BB': 53, 'BC': 54, 'BD': 55, 'BE': 56, 'BF': 57,
         'BG': 58, 'BH': 59, 'BI': 60, 'BJ': 61, 'BK': 62, 'BL': 63, 'BM': 64, 'BN': 65, 'BO': 66, 'BP': 67, 'BQ': 68,
         'BR': 69, 'BS': 70, 'BT': 71, 'BU': 72, 'BV': 73, 'BW': 74, 'BX': 75, 'BY': 76, 'BZ': 77, 'CA': 78, 'CB': 79,
         'CC': 80, 'CD': 81, 'CE': 82, 'CF': 83, 'CG': 84, 'CH': 85, 'CI': 86, 'CJ': 87, 'CK': 88, 'CL': 89, 'CM': 90,
         'CN': 91, 'CO': 92, 'CP': 93, 'CQ': 94, 'CR': 95, 'CS': 96, 'CT': 97, 'CU': 98, 'CV': 99, 'CW': 100, 'CX': 101,
         'CY': 102, 'CZ': 103, 'DA': 104, 'DB': 105, 'DC': 106, 'DD': 107, 'DE': 108, 'DF': 109, 'DG': 110, 'DH': 111,
         'DI': 112, 'DJ': 113, 'DK': 114, 'DL': 115, 'DM': 116, 'DN': 117, 'DO': 118, 'DP': 119, 'DQ': 120, 'DR': 121,
         'DS': 122, 'DT': 123, 'DU': 124, 'DV': 125, 'DW': 126, 'DX': 127, 'DY': 128, 'DZ': 129, 'EA': 130, 'EB': 131,
         'EC': 132, 'ED': 133, 'EE': 134, 'EF': 135, 'EG': 136, 'EH': 137, 'EI': 138, 'EJ': 139, 'EK': 140, 'EL': 141,
         'EM': 142, 'EN': 143, 'EO': 144, 'EP': 145, 'EQ': 146, 'ER': 147, 'ES': 148, 'ET': 149, 'EU': 150, 'EV': 151,
         'EW': 152, 'EX': 153, 'EY': 154, 'EZ': 155, 'FA': 156, 'FB': 157, 'FC': 158, 'FD': 159, 'FE': 160, 'FF': 161,
         'FG': 162, 'FH': 163, 'FI': 164, 'FJ': 165, 'FK': 166, 'FL': 167, 'FM': 168, 'FN': 169, 'FO': 170, 'FP': 171,
         'FQ': 172, 'FR': 173, 'FS': 174, 'FT': 175, 'FU': 176, 'FV': 177, 'FW': 178, 'FX': 179, 'FY': 180, 'FZ': 181,
         'GA': 182, 'GB': 183, 'GC': 184, 'GD': 185, 'GE': 186, 'GF': 187, 'GG': 188, 'GH': 189, 'GI': 190, 'GJ': 191,
         'GK': 192, 'GL': 193, 'GM': 194, 'GN': 195, 'GO': 196, 'GP': 197, 'GQ': 198, 'GR': 199, 'GS': 200, 'GT': 201,
         'GU': 202, 'GV': 203, 'GW': 204, 'GX': 205, 'GY': 206, 'GZ': 207, 'HA': 208, 'HB': 209, 'HC': 210, 'HD': 211,
         'HE': 212, 'HF': 213, 'HG': 214, 'HH': 215, 'HI': 216, 'HJ': 217, 'HK': 218, 'HL': 219, 'HM': 220, 'HN': 221,
         'HO': 222, 'HP': 223, 'HQ': 224, 'HR': 225, 'HS': 226, 'HT': 227, 'HU': 228, 'HV': 229, 'HW': 230, 'HX': 231,
         'HY': 232, 'HZ': 233, 'IA': 234, 'IB': 235, 'IC': 236, 'ID': 237, 'IE': 238, 'IF': 239, 'IG': 240, 'IH': 241,
         'II': 242, 'IJ': 243, 'IK': 244, 'IL': 245, 'IM': 246, 'IN': 247, 'IO': 248, 'IP': 249, 'IQ': 250, 'IR': 251,
         'IS': 252, 'IT': 253, 'IU': 254, 'IV': 255, 'IW': 256, 'IX': 257, 'IY': 258, 'IZ': 259, 'JA': 260, 'JB': 261,
         'JC': 262, 'JD': 263, 'JE': 264, 'JF': 265, 'JG': 266, 'JH': 267, 'JI': 268, 'JJ': 269, 'JK': 270, 'JL': 271,
         'JM': 272, 'JN': 273, 'JO': 274, 'JP': 275, 'JQ': 276, 'JR': 277, 'JS': 278, 'JT': 279, 'JU': 280, 'JV': 281,
         'JW': 282, 'JX': 283, 'JY': 284, 'JZ': 285, 'KA': 286, 'KB': 287, 'KC': 288, 'KD': 289, 'KE': 290, 'KF': 291,
         'KG': 292, 'KH': 293, 'KI': 294, 'KJ': 295, 'KK': 296, 'KL': 297, 'KM': 298, 'KN': 299, 'KO': 300, 'KP': 301,
         'KQ': 302, 'KR': 303, 'KS': 304, 'KT': 305, 'KU': 306, 'KV': 307, 'KW': 308, 'KX': 309, 'KY': 310, 'KZ': 311,
         'LA': 312, 'LB': 313, 'LC': 314, 'LD': 315, 'LE': 316, 'LF': 317, 'LG': 318, 'LH': 319, 'LI': 320, 'LJ': 321,
         'LK': 322, 'LL': 323, 'LM': 324, 'LN': 325, 'LO': 326, 'LP': 327, 'LQ': 328, 'LR': 329, 'LS': 330, 'LT': 331,
         'LU': 332, 'LV': 333, 'LW': 334, 'LX': 335, 'LY': 336, 'LZ': 337, 'MA': 338, 'MB': 339, 'MC': 340, 'MD': 341,
         'ME': 342, 'MF': 343, 'MG': 344, 'MH': 345, 'MI': 346, 'MJ': 347, 'MK': 348, 'ML': 349, 'MM': 350, 'MN': 351,
         'MO': 352, 'MP': 353, 'MQ': 354, 'MR': 355, 'MS': 356, 'MT': 357, 'MU': 358, 'MV': 359, 'MW': 360, 'MX': 361,
         'MY': 362, 'MZ': 363, 'NA': 364, 'NB': 365, 'NC': 366, 'ND': 367, 'NE': 368, 'NF': 369, 'NG': 370, 'NH': 371,
         'NI': 372, 'NJ': 373, 'NK': 374, 'NL': 375, 'NM': 376, 'NN': 377, 'NO': 378, 'NP': 379, 'NQ': 380, 'NR': 381,
         'NS': 382, 'NT': 383, 'NU': 384, 'NV': 385, 'NW': 386, 'NX': 387, 'NY': 388, 'NZ': 389, 'OA': 390, 'OB': 391,
         'OC': 392, 'OD': 393, 'OE': 394, 'OF': 395, 'OG': 396, 'OH': 397, 'OI': 398, 'OJ': 399, 'OK': 400, 'OL': 401,
         'OM': 402, 'ON': 403, 'OO': 404, 'OP': 405, 'OQ': 406, 'OR': 407, 'OS': 408, 'OT': 409, 'OU': 410, 'OV': 411,
         'OW': 412, 'OX': 413, 'OY': 414, 'OZ': 415, 'PA': 416, 'PB': 417, 'PC': 418, 'PD': 419, 'PE': 420, 'PF': 421,
         'PG': 422, 'PH': 423, 'PI': 424, 'PJ': 425, 'PK': 426, 'PL': 427, 'PM': 428, 'PN': 429, 'PO': 430, 'PP': 431,
         'PQ': 432, 'PR': 433, 'PS': 434, 'PT': 435, 'PU': 436, 'PV': 437, 'PW': 438, 'PX': 439, 'PY': 440, 'PZ': 441,
         'QA': 442, 'QB': 443, 'QC': 444, 'QD': 445, 'QE': 446, 'QF': 447, 'QG': 448, 'QH': 449, 'QI': 450, 'QJ': 451,
         'QK': 452, 'QL': 453, 'QM': 454, 'QN': 455, 'QO': 456, 'QP': 457, 'QQ': 458, 'QR': 459, 'QS': 460, 'QT': 461,
         'QU': 462, 'QV': 463, 'QW': 464, 'QX': 465, 'QY': 466, 'QZ': 467, 'RA': 468, 'RB': 469, 'RC': 470, 'RD': 471,
         'RE': 472, 'RF': 473, 'RG': 474, 'RH': 475, 'RI': 476, 'RJ': 477, 'RK': 478, 'RL': 479, 'RM': 480, 'RN': 481,
         'RO': 482, 'RP': 483, 'RQ': 484, 'RR': 485, 'RS': 486, 'RT': 487, 'RU': 488, 'RV': 489, 'RW': 490, 'RX': 491,
         'RY': 492, 'RZ': 493, 'SA': 494, 'SB': 495, 'SC': 496, 'SD': 497, 'SE': 498, 'SF': 499, 'SG': 500}
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
            days_due = df_temp['Days due'].values.tolist()
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
                                                             'Total working from home'] / num_of_working_days_current_month
        df_all_current_month['% of sick leave'] = df_all_current_month[
                                                      'Total sick leave'] / num_of_working_days_current_month
        df_all_current_month['% of late'] = df_all_current_month['Late'] / num_of_working_days_current_month
        df_all_current_month['% of no log in'] = df_all_current_month[
                                                     'no log in'] / num_of_working_days_current_month

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
