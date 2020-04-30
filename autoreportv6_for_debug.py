import os
import pandas as pd
import numpy as np
import datetime

#set global parameters
path = 'C:/Test'
in_file = '2020 Staff holiday record - live-v2.xlsx'
sheet = '2020 individual attendance'
status = ['Active']
month_start_col = 'DQ'
yr_start_col = 'AD'
days_in_month = 30
num_of_working_days_current_month = 22
num_of_working_days_current_year = 254
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
num_of_days_each_month = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30]

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
        global end_col
        global year_start_col
        global df2
        global df1
        global key
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
        end_col = int(start_col) + self.days_in_month
        df1 = pd.read_excel(self.in_file, sheet_name=self.sheet)

        col_range = list(range(start_col, end_col))
        col_range.insert(0, 0)
        col_range.insert(1, 1)
        col_range.insert(2, 2)
        col_range.insert(3, 3)
        # all rows in the report month
        df2 = df1.iloc[:, col_range]

    def activeUsersonly(self):
        # filter for all active users in the report month. df3 is for current month
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
        df_all_current_month['Total working from home in current month'] = df_all_current_month['1/2 working from home'] * 0.5 + \
                                                          df_all_current_month['Working from home']
        df_all_current_month['Total sick leave in current month'] = df_all_current_month['1/2 sick leave'] * 0.5 + df_all_current_month[
            'Sick leave']

    def yeartoMonth(self):
        global df_all_current_month
        global df_all_ytm
        global df2_full_yr_annual_leave
        # generate dataframe for the total annual leave taken for the full year
        full_yr_annual_leave_range = list(range(year_start_col, year_start_col + 366))
        full_yr_annual_leave_range.insert(0, 0)
        full_yr_annual_leave_range.insert(1, 1)
        full_yr_annual_leave_range.insert(2, 2)
        full_yr_annual_leave_range.insert(3, 3)
        
        df2_full_yr_annual_leave = df1.iloc[:, full_yr_annual_leave_range]
        
        df3_full_yr_annual_leave = df2_full_yr_annual_leave[df2_full_yr_annual_leave.Status.isin(self.status)]
        df3_full_yr_annual_leave = df3_full_yr_annual_leave.astype(np.object)
              
        # calculate the total annual leave taken to month (up to and including current month)
        ytm_col_range = list(range(year_start_col, end_col))
        ytm_col_range.insert(0, 0)
        ytm_col_range.insert(1, 1)
        ytm_col_range.insert(2, 2)
        ytm_col_range.insert(3, 3)

        df2_ytm = df1.iloc[:, ytm_col_range]

        # filter for all active users in the report month (up to and including current month)
        df3_ytm = df2_ytm[df2_ytm.Status.isin(self.status)]
        df3_ytm = df3_ytm.astype(np.object)

        d_ytm = {}
        for k in range(len(self.offices)):
            df_temp_ytm = df3_ytm[df3_ytm['Office'] == self.offices[k]]
            df_temp_ytm = df_temp_ytm.astype(np.object)
            df_temp_full_yr_annual = df3_full_yr_annual_leave[df3_full_yr_annual_leave['Office'] == self.offices[k]]
            df_temp_full_yr_annual = df_temp_full_yr_annual.astype(np.object)
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
                    print('{} took {} days of {} year to month (up to and including current month)'.format(staff_name, num, i))
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

        #calculate annuale leave metric
        df_all_ytm['Total annual leave taken up to and including current month'] = df_all_ytm['1/2 Annual leave'] * 0.5 + df_all_ytm[
            'Annual leave']
        annual_leave_ytm = df_all_ytm['Total annual leave taken up to and including current month'].values.tolist()
        annual_leave_ytm_series = pd.Series(annual_leave_ytm)
        df_all_current_month.insert(loc=3,
                                    column='Total annual leave taken up to and including current month',
                                    value=annual_leave_ytm_series)
        df_all_current_month['Total annual leave remaining'] = df_all_current_month['Days due'] - df_all_current_month[
            'Total annual leave taken up to and including current month']
        #calculate wfh metric
        df_all_ytm['Total working from home up to and including current month'] = df_all_ytm['1/2 working from home'] * 0.5 + df_all_ytm['Working from home']
        wfh_ytm = df_all_ytm['Total working from home up to and including current month'].values.tolist()
        wfh_ytm_series = pd.Series(wfh_ytm)
        df_all_current_month.insert(loc=4,
                                    column='Total working from home taken up to and including current month',
                                    value=wfh_ytm_series)
        df_all_current_month['Current month % of working from home'] = df_all_current_month[
                                                             'Total working from home in current month'] / num_of_working_days_current_month
        df_all_current_month['Year to month % of working from home'] = (df_all_current_month[
                                                                            'Total working from home taken up to and including current month']) / num_of_working_days_current_year
        # calculate sick leave metric
        df_all_ytm['Total sick leave up to and including current month'] = df_all_ytm['1/2 sick leave'] * 0.5 + \
                                                                  df_all_ytm['Sick leave']
        sick_leave_ytm = df_all_ytm['Total sick leave up to and including current month'].values.tolist()
        sick_leave_ytm_series = pd.Series(sick_leave_ytm)
        df_all_current_month.insert(loc=5,
                                    column='Total sick leave up to and including current month',
                                    value=sick_leave_ytm_series)
        df_all_current_month['Current month % of sick leave'] = df_all_current_month[
                                                                           'Total sick leave in current month'] / num_of_working_days_current_month
        df_all_current_month['Year to month % of sick leave'] = (df_all_current_month[
                                                                            'Total sick leave up to and including current month']) / num_of_working_days_current_year

        # calculate late metric
        late_ytm = df_all_ytm['Late'].values.tolist()
        late_ytm_series = pd.Series(late_ytm)
        df_all_current_month.insert(loc=6,
                                    column='Total late days up to and including current month',
                                    value=late_ytm_series)
        df_all_current_month['Current month % of late'] = df_all_current_month[
                                                              'Late'] / num_of_working_days_current_month

        df_all_current_month['Year to month % of late'] = (df_all_ytm['Late']) / num_of_working_days_current_year

        # calculate no log in metric
        no_log_in_ytm = df_all_ytm['no log in'].values.tolist()
        no_log_in_ytm_series = pd.Series(no_log_in_ytm)
        df_all_current_month.insert(loc=7,
                                    column='No log in days up to and including current month',
                                    value=no_log_in_ytm_series)
        df_all_current_month['Current month % of no log in'] = df_all_current_month[
                                                                   'no log in'] / num_of_working_days_current_month

        df_all_current_month['Year to month % of no log in'] = (df_all_current_month['no log in']
                                                           + df_all_ytm['no log in']) / num_of_working_days_current_year

        # list(df_all_current_month.columns.values)

        df_all_current_month = df_all_current_month[['Name',
                                                     'Office',
                                                     'Days due',
                                                     'Total annual leave remaining',
                                                     'Total annual leave taken up to and including current month',
                                                     'Total annual leave taken in current month',
                                                     'Out of office work', 'work in China', 'work in US',
                                                     'Late', 'Total late days up to and including current month',
                                                     'Current month % of late', 'Year to month % of late',
                                                     'Working from home', '1/2 working from home', 'Total working from home in current month',
                                                     'Total working from home taken up to and including current month',
                                                     'Current month % of working from home',
                                                     'Year to month % of working from home',
                                                     'Sick leave', '1/2 sick leave', 'Total sick leave in current month',
                                                     'Total sick leave up to and including current month', 'Current month % of sick leave', 'Year to month % of sick leave',
                                                     'Personal reason',
                                                     'Bank Holiday',
                                                     'Weekend', 'Unpaid leave', 'Absent', 'TOIL',
                                                     'no log in', 'No log in days up to and including current month',
                                                     'Current month % of no log in', 'Year to month % of no log in',
                                                     '1/2 Annual leave', 'Annual leave']]

        print(df_all_current_month.to_string())

    def exportResults(self):
        df_all_current_month.to_excel('Results.xlsx', sheet_name='Current Month')
        print('Current month and YTM report generation complete!')

    def annualOverview(self):
        global num_of_days_each_month
        global df2
        global df_ov_all_final
        month_list = ['January', 'February',
                      'March',
                      'April',
                      'May',
                      'June',
                      'July',
                      'August',
                      'September',
                      'October',
                      'November', 'December']
        days_in_month_list = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        cum_sum = np.cumsum(num_of_days_each_month)
        d_ov = {}
        # calculate column range for each month:
        # first month is 0 because there is no need to calculate Jan start column
        counter = 1

        for i in range(len(num_of_days_each_month)):
            month_start_col_idx = key[yr_start_col] + cum_sum[i]
            rev_key = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E', 5: 'F', 6: 'G', 7: 'H', 8: 'I', 9: 'J', 10: 'K', 11: 'L',
                       12: 'M', 13: 'N', 14: 'O', 15: 'P', 16: 'Q', 17: 'R', 18: 'S', 19: 'T', 20: 'U', 21: 'V',
                       22: 'W', 23: 'X', 24: 'Y', 25: 'Z', 26: 'AA', 27: 'AB', 28: 'AC', 29: 'AD', 30: 'AE', 31: 'AF',
                       32: 'AG', 33: 'AH', 34: 'AI', 35: 'AJ', 36: 'AK', 37: 'AL', 38: 'AM', 39: 'AN', 40: 'AO',
                       41: 'AP', 42: 'AQ', 43: 'AR', 44: 'AS', 45: 'AT', 46: 'AU', 47: 'AV', 48: 'AW', 49: 'AX',
                       50: 'AY', 51: 'AZ', 52: 'BA', 53: 'BB', 54: 'BC', 55: 'BD', 56: 'BE', 57: 'BF', 58: 'BG',
                       59: 'BH', 60: 'BI', 61: 'BJ', 62: 'BK', 63: 'BL', 64: 'BM', 65: 'BN', 66: 'BO', 67: 'BP',
                       68: 'BQ', 69: 'BR', 70: 'BS', 71: 'BT', 72: 'BU', 73: 'BV', 74: 'BW', 75: 'BX', 76: 'BY',
                       77: 'BZ', 78: 'CA', 79: 'CB', 80: 'CC', 81: 'CD', 82: 'CE', 83: 'CF', 84: 'CG', 85: 'CH',
                       86: 'CI', 87: 'CJ', 88: 'CK', 89: 'CL', 90: 'CM', 91: 'CN', 92: 'CO', 93: 'CP', 94: 'CQ',
                       95: 'CR', 96: 'CS', 97: 'CT', 98: 'CU', 99: 'CV', 100: 'CW', 101: 'CX', 102: 'CY', 103: 'CZ',
                       104: 'DA', 105: 'DB', 106: 'DC', 107: 'DD', 108: 'DE', 109: 'DF', 110: 'DG', 111: 'DH',
                       112: 'DI', 113: 'DJ', 114: 'DK', 115: 'DL', 116: 'DM', 117: 'DN', 118: 'DO', 119: 'DP',
                       120: 'DQ', 121: 'DR', 122: 'DS', 123: 'DT', 124: 'DU', 125: 'DV', 126: 'DW', 127: 'DX',
                       128: 'DY', 129: 'DZ', 130: 'EA', 131: 'EB', 132: 'EC', 133: 'ED', 134: 'EE', 135: 'EF',
                       136: 'EG', 137: 'EH', 138: 'EI', 139: 'EJ', 140: 'EK', 141: 'EL', 142: 'EM', 143: 'EN',
                       144: 'EO', 145: 'EP', 146: 'EQ', 147: 'ER', 148: 'ES', 149: 'ET', 150: 'EU', 151: 'EV',
                       152: 'EW', 153: 'EX', 154: 'EY', 155: 'EZ', 156: 'FA', 157: 'FB', 158: 'FC', 159: 'FD',
                       160: 'FE', 161: 'FF', 162: 'FG', 163: 'FH', 164: 'FI', 165: 'FJ', 166: 'FK', 167: 'FL',
                       168: 'FM', 169: 'FN', 170: 'FO', 171: 'FP', 172: 'FQ', 173: 'FR', 174: 'FS', 175: 'FT',
                       176: 'FU', 177: 'FV', 178: 'FW', 179: 'FX', 180: 'FY', 181: 'FZ', 182: 'GA', 183: 'GB',
                       184: 'GC', 185: 'GD', 186: 'GE', 187: 'GF', 188: 'GG', 189: 'GH', 190: 'GI', 191: 'GJ',
                       192: 'GK', 193: 'GL', 194: 'GM', 195: 'GN', 196: 'GO', 197: 'GP', 198: 'GQ', 199: 'GR',
                       200: 'GS', 201: 'GT', 202: 'GU', 203: 'GV', 204: 'GW', 205: 'GX', 206: 'GY', 207: 'GZ',
                       208: 'HA', 209: 'HB', 210: 'HC', 211: 'HD', 212: 'HE', 213: 'HF', 214: 'HG', 215: 'HH',
                       216: 'HI', 217: 'HJ', 218: 'HK', 219: 'HL', 220: 'HM', 221: 'HN', 222: 'HO', 223: 'HP',
                       224: 'HQ', 225: 'HR', 226: 'HS', 227: 'HT', 228: 'HU', 229: 'HV', 230: 'HW', 231: 'HX',
                       232: 'HY', 233: 'HZ', 234: 'IA', 235: 'IB', 236: 'IC', 237: 'ID', 238: 'IE', 239: 'IF',
                       240: 'IG', 241: 'IH', 242: 'II', 243: 'IJ', 244: 'IK', 245: 'IL', 246: 'IM', 247: 'IN',
                       248: 'IO', 249: 'IP', 250: 'IQ', 251: 'IR', 252: 'IS', 253: 'IT', 254: 'IU', 255: 'IV',
                       256: 'IW', 257: 'IX', 258: 'IY', 259: 'IZ', 260: 'JA', 261: 'JB', 262: 'JC', 263: 'JD',
                       264: 'JE', 265: 'JF', 266: 'JG', 267: 'JH', 268: 'JI', 269: 'JJ', 270: 'JK', 271: 'JL',
                       272: 'JM', 273: 'JN', 274: 'JO', 275: 'JP', 276: 'JQ', 277: 'JR', 278: 'JS', 279: 'JT',
                       280: 'JU', 281: 'JV', 282: 'JW', 283: 'JX', 284: 'JY', 285: 'JZ', 286: 'KA', 287: 'KB',
                       288: 'KC', 289: 'KD', 290: 'KE', 291: 'KF', 292: 'KG', 293: 'KH', 294: 'KI', 295: 'KJ',
                       296: 'KK', 297: 'KL', 298: 'KM', 299: 'KN', 300: 'KO', 301: 'KP', 302: 'KQ', 303: 'KR',
                       304: 'KS', 305: 'KT', 306: 'KU', 307: 'KV', 308: 'KW', 309: 'KX', 310: 'KY', 311: 'KZ',
                       312: 'LA', 313: 'LB', 314: 'LC', 315: 'LD', 316: 'LE', 317: 'LF', 318: 'LG', 319: 'LH',
                       320: 'LI', 321: 'LJ', 322: 'LK', 323: 'LL', 324: 'LM', 325: 'LN', 326: 'LO', 327: 'LP',
                       328: 'LQ', 329: 'LR', 330: 'LS', 331: 'LT', 332: 'LU', 333: 'LV', 334: 'LW', 335: 'LX',
                       336: 'LY', 337: 'LZ', 338: 'MA', 339: 'MB', 340: 'MC', 341: 'MD', 342: 'ME', 343: 'MF',
                       344: 'MG', 345: 'MH', 346: 'MI', 347: 'MJ', 348: 'MK', 349: 'ML', 350: 'MM', 351: 'MN',
                       352: 'MO', 353: 'MP', 354: 'MQ', 355: 'MR', 356: 'MS', 357: 'MT', 358: 'MU', 359: 'MV',
                       360: 'MW', 361: 'MX', 362: 'MY', 363: 'MZ', 364: 'NA', 365: 'NB', 366: 'NC', 367: 'ND',
                       368: 'NE', 369: 'NF', 370: 'NG', 371: 'NH', 372: 'NI', 373: 'NJ', 374: 'NK', 375: 'NL',
                       376: 'NM', 377: 'NN', 378: 'NO', 379: 'NP', 380: 'NQ', 381: 'NR', 382: 'NS', 383: 'NT',
                       384: 'NU', 385: 'NV', 386: 'NW', 387: 'NX', 388: 'NY', 389: 'NZ', 390: 'OA', 391: 'OB',
                       392: 'OC', 393: 'OD', 394: 'OE', 395: 'OF', 396: 'OG', 397: 'OH', 398: 'OI', 399: 'OJ',
                       400: 'OK', 401: 'OL', 402: 'OM', 403: 'ON', 404: 'OO', 405: 'OP', 406: 'OQ', 407: 'OR',
                       408: 'OS', 409: 'OT', 410: 'OU', 411: 'OV', 412: 'OW', 413: 'OX', 414: 'OY', 415: 'OZ',
                       416: 'PA', 417: 'PB', 418: 'PC', 419: 'PD', 420: 'PE', 421: 'PF', 422: 'PG', 423: 'PH',
                       424: 'PI', 425: 'PJ', 426: 'PK', 427: 'PL', 428: 'PM', 429: 'PN', 430: 'PO', 431: 'PP',
                       432: 'PQ', 433: 'PR', 434: 'PS', 435: 'PT', 436: 'PU', 437: 'PV', 438: 'PW', 439: 'PX',
                       440: 'PY', 441: 'PZ', 442: 'QA', 443: 'QB', 444: 'QC', 445: 'QD', 446: 'QE', 447: 'QF',
                       448: 'QG', 449: 'QH', 450: 'QI', 451: 'QJ', 452: 'QK', 453: 'QL', 454: 'QM', 455: 'QN',
                       456: 'QO', 457: 'QP', 458: 'QQ', 459: 'QR', 460: 'QS', 461: 'QT', 462: 'QU', 463: 'QV',
                       464: 'QW', 465: 'QX', 466: 'QY', 467: 'QZ', 468: 'RA', 469: 'RB', 470: 'RC', 471: 'RD',
                       472: 'RE', 473: 'RF', 474: 'RG', 475: 'RH', 476: 'RI', 477: 'RJ', 478: 'RK', 479: 'RL',
                       480: 'RM', 481: 'RN', 482: 'RO', 483: 'RP', 484: 'RQ', 485: 'RR', 486: 'RS', 487: 'RT',
                       488: 'RU', 489: 'RV', 490: 'RW', 491: 'RX', 492: 'RY', 493: 'RZ', 494: 'SA', 495: 'SB',
                       496: 'SC', 497: 'SD', 498: 'SE', 499: 'SF', 500: 'SG'}
            month_start_col = str(rev_key[month_start_col_idx])
            # use the calculated start column values to generate raw input file. One raw file is generated per month
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Generating the raw file for month starting at column: {}'.format(
                month_start_col))
            days_in_month = days_in_month_list[i]
            GenerateReport(path, in_file, sheet, month_start_col, days_in_month, yr_start_col, status, offices,
                           report_items).inputFile()
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Finished generating the raw file for month starting at column: {}'.format(
                month_start_col))
            # raw file is generated as df2
            drop_col = [1, 2, 3]
            df2 = df2.drop(df2.columns[drop_col], axis=1)
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Generating df_ov_per_month for month starting at column: {}'.format(
                month_start_col))
            df_ov_per_month = df2.astype(np.object)
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Finished generating df_ov_per_month for month starting at column: {}'.format(
                month_start_col))
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Generating count_list for month starting at column: {}'.format(
                month_start_col))
            count_list = [[] for n in range(len(
                df_ov_per_month.columns))]  # Create a list of lists based on the number of columns. Each column will be an item (each item is a list) in this list
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Finished generating count_list for month starting at column: {}'.format(
                month_start_col))
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Start iterating through the columns in df_ov_per_month for month starting at column: {}'.format(
                month_start_col))
            for l in range(len(df_ov_per_month.columns)):
                print('[' + datetime.datetime.now().strftime(
                    "%d.%b %Y %H:%M:%S") + '] Looping column {} of df_ov_per_month for month starting at column: {}'.format(
                    l,
                    month_start_col))
                # iterate through all columns in the raw file. Find out which row contains 'Annual leave'. Then use these rows to form a new table called df_temp
                df_temp = df_ov_per_month[df_ov_per_month.iloc[:, l].fillna(0).astype(str).str.contains('Annual leave')]
                count_list[l].append(df_temp.iloc[:, 0].to_string(
                    index=False))  # each table is converted to a list of names. Each list of name is appended to the list of lists
            print('[' + datetime.datetime.now().strftime(
                "%d.%b %Y %H:%M:%S") + '] Starting to clean up the count_list of df_ov_per_month for month starting at column: {}'.format(
                month_start_col))
            for sub, o in zip(count_list, range(len(count_list))):  # clean up the list to get just the names
                print('[' + datetime.datetime.now().strftime(
                    "%d.%b %Y %H:%M:%S") + '] Cleaning up column {} of the count_list of df_ov_per_month for month starting at column: {}'.format(
                    o,
                    month_start_col))
                count_list[o] = str(sub).replace('Series([], )', '')
                print('Showing count_list in loop contents for df_ov_per_month starting at column: {}'.format(
                    month_start_col))
                print(count_list)
            print('Showing count_list contents for df_ov_per_month starting at column: {}'.format(month_start_col))
            count = counter - 1
            d_ov[month_list[count]] = count_list
            counter += 1

        df_ov_all = pd.DataFrame()

        for i, j in zip(range(len(month_list)), month_list):
            df_ov_all.insert(loc=i, column=j, value=pd.Series(d_ov[j]))

        df_ov_all_final = df_ov_all.drop([0])
        #print(df_ov_all_final.to_string())

    def exportOverview(self):
        df_ov_all_final.to_excel('Overview.xlsx', sheet_name='Overview')
        print('Annual overview report generation complete!')


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

#generate overview report to show who took annual leave each day
initiate.annualOverview()
initiate.exportOverview()