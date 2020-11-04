import Functions.operational_functionality as ofn
import pandas as pd
import xlrd


def branch_wise_stocks_Records():
    print('Branch wise stock ...')
    df = pd.read_excel('Data/gpm_data.xlsx')
    excel_data_df = df[['BSL NO', 'BRAND', 'ISL NO',
                        'Item Name', 'UOM', 'BOG', 'BSL',
                        'COM', 'COX', 'CTG',
                        'CTN', 'DNJ', 'FEN', 'FRD', 'GZP',
                        'HZJ', 'JES', 'KHL', 'KRN', 'KSG',
                        'KUS', 'MHK', 'MIR', 'MLV',
                        'MOT', 'MYM', 'NAJ', 'NOK', 'PAT',
                        'PBN', 'RAJ', 'RNG', 'SAV', 'SYL',
                        'TGL', 'VRB']]

    bslno = ofn.create_dup_count_list(excel_data_df, 'BSL NO')
    brand = ofn.create_dup_count_list(excel_data_df, 'BRAND')

    wb = xlrd.open_workbook('Data/gpm_data.xlsx')
    sh = wb.sheet_by_name('Sheet1')
    print('branch Wise dataset Start printing in HTML')

    tabletd = ""
    for i in range(1, sh.nrows):
        tabletd = tabletd + "<tr>\n"
        for j in range(0, 1):
            # BSL NO
            if (bslno[i - 1] != 0):
                tabletd = tabletd + "<td class=\"serial\" rowspan=\"" + str(bslno[i - 1]) + "\"> " + str(
                    int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(1, 2):
            # Brand
            if (brand[i - 1] != 0):
                tabletd = tabletd + "<td class=\"brandtd\" rowspan=\"" + str(brand[i - 1]) + "\">" + str(
                    sh.cell_value(i, j)) + "</td>\n"

        for j in range(2, 3):
            # ITemNo
            tabletd = tabletd + "<td class=\"central\">" + str(int((sh.cell_value(i, j)))) + "</td>\n"

        for j in range(3, 4):
            # item name
            tabletd = tabletd + "<td class=\"description\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(4, 5):
            # unit
            tabletd = tabletd + "<td class=\"style1\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(5, 6):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(6, 7):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(7, 8):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(8, 9):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(9, 10):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(10, 11):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(11, 12):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(12, 13):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(13, 14):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(14, 15):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(15, 16):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(16, 17):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(17, 18):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(18, 19):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(19, 20):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(20, 21):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(21, 22):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(22, 23):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(23, 24):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(24, 25):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(25, 26):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(26, 27):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(27, 28):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(28, 29):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(29, 30):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(30, 31):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(31, 32):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(32, 33):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(33, 34):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(34, 35):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        for j in range(35, 36):
            tabletd = tabletd + "<td class=\"remarks\"style=\"background-color:" + \
                      str(ofn.branch_warning(sh.cell_value(i, j))) + "\">" + str(int(sh.cell_value(i, j))) + "</td>\n"

        table1 = tabletd + "</tr>\n"
    print("No Sales table Created")
    return table1
