import os
import msoffcrypto
import xlrd
from datetime import datetime
import logging
import re
import pandas as pd


inputDir = r'\\ '
車両情報_output = r'\\' + '/車両情報.xlsx'
付帯売上_output = r'\\' + '/付帯売上.xlsx'
err_output = r'\\' + '/err_output.xlsx'

# today = datetime.datetime.now()

japanCalendar = {
    "2020":("R02", "令和02"),
    "2021":("R03", "令和03"),
    "2022":("R04", "令和04"),
    "2023":("R05", "令和05")
}


# year = today.year
# month = today.month
# r_year = year - 2018
# folder_name = "令和0"+ str(r_year) + "年0" +str(month) + "月各店分"
 
# print(folder_name)

days = ["月", "火", "水", "木", "金", "土", "日"]

used = {
    "二輪新車": "新車",
    "二輪中古": "中古",
    "四輪新車": "新車",
    '四輪中古': "中古"
}
def check_year(input_str):
    pattern = r'^\d{4}$'
    if not re.match(pattern, input_str):
        return False

    year = int(input_str)
    if year <= 2020 or year > datetime.now().year:
        return False
    return True

def is_valid_month(input_str):
    pattern = r'^\d{2}$'
    if re.match(pattern, input_str) is not None:
        month_num = int(input_str)
        return 1 <= month_num <= 12
    return False
# logging.basicConfig(level=logging.DEBUG)

def iterate_files(inputDir):
    user_input_year = input("Please enter year (e.g., 2023) (enter 'q' to stop): ")
    while True:
        if user_input_year.lower() == 'q':
            print("Stop application.")
            return None
        if check_year(user_input_year):
            print(f"Selected year {user_input_year} successfully.")
            break
        else:
            print(f"{user_input_year} is not a valid year, please try again (e.g., 2023).")
            user_input_year = input("Please enter year (enter 'q' to stop): ")

    user_input_month = input("Please enter month (e.g., 05) (enter 'q' to stop): ")
    while True:
        if user_input_month.lower() == 'q':
            print("Stop application.")
            return None
        if is_valid_month(user_input_month):
            print(f"Selected month {user_input_month} successfully.")
            break
        else:
            print(f"{user_input_month} is not a valid month, please try again (e.g., 05).")
            user_input_month = input("Please enter month (enter 'q' to stop): ")


    inputDir += f"/{user_input_year}年度({japanCalendar[user_input_year][0]}年)各店分/{japanCalendar[user_input_year][1]}年{user_input_month}月各店分"

    for filename in os.listdir(inputDir):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            password = '0' + filename[:2]
            read_password_protected_excel(inputDir + '/' + filename, password)


def read_password_protected_excel(file_path, password):
    try:
        # Create an encrypted file object
        enc_file = msoffcrypto.OfficeFile(open(file_path, "rb"))

        # Attempt to decrypt the file using the provided password
        enc_file.load_key(password=password)

        # Create a new decrypted file object
        dec_file = open(file_path + ".tmp", "wb")

        # Decrypt the file and save it as a temporary file
        enc_file.decrypt(dec_file)
        dec_file.close()

        df_車両情報, df_付帯売上, df_err = process_data(file_path + ".tmp")
        write_to_excel(df_車両情報, df_付帯売上, df_err, 車両情報_output, 付帯売上_output,
                      err_output)
    except Exception as e:
        print(f"Caught an exception: {e}")
    finally:
        # Remove the temporary file
        os.remove(file_path + ".tmp")

    # Return data
    return


def write_to_excel(df_車両情報, df_付帯売上, df_err, 車両情報_output, 付帯売上_output, err_output):
    if os.path.isfile(車両情報_output):
        existing_df = pd.read_excel(車両情報_output)   #fileがあった場合、読み込んで連結
        df_車両情報 = pd.concat([existing_df,df_車両情報], ignore_index=True)
        
    # df_車両情報=df_車両情報.drop_duplicates(keep='first', inplace=True)
    else:
        df_車両情報=df_車両情報
    with pd.ExcelWriter(車両情報_output) as writer:
        df_車両情報 = df_車両情報.sort_values( by=['店舗','契約日付'], ascending=[True, True])
        df_車両情報.to_excel(writer, index=False)     

    if os.path.isfile(付帯売上_output):
        existing_df = pd.read_excel(付帯売上_output)
        df_付帯売上 = pd.concat([existing_df, df_付帯売上], ignore_index=True)

    # df_付帯売上= df_付帯売上.drop_duplicates(keep = 'first', inplace = True)

    else:
         df_付帯売上= df_付帯売上
    with pd.ExcelWriter(付帯売上_output) as writer:
       
        df_付帯売上 = df_付帯売上.sort_values(by=['店舗','契約日付'], ascending=[True, True])
        df_付帯売上.to_excel(writer, index=False)      
                
    if os.path.isfile(err_output):
        existing_df = pd.read_excel(err_output)
        df_err = pd.concat([existing_df, df_err], ignore_index =True)
    # df_err= df_err.drop_duplicates(keep ='first', inplace = True)
    else:
        df_err=df_err  
    with pd.ExcelWriter(err_output) as writer: 
        df_err.to_excel(writer, index = False)


def process_data(filePath):
    workbook = xlrd.open_workbook(filePath)

    df_車両情報 = pd.DataFrame(
        columns=['契約日付', '曜日', '店舗', '新古区分', 'メーカー', '車種', '排気量', '色', 'フレームNo', '売上金額']
    )
    df_付帯売上 = pd.DataFrame(
        columns=['契約日付', '曜日', '店舗', '修理(パーツ)', '修理(工賃)', '保険', '外注工賃', 'ブルーマウンテン売上']
    )
    df_err = pd.DataFrame(
        columns=['File Path', 'Sheet Name', 'Err', 'Action']
    )

    for sheet_name in workbook.sheet_names()[3:]:
        logging.debug("filePath： " + filePath + " sheet_name: " + sheet_name)
        # sheet_name = '１日'
        sheet = workbook.sheet_by_name(sheet_name)
        invalidSheet = False
        checkedData = {
            '契約日付': False,
            '店舗': False
        }
        collectedData = {
            '契約日付': '',
            '曜日': '',
            '店舗': '',
            '新古区分': '',
            'メーカー': None,
            '車種': None,
            '排気量': None,
            '色': None,
            'フレームNo': None,
            '売上金額': None,
            '修理(パーツ)': None,
            '修理(工賃)': None,
            '保険': None,
            '外注工賃': None,
            'ブルーマウンテン売上': None
        }
        # i: row
        # j: column
        for i in range(0,43):
            row = sheet.row_values(i)
            for j in range(len(row)):
                cell = sheet.cell(i, j)
                if cell.ctype == 0:
                    continue
                elif not checkedData['契約日付'] and cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt = datetime(*xlrd.xldate_as_tuple(cell.value, workbook.datemode))
                        collectedData['契約日付'] = dt.strftime('%Y-%m-%d')
                        day_of_week = dt.weekday()
                        collectedData['曜日'] = days[day_of_week]
                        checkedData['契約日付'] = True
                    except Exception as e:
                        print(f"Caught an exception when parsing 契約日付: {cell.value} {e}")
                        err = {
                            'File Path': filePath,
                            'File Name': os.path.splitext(os.path.basename(filePath)),
                            'Sheet Name': sheet_name,
                            'Err': 'Caught an exception when parsing 契約日付',
                            'Action': 'Raised an exception and skipped this sheet'
                        }
                        df_err = df_err._append(err, ignore_index=True)
                        invalidSheet = True
                        break
                elif not checkedData['店舗'] and cell.ctype == 1 and '店名（' in str(cell.value):
                    j += 2
                    collectedData['店舗'] = sheet.cell(i, j).value
                    checkedData['店舗'] = True
                elif cell.ctype == 1 and re.sub(r'\s', '', str(cell.value)) in used.keys():
                    collectedData['新古区分'] = (used[re.sub(r'\s', '', str(cell.value))])
                # elif cell.ctype == 1 and ('メーカー' in str(cell.value) or 'ﾒｰｶｰ' == cell.value):
                #     collectedData['メーカー'] = (cell, i, j)
                elif cell.ctype == 1 and '車種' in str(cell.value):
                    collectedData['車種'] = (cell, i, j)
                elif cell.ctype == 1 and '排気量' in str(cell.value):
                    collectedData['排気量'] = (cell, i, j)
                elif cell.ctype == 1 and '色' in str(cell.value):
                    collectedData['色'] = (cell, i, j)
                elif cell.ctype == 1 and 'ﾌﾚｰﾑNo' in str(cell.value):
                    collectedData['フレームNo'] = (cell, i, j)
                elif cell.ctype == 1 and '売上金額' in str(cell.value):
                    collectedData['売上金額'] = (cell, i, j)
                elif cell.ctype == 1 and '修理（パーツ）Ｆ' in re.sub(r'\s', '', str(cell.value)):
                    v = sheet.cell(i, j + 3).value
                    if v == '':
                        collectedData['修理(パーツ)'] = 0
                    else:
                        collectedData['修理(パーツ)'] = float(sheet.cell(i, j + 3).value)
                    # collectedData['修理(パーツ)'] = sheet.cell(i, j + 3).value
                    # logging.debug('修理(パーツ) is ' + cell.value)
                elif cell.ctype == 1 and '修理（工賃）Ｇ' in re.sub(r'\s', '', str(cell.value)):
                    v = sheet.cell(i, j + 3).value
                    if v == '':
                        collectedData['修理(工賃)'] = 0
                    else:
                        try:
                            collectedData['修理(工賃)'] = float(sheet.cell(i, j + 3).value)
                        except Exception as e:
                            print(f"Caught an exception when parsing 修理(工賃): {sheet.cell(i, j + 3).value} {e}")
                            err = {
                                'File Path': filePath,
                                'File Name': os.path.splitext(os.path.basename(filePath)),
                                'Sheet Name': sheet_name,
                                'Err': 'Caught an exception when parsing 修理(工賃)' + str(e),
                                'Action': 'Raised an exception and skipped this sheet'
                            }
                            df_err = df_err._append(err, ignore_index=True)
                            invalidSheet = True
                            break
                    # collectedData['修理(工賃)'] = sheet.cell(i, j + 3).value
                    logging.debug('修理(工賃) is ' + cell.value)
                elif cell.ctype == 1 and '保険Ｈ' in re.sub(r'\s', '', str(cell.value)):
                    v = sheet.cell(i, j + 3).value
                    if v == '':
                        collectedData['保険'] = 0
                    else:
                        collectedData['保険'] = float(sheet.cell(i, j + 3).value)
                    # collectedData['保険'] = sheet.cell(i, j + 3).value
                    logging.debug('保険 is ' + cell.value)
                elif cell.ctype == 1 and '外注工賃Ｉ' in re.sub(r'\s', '', str(cell.value)):
                    v = sheet.cell(i, j + 3).value
                    if v == '':
                        collectedData['外注工賃'] = 0
                    else:
                        collectedData['外注工賃'] = float(sheet.cell(i, j + 3).value)
                    # collectedData['外注工賃'] = sheet.cell(i, j + 3).value
                    logging.debug('外注工賃 is ' + cell.value)
                elif cell.ctype == 1 and 'ブルーマウンテン売上Ｋ' in re.sub(r'\s', '', str(cell.value)):
                    v = sheet.cell(i, j + 7).value
                    if v == '':
                        collectedData['ブルーマウンテン売上'] = 0
                    else:
                        collectedData['ブルーマウンテン売上'] = float(sheet.cell(i, j + 7).value)
                    # collectedData['ブルーマウンテン売上'] = sheet.cell(i, j + 7).value == '' ? 0 : float(sheet.cell(i, j + 7).value)
                    logging.debug('ブルーマウンテン売上 is ' + cell.value)
            if invalidSheet:
                break
            if collectedData['車種'] is not None:
                newCell = sheet.cell(collectedData['車種'][1], 1)
                tempRow = collectedData['車種'][1]
                while (newCell.value):
                    tempRow += 1
                    newCell = sheet.cell(tempRow, 1)
                    # print(newCell)

                if tempRow == collectedData['車種'][1] + 1:
                    collectedData['メーカー'] = None
                    collectedData['車種'] = None
                    collectedData['排気量'] = None
                    collectedData['色'] = None
                    collectedData['フレームNo'] = None
                    collectedData['売上金額'] = None
                    continue

                if (collectedData['契約日付'] and collectedData['曜日'] and collectedData[
                    '店舗'] and collectedData['新古区分'] and collectedData['車種'] and
                        collectedData[
                            '排気量'] and
                        collectedData['色'] and collectedData['フレームNo'] and collectedData['売上金額']):
                    for carRow in range(collectedData['車種'][1] + 1, tempRow):
                        new_row = {
                            '契約日付': collectedData['契約日付'],
                            '曜日': collectedData['曜日'],
                            '店舗': collectedData['店舗'],
                            '新古区分': collectedData['新古区分'],
                            'メーカー': sheet.cell(carRow, 1).value,
                            '車種': sheet.cell(carRow, 2).value,
                            '排気量': sheet.cell(carRow, 3).value,
                            '色': sheet.cell(carRow, 4).value,
                            'フレームNo': sheet.cell(carRow, 5).value,
                            '売上金額': sheet.cell(carRow, 9).value
                        }
                        df_車両情報 = df_車両情報._append(new_row, ignore_index=True)
                        # logging.debug('added new row: ' + str(new_row))
                    collectedData['メーカー'] = None
                    collectedData['車種'] = None
                    collectedData['排気量'] = None
                    collectedData['色'] = None
                    collectedData['フレームNo'] = None
                    collectedData['売上金額'] = None
                else:
                    err = {
                        'File Path': filePath,
                        'File Name': os.path.splitext(os.path.basename(filePath)),
                        'Sheet Name': sheet_name,
                        'Err': 'Invalid format, missing 曜日/新古区分/メーカー/車種/排気量/色/フレームNo/売上金額',
                        'Action': 'Raised an exception '
                    }
                    df_err = df_err._append(err, ignore_index=True)
                    # raise Exception('Invalid format in ' + str(filePath) + ' ' + str(sheet_name))
        if invalidSheet:
            continue
        if collectedData['修理(パーツ)'] is not None and collectedData['修理(工賃)'] is not None and collectedData[
            '保険'] is not None and collectedData['外注工賃'] is not None and collectedData['ブルーマウンテン売上'] is not None:
            if collectedData['修理(パーツ)'] ==0 and collectedData['修理(工賃)'] ==0 and collectedData[
            '保険'] ==0 and collectedData['外注工賃'] ==0 and collectedData['ブルーマウンテン売上'] ==0:
                continue
            else: 
                new_row = {
                    '契約日付': collectedData['契約日付'],
                    '曜日': collectedData['曜日'],
                    '店舗': collectedData['店舗'],
                    '修理(パーツ)': collectedData['修理(パーツ)'],
                    '修理(工賃)': collectedData['修理(工賃)'],
                    '保険': collectedData['保険'],
                    '外注工賃': collectedData['外注工賃'],
                    'ブルーマウンテン売上': collectedData['ブルーマウンテン売上']
                }
            df_付帯売上 = df_付帯売上._append(new_row, ignore_index=True)
        else:
            err = {
                    'File Path': filePath,
                    'File Name': os.path.splitext(os.path.basename(filePath)),
                    'Sheet Name': sheet_name,
                    'Err': 'Invalid format, missing 修理(パーツ)/修理(工賃)/保険/外注工賃/ブルーマウンテン売上',
                    'Action': 'Raised an exception '
                }
            df_err = df_err._append(err, ignore_index=True)

    return df_車両情報, df_付帯売上, df_err


# if __name__ == "__main__":
#     if os.path.exists(付帯売上_output):
#         os.remove(付帯売上_output)
#         print(付帯売上_output + ' has been deleted')
#     if os.path.exists(車両情報_output):
#         os.remove(車両情報_output)
#         print(車両情報_output + ' has been deleted')
#     if os.path.exists(err_output):
#         os.remove(err_output)
#         print(err_output + ' has been deleted')
iterate_files(inputDir)