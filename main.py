import os
import nptdms as tdms
import xlwings as xw
from datetime import datetime, timedelta

# import pandas as pd

# from pathlib import Path  # 导入Path类，用于处理文件路径

# pyinstaller --upx-dir "D:\UPX\upx-4.2.3-win64" -F main.py
# --clean


# Get the current working directory
# cwd = Path.cwd().joinpath("tdms-convert/")
# 获得当前运行目录
cwd = os.getcwd()
# 如果tdms文件夹不存在则创建
if not os.path.exists(os.path.join(cwd, "tdms")):
    os.mkdir(os.path.join(cwd, "tdms"))


# app = xw.App()
# wb = xw.App()

date_format = "%y %m %d %H %M"  # 日期格式


# Create a list of all TDMS files in the "tdms" directory
tdms_files = []
for file in os.listdir(os.path.join(cwd, "tdms/")):
    if file.endswith(".tdms"):
        tdms_files.append(file)

# Convert each TDMS file to a CSV file
for tdms_file in tdms_files:
    print("Converting " + tdms_file + " to CSV...")
    filename = tdms_file[:-5]
    # Open the TDMS file
    with tdms.TdmsFile.open(
        os.path.join(cwd, "tdms", tdms_file), raw_timestamps="True"
    ) as tdms_file:
        # Create a Pandas DataFrame from the TDMS file
        df = tdms_file.as_dataframe()
        print(df)
        # print df column names
        group = tdms_file["COBRA THERMAL DATA"]  # sheet name
        channel = group["TIME"]  # column name
        channel_data = channel[:]  # column data
        channel_properties = channel.properties
        print(channel_data, channel_properties, tdms_file.properties["name"])
        # tdms_file.properties -> OrderedDict({'name': 'J3100000600073 - NI-sbRIO-9651-01ea5315 - 24 01 25 09 20'})
        # get the datetime from the file name  24 01 25 09 20
        create_time = tdms_file.properties["name"].split(" - ")[-1]
        # convert ' 24 01 25 09 20' to '24-01-25 09:20:00'
        create_time = datetime.strptime(create_time, date_format)
        print("create time: ", create_time, type(create_time))

        rows = df.shape[0]
        print("rows: ", rows)
        # Save the DataFrame to a CSV file
        # remove '/'COBRA THERMAL DATA'/'' from title

        # remove column /'Network'/'Network'  and /'Errors'/'Untitled'
        try:
            # read file to drop columns
            drop_file_path = os.path.join(cwd, "drop_columns.txt")
            drop_columns = []
            with open(drop_file_path, "r") as f:
                drop_columns = f.read().splitlines()
            print("drop columns: ", drop_columns)
            df = df.drop(
                drop_columns,
                axis=1,
            )

        except:
            print("no such column")
        print("Continue...")
        # insert datetime column to the first column
        df.insert(0, "DateTime", df.index)
        df.insert(1, "date", df.index)
        df.insert(2, "time_parse", df.index)
        # add data to the datetime column
        # time-parse = column 'TIME' / 86400
        df["time_parse"] = channel_data / 86400
        df["date"] = create_time
        # //2024-01-25 09:20:00 <class 'datetime.datetime'>
        df["DateTime"] = create_time

        df.columns = df.columns.str.replace("COBRA THERMAL DATA", "")
        # replace /' and '
        df.columns = df.columns.str.replace("'", "")
        # replace /' and '
        df.columns = df.columns.str.replace("/'", "")
        df.columns = df.columns.str.replace("//", "")
        print("write to csv/xlsb file...")
        # 写入到 当前目录下的csv文件夹中 如果文件夹不存在则创建
        if not os.path.exists(os.path.join(cwd, "csv")):
            os.mkdir(os.path.join(cwd, "csv"))
        path_ = os.path.join(cwd, "csv/" + filename + ".csv")

        # df.to_csv(path_, index=False)
        # wb = xw.Book(path_)
        # wb.set_mock_caller()
        # print(path_.replace(".csv", ".xlsb"))
        # app = xw.App(visible=False)
        # wb = app.books.open(path_.replace(".csv", ".xlsb"))
        wb = xw.Book()
        wb.sheets[0]["A1"].value = df
        # B 2 : B rows = C2+D2 ... C3+D3
        wb.sheets[0]["B2"].value = "=C2+D2"
        wb.sheets[0]["B2"].api.AutoFill(wb.sheets[0].range("B2:B" + str(rows)).api, 4)
        # B:B format yyyy/m/d hh:mm:ss
        wb.sheets[0]["B:B"].api.NumberFormat = "yyyy/m/d hh:mm:ss"
        # set A1 to red background color and white font color
        wb.sheets[0]["A1"].api.Font.Color = -16777216
        wb.sheets[0]["A1"].api.Interior.Color = 255
        # Terry terry.he@neworld.com.hk
        wb.sheets[0][
            "A1"
        ].value = "If you need further assistance, please contact Terry [terry.he@neworld.com.hk]"
        # delete column 'C:D'
        print("remove old data")
        # 将 B 的formular 值赋值给 B
        wb.sheets[0]["B2:B" + str(rows)].api.Copy()
        wb.sheets[0]["B2:B" + str(rows)].api.PasteSpecial(-4163)
        wb.sheets[0]["C:D"].api.Delete()
        df = df.drop(
            ["date", "time_parse"],
            axis=1,
        )
        df.to_csv(path_, index=False)
        # print wb columns
        will_drop_columns = wb.sheets[0].range("A1").expand("right").value
        print(will_drop_columns)
        # wb drop column  , 多重防护, 防止没删掉, 用index,  倒叙删除

        for i in range(len(will_drop_columns) - 1, -1, -1):
            if will_drop_columns[i] in drop_columns:
                wb.sheets[0].range((1, i + 1)).api.EntireColumn.Delete()
                print("delete column: ", will_drop_columns[i])
        wb.save(path_.replace(".csv", ".xlsb"))
        wb.close()
    tdms_file.close()


# Print a message to the console
print("TDMS files converted to CSV files successfully!")
# app.kill()
