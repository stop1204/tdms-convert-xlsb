import os
import nptdms as tdms
import xlwings as xw
from datetime import datetime, timedelta

# import pandas as pd

# from pathlib import Path  # 导入Path类，用于处理文件路径

# pyinstaller --upx-dir "D:\UPX\upx-4.2.3-win64" -F main.py
# --clean


debug = False  # Set to True for debugging, False for production
# Get the current working directory
# cwd = Path.cwd().joinpath("tdms-convert/")
# 获得当前运行目录
cwd = os.getcwd()
# 如果tdms文件夹不存在则创建
if not os.path.exists(os.path.join(cwd, "tdms")):
    os.mkdir(os.path.join(cwd, "tdms"))


date_format = "%y %m %d %H %M"  # 日期格式


# Create a list of all TDMS files in the "tdms" directory
tdms_files = []
for file in os.listdir(os.path.join(cwd, "tdms/")):
    if file.endswith(".tdms"):
        tdms_files.append(file)


def process_tdms_file(tdms_file, app):
    try:
        print("Converting " + tdms_file + " to CSV...")
        filename = tdms_file[:-5]
        with tdms.TdmsFile.open(
                os.path.join(cwd, "tdms", tdms_file), raw_timestamps="True"
        ) as tdms_file_obj:
            df = tdms_file_obj.as_dataframe()
            if debug:
                print(df)
            group = tdms_file_obj["COBRA THERMAL DATA"]
            channel = group["TIME"]
            channel_data = channel[:]
            channel_properties = channel.properties
            if debug:

                print(channel_data, channel_properties, tdms_file_obj.properties["name"])
            create_time = tdms_file_obj.properties["name"].split(" - ")[-1]
            create_time = datetime.strptime(create_time, date_format)
            if debug:

                print("create time: ", create_time, type(create_time))
            rows = df.shape[0]
            drop_columns = []
            try:
                drop_file_path = os.path.join(cwd, "drop_columns.txt")
                with open(drop_file_path, "r") as f:
                    drop_columns = f.read().splitlines()
                if debug :

                    print("drop columns: ", drop_columns)
                df = df.drop(
                    drop_columns,
                    axis=1,
                )
            except Exception as e:
                print("no such column", e)
            if debug:

                print("Continue...")
            df.insert(0, "DateTime", df.index)
            df.insert(1, "date", df.index)
            df.insert(2, "time_parse", df.index)
            if len(channel_data) == len(df):
                df["time_parse"] = channel_data / 86400
            else:
                if debug :

                    print(f"Length mismatch: channel_data({len(channel_data)}) != df({len(df)})")
                df["time_parse"] = [None] * len(df)
            df["date"] = create_time
            # df["DateTime"] = create_time
            # drop na df.column data top 10  /'COBRA THERMAL DATA'/'TIME'
            df = df.dropna(subset=["/'COBRA THERMAL DATA'/'TIME'"])
            df["DateTime"] = [create_time + timedelta(seconds=float(t)) for t in df["/'COBRA THERMAL DATA'/'TIME'"]]
            # df["DateTime"] = df["DateTime"].dt.strftime("%Y/%m/%d %H:%M:%S")
            df.columns = df.columns.str.replace("COBRA THERMAL DATA", "")
            df.columns = df.columns.str.replace("'", "")
            df.columns = df.columns.str.replace("/'", "")
            df.columns = df.columns.str.replace("//", "")
            if debug:

                print("write to csv/xlsb file...")
            if not os.path.exists(os.path.join(cwd, "csv")):
                os.mkdir(os.path.join(cwd, "csv"))
            if not os.path.exists(os.path.join(cwd, "xlsb")):
                os.mkdir(os.path.join(cwd, "xlsb"))
            path_ = os.path.join(cwd, "csv/" + filename + ".csv")


            # because droped time's NA data
            rows = df.shape[0]

            wb = app.books.add()
            wb.sheets[0]["A1"].value = df
            wb.sheets[0]["B2"].value = "=C2+D2"
            wb.sheets[0]["B2"].api.AutoFill(wb.sheets[0].range("B2:B" + str(rows)).api, 4)
            wb.sheets[0]["B:B"].api.NumberFormat = "yyyy/m/d hh:mm:ss"
            wb.sheets[0]["A1"].api.Font.Color = -16777216
            wb.sheets[0]["A1"].api.Interior.Color = 255
            wb.sheets[0]["A1"].value = "If you need further assistance, please contact Terry [terry.he@neworld.com.hk]"
            if debug:

                print("remove old data")


            # fix datetime + time

            #print the rows count in wb and df

            # Write the DateTime column to the correct Excel range
            wb.sheets[0]["B2:B" + str( rows + 1 )].value = [[dt] for dt in df["DateTime"].tolist()]


            wb.sheets[0]["B2:B" + str(rows)].api.Copy()
            wb.sheets[0]["B2:B" + str(rows)].api.PasteSpecial(-4163)
            wb.sheets[0]["C:D"].api.Delete()
            df = df.drop(
                ["date", "time_parse"],
                axis=1,
            )
            for column in drop_columns:
                if column in df.columns:
                    df = df.drop(column, axis=1)
            df.to_csv(path_, index=False)
            will_drop_columns = wb.sheets[0].range("A1").expand("right").value
            if debug:

                print(will_drop_columns)
            for i in range(len(will_drop_columns) - 1, -1, -1):
                if will_drop_columns[i] in drop_columns:
                    wb.sheets[0].range((1, i + 1)).api.EntireColumn.Delete()
                    if debug :

                        print("delete column: ", will_drop_columns[i])
            wb.save(path_.replace("csv", "xlsb"))
            print("Converted")
            wb.close()
    except Exception as e:
        print("failed to convert", tdms_file, e)


if __name__ == "__main__":
    max_threads = 1  # Only one thread to avoid Excel COM issues
    app = xw.App(visible=False)
    try:
        for tdms_file in tdms_files:
            process_tdms_file(tdms_file, app)
        print("TDMS files converted to CSV files successfully!")
    finally:
        app.quit()
