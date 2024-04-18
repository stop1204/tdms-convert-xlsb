# 默认输出 xlsb
  - root:
  -  ./csv   // output path
  -  ./tdms  // input path
  -  main.exe
  -  drop_columns.txt // drop columns name
## 输出CSV
        # df.to_csv(path_, index=False)
        # wb = xw.Book(path_)
        # wb.set_mock_caller()
        # print(path_.replace(".csv", ".xlsb"))
        # app = xw.App(visible=False)
        # wb = app.books.open(path_.replace(".csv", ".xlsb"))

        
