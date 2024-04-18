# folder structure
  - root:
  -  ./csv   // output path
  -  ./tdms  // input path
  -  ./xlsb  // output path
  -  main.exe
  -  drop_columns.txt // drop columns name
  
## Output CSV/XLSB
        # df.to_csv(path_, index=False)
        # wb = xw.Book(path_)
        # wb.set_mock_caller()
        # print(path_.replace(".csv", ".xlsb"))
        # app = xw.App(visible=False)
        # wb = app.books.open(path_.replace(".csv", ".xlsb"))

        
## CSV files need to be manually converted to time format: 
Formula: ```=DateTime + TIME/86400```