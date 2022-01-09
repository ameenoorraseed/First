from numpy import add
import pandas as pd
from pathlib import Path

def excel_diff(path_OLD, path_NEW): 
  
    df_OLD = pd.read_excel(path_OLD, sheet_name=0, keep_default_na=False)  
    df_NEW = pd.read_excel(path_NEW, sheet_name=1, keep_default_na=False)

    values = []
   
              
    print('Old '+str(len(df_OLD.columns)))
    print("New "+str(len(df_NEW.columns)))
    

    added_col=list(set(df_NEW.columns)-set(df_OLD.columns))
    removed_col=list(set(df_OLD.columns)-set(df_NEW.columns))
    key = ["Name", "NeType", "Category"]
    df_OLD = df_OLD.set_index(key)
    df_NEW = df_NEW.set_index(key)
    dfDiff = df_NEW.copy()
    cols_OLD = df_OLD.columns
    cols_NEW = df_NEW.columns

    sharedCols = list(set(cols_OLD).intersection(cols_NEW))
    col_notin_new=list(set(cols_OLD).difference(cols_NEW))
    col_notin_old=list(set(cols_NEW).difference(cols_OLD))
    print('Col not in old= ',col_notin_old)
    print('Col not in new= ',col_notin_new)
    dfDiff.insert(0, "Comment", '')
    

    values.append('Added Columns:{}'.format([c for c in col_notin_old]))
    values.append('Removed Columns:{}'.format([c for c in col_notin_new]))


    sx=[]
    for row in dfDiff.index:
        if (row in df_OLD.index) and (row in df_NEW.index):
            for col in sharedCols:
                value_OLD = df_OLD.loc[row, col]
                value_NEW = df_NEW.loc[row, col]
                if value_OLD == value_NEW:
                    dfDiff.loc[row, col] = df_NEW.loc[row, col]
                else:
                    sx.append(col)
                    dfDiff.loc[row,'Comment']='Modified Columns: {}'.format([c for c in sx])
            sx.clear()
    
    #writing all the comments at once in dataframe
    for i in range(len(values)):
        dfDiff['Comment'][i]=values[i]

    writer = pd.ExcelWriter(r"D:\output-delta.xlsx", engine="xlsxwriter")
    dfDiff.to_excel(writer, sheet_name="Delta", index=True)
    writer.save()

def main():
    path_OLD = Path(r"F:\facebook\Surabah\First-main (1)\First-main\Input.xlsx")
    path_NEW = Path(r"F:\facebook\Surabah\First-main (1)\First-main\Input.xlsx")
    excel_diff(path_OLD, path_NEW)

if __name__ == "__main__":
    main()

