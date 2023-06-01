#import library pandas
import pandas as pd

## create data frame from the excel file placed in your folder
df = pd.read_excel(r"C:\Users\u130628\Downloads\O2C March Data.xlsx", sheet_name=0, skiprows=13, index_col=None)

## Add columns by merging existing columns
df.insert(2,"Key Indicator", df['Unnamed: 1'].astype(str) + df["Question Description"])

df.insert(9,"Answer",'')
for i in df.index:
    if df['Survey Answer'][i] != "#":
        ans = df['Survey Answer'][i]
    else:
        ans = df['Survey Answer Text'][i]
    df['Answer'][i] = ans

## replace values in the column
df.loc[df['Question Description'].str.contains('detailed'), 'Question Description'] = 'What was the detailed cause of the issue?'

## Get all the unique values for survey questions
survey_que = (df['Question Description'].unique())

## create another copy of this dataframe
df_upd = df.copy()

## drop columns from copied data frame that are not required
df_upd.drop(df.columns[[2,6,7,8,9]],axis= 1,inplace=True)

## rename unnamed columns
df_upd.rename(columns={"Unnamed: 1":"Ticket #", "Unnamed: 9":"Account #"}, inplace=True)

## remove duplicate rows
df_upd.drop_duplicates(inplace=True)

## create empty new columns for each question
n = 2
for i in survey_que:
    df_upd.insert(n,i,'')
    n = n+1

## insert answers in each columns as per the values in the main dataframe
for col in df_upd.columns[2:9]:
    for i in df_upd.index:
        merge = str(df_upd['Ticket #'][i]) + col
        if merge in df['Key Indicator'].values:
            loc = df[df['Key Indicator']==merge].index.values.tolist()
            answer = df['Answer'][loc[0]]
            df_upd[col][i] = answer
        else:
            df_upd[col][i] = "NA"


## change date format for below columns
df_upd['Reported On Date']= df_upd['Reported On Date'].dt.strftime('%m/%d/%y')
df_upd['Completion Due Date']= df_upd['Completion Due Date'].dt.strftime('%m/%d/%y')

## import dataframe to excel
df_upd.to_excel("O2C_Updated.xlsx", index=False)
