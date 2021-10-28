#!/usr/bin/env python3

def sum_legends(my_df):

    KB = []
    D = []
    LOA = []
    TB1 = []
    TB2 = []
    ST = []
    for each in my_df.index:
        x = (my_df.loc[each] == "KB").sum()
        KB.append(x)
        y = (my_df.loc[each] == "D").sum()
        D.append(y)
        z = (my_df.loc[each] == "LOA").sum()
        LOA.append(z)
        h = (my_df.loc[each] == "TB1").sum()
        TB1.append(h)
        n = (my_df.loc[each] == "TB2").sum()
        TB2.append(n)
        m = (my_df.loc[each] == "ST").sum()
        ST.append(m)
# Adding the separate sets to the main data frame
    my_df["KB"] = KB
    my_df["LOA"] = LOA
    my_df["D"] = D
    my_df["TB1"] = TB1
    my_df["TB2"] = TB2
    my_df["ST"] = ST
    return my_df


def clean_df_na_set_index(my_df,col_num):
    my_df = my_df.dropna(subset=[col_num])
    my_df[col_num] = my_df[col_num].astype(int)
    # my_df.loc[:,col_num] = my_df.loc[:,col_num].astype(int)
    my_df = my_df.set_index(col_num)
    return my_df

def merge_two_dataframe(df_crew, df_cost):
    # clean dataframe
    # df_cost.drop(columns=['BL', 'COVID' ,'Sub BL', 'Employee Name', 'Type', 'Function Description','International Type'], axis=1, inplace=True)
    df_cost.fillna(0, inplace=True)
    df = df_crew.join(df_cost)
    df.dropna(subset=['Wellsite Rate (J1)','Base Rate (Daily Rate)'],inplace=True)
    return df





