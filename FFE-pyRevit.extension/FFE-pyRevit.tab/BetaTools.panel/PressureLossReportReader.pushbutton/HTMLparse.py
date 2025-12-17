# import pandas as pd






def HTMLparse(table_CP, table_Duct, table_Fittings):

    ### Get Critical Path from report ###
    list_CriticalPath = table_CP.iloc[-1,1].split(' ')[3].split('-')


    ### Get "Detail Information of Straight Segment by Sections" table from report ###
    dfTable_Duct = table_Duct

    ### Get "Fitting and Accessory Loss Coefficient Summary by Sections" table from report ###
    dfTable_Fittings = table_Fittings


    ### Drop all rows where 'Mark' value equals "NA" or "NaN" || Create new DataFrames ###
    df_DuctSections = dfTable_Duct.loc[dfTable_Duct.iloc[:,-1].isnull()]
    df_FittingsSections = dfTable_Fittings.loc[dfTable_Fittings.iloc[:,-1].isnull()]


    ### list of column names ###
    list_DuctSections_columns = list(df_DuctSections.columns)
    list_FittingsSections_columns = list(df_FittingsSections.columns)


    ### Drop unwanted columns ###
    df_DuctSections = df_DuctSections.drop(columns=list_DuctSections_columns[2:])
    df_FittingsSections = df_FittingsSections.drop(columns=list_FittingsSections_columns[2:])


    ### Rename Columns ###
    df_DuctSections.columns = ['Section', 'Details']
    df_FittingsSections.columns = ['Section', 'Details']


    ### Drop all rows where 'Mark' value DOES NOT equal "NA" or "NaN" || Create new DataFrame ###
    df_DuctReport = dfTable_Duct.loc[dfTable_Duct.iloc[:,-1].notnull()] 
    df_FittingsReport = dfTable_Fittings.loc[dfTable_Fittings.iloc[:,-1].notnull()]


    ### Insert "Section" column at position 0 of Dataframe ###
    df_DuctReport.insert(0, 'Section', 0) 
    df_FittingsReport.insert(0, 'Section', 0)


    ### Collect "Element ID" without duplicates ###
    set_DuctID = list(set(df_DuctReport['Element ID'])) 
    set_FittingsID = list(set(df_FittingsReport['Element ID']))


    ### DUCT: Loop through "Element ID" and add "Section" value ###
    for Row in set_DuctID:
        list_Section = list(df_DuctSections['Section'].loc[df_DuctSections['Details'].str.contains(str(Row))])
        list_Index = list(df_DuctReport.loc[df_DuctReport['Element ID'] == Row].index)
        df_DuctReport.loc[list_Index, 'Section'] = list_Section


    ### FITTINGS: Loop through "Element ID" and add "Section" value ###
    for Row in set_FittingsID:
        list_Section = list(df_FittingsSections['Section'].loc[df_FittingsSections['Details'].str.contains(str(Row))])
        list_Index = list(df_FittingsReport.loc[df_FittingsReport['Element ID'] == Row].index)
        df_FittingsReport.loc[list_Index, 'Section'] = list_Section


    df_DuctReport.loc[: , 'Category'] = 'Duct'
    df_FittingsReport.loc[: , 'Category'] = 'Fitting'


    ### Renumber index ###
    df_DuctReport = df_DuctReport.reset_index(drop=True)
    df_FittingsReport = df_FittingsReport.reset_index(drop=True)



    df_DuctReport.loc[: , 'Section'] = df_DuctReport['Section'].apply(str)
    df_FittingsReport.loc[: , 'Section'] = df_FittingsReport['Section'].apply(str)



    ### Convert Values (string to int64 or float64) ###
    df_DuctReport['Velocity'] = df_DuctReport['Velocity'].str.replace(' FPM', '')
    df_DuctReport['Velocity'] = df_DuctReport['Velocity'].astype('int64')

    df_DuctReport['Friction'] = df_DuctReport['Friction'].str.replace(' in-wg/100ft', '')
    df_DuctReport['Friction'] = df_DuctReport['Friction'].astype('float64')

    df_DuctReport['Flow'] = df_DuctReport['Flow'].str.replace(' CFM', '')
    df_DuctReport['Flow'] = df_DuctReport['Flow'].astype('int64')

    df_DuctReport['Pressure Loss'] = df_DuctReport['Pressure Loss'].str.replace(' in-wg', '')
    df_DuctReport['Pressure Loss'] = df_DuctReport['Pressure Loss'].astype('float64')
    df_FittingsReport['Pressure Loss'] = df_FittingsReport['Pressure Loss'].str.replace(' in-wg', '')
    df_FittingsReport['Pressure Loss'] = df_FittingsReport['Pressure Loss'].astype('float64')


    ### Drop all Fittings with Pressure Loss of 0 (Zero) ###
    # df_FittingsReport = df_FittingsReport[df_FittingsReport['Pressure Loss'] != 0]


    ### Filter Ducts that are in Revit's Critical Path ###
    df_DuctCP = df_DuctReport[df_DuctReport['Section'].isin(list_CriticalPath)]
    df_FittingsCP = df_FittingsReport[df_FittingsReport['Section'].isin(list_CriticalPath)]



    ### Convert Feet & Inches to Decimal Feet ###
    # df_DuctReport = df_DuctReport[['Length']].assign(**df_DuctReport['Length'].str.extract(r"(?P<Feet>\d+)'-\s?(?P<Inches>\d+)\s?(?P<Num>\d+)?\/?(?P<Dem>\d+)?\"").astype(float).fillna(0))
    # df_DuctReport['Length Decimal'] = df_DuctReport.eval("Feet + Inches / 12") + np.where(df_DuctReport.Num == 0,0,(df_DuctReport["Num"]/df_DuctReport["Dem"])/12)



    ### Combine Duct & Fitting DataFrames ###
    df_DuctData = pd.concat([df_FittingsReport, df_DuctReport])
    df_DuctData_CP = pd.concat([df_FittingsCP, df_DuctCP])


    ### Add 'Flow' value to Fitting rows ###
    Sections = df_DuctData['Section'].unique()
    for i in Sections:
        df_DuctData.loc[df_DuctData['Section'] == i, 'Flow'] = df_DuctData.loc[df_DuctData['Section'] == i, 'Flow'].max()
        df_DuctData_CP.loc[df_DuctData_CP['Section'] == i, 'Flow'] = df_DuctData_CP.loc[df_DuctData_CP['Section'] == i, 'Flow'].max()


    ### Renumber index ###
    df_DuctData = df_DuctData.reset_index(drop=True)
    df_DuctData_CP = df_DuctData_CP.reset_index(drop=True)


    ### Rearrange Columns ###
    ColumnIndex = ["System Name", "Category", "Element ID", "Type Mark", "ASHRAE Table", "Comments", "Section", "Size", "Flow", "Length", "Velocity", "Friction", "Pressure Loss"]

    df_DuctData = df_DuctData.reindex(columns = ColumnIndex)
    df_DuctData_CP = df_DuctData_CP.reindex(columns = ColumnIndex)


    ### Rename column "Comments" to "Critical Path" ###
    df_DuctData = df_DuctData.rename(columns={"Comments" : "Critical Path"})
    df_DuctData_CP = df_DuctData_CP.rename(columns={"Comments" : "Critical Path"})


    list_dfData = [df_DuctData, df_DuctData_CP]
    return list_dfData