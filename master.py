import pandas as pd
import difflib 
import pprint
import phonenumbers
from email.utils import parseaddr
import numpy as np 

def validNumber(phone_number):
    """Checks if phone Number is valid
    
    Reference:
    https://stackoverflow.com/questions/15258708/python-trying-to-check-for-a-valid-phone-number

    Args:
        phone_number (str): string representing a phone number

    Returns:
        Tre: _description_
    """
    if len(phone_number) != 12:
        return False
    for i in range(12):
        if i in [3,7]:
            if phone_number[i] != '-':
                return False
        elif not phone_number[i].isalnum():
            return False
    return True

if __name__=="__main__":
    master_workbook='PB_CLE_CRM.xlsx'
    master_worksheet='PB CLE Master List'
    
    workbooks_to_combine = ['PB_CLE_CRM.xlsx']
    worksheets_to_combine = ['Jan 9th RSVP']
    
    master = pd.ExcelFile(master_workbook)
    df_master = master.parse(master_worksheet)

    for i in range(len(workbooks_to_combine)):
        book = pd.ExcelFile(workbooks_to_combine[i])
        df_sheet = book.parse(worksheets_to_combine[i])
        
        '''Update data in existing columns of master'''
        print(f'Checking column matches for {master_workbook}/{master_worksheet} with {workbooks_to_combine[i]}/{worksheets_to_combine[i]}')
        column_map = dict() # Build column map
        for c in df_master.columns:
            matches = difflib.get_close_matches(c,df_sheet.columns)
            if len(matches)>0:
                print(f'"{c}" matches "{matches[0]}"')
                column_map[c] = matches[0]
        
        print('Column map for {master_workbook}/{master_worksheet} to {workbooks_to_combine[i]}/{worksheets_to_combine[i]}')
        pprint.pprint(column_map)

        # '''Update data'''
        # # Check for Lastname, Firstname matches
        df_master_clean = df_master.dropna(subset=['Last Name','First Name']) # Drop Nan for lastname and firstname
        df_master_clean["Phone"].replace(np.nan, '', regex=True,inplace=True)
        df_master_clean["Email"].replace(np.nan, '', regex=True,inplace=True)
        df_master_clean["Phone"] = df_master_clean["Phone"].astype("string")
        df_master_clean["Email"] = df_master_clean["Email"].astype("string")
        # mLastNames = list(map(str.strip, list(df_master_clean['Last Name'])))
        # mFirstNames = list(map(str.strip, list(df_master_clean['First Name'])))
        # mEmails = list(map(str.strip, list(df_master_clean['Email'].)))
        # mEmails2 = list(map(str.strip, list(df_master_clean['Email 2'])))
        # mPhones = list(map(str.strip, list(df_master_clean['Phone'])))

        df_sheet_clean = df_sheet.dropna(subset=["First name","Last name"]) # Drop Nan for lastname and firstname
        df_sheet_clean["Phone"].replace(np.nan, '', regex=True,inplace=True)
        df_sheet_clean["Email"].replace(np.nan, '', regex=True,inplace=True)
        df_sheet_clean["Phone"] = df_sheet_clean["Phone"].astype("string")
        df_sheet_clean["Email"] = df_sheet_clean["Email"].astype("string")
        df_sheet_clean["I live in Cleveland Ward # "].replace(np.nan, -1, regex=True,inplace=True)
        # sFirstNames = list(map(str.strip, list(df_sheet_clean[column_map['First Name']])))
        # sLastNames = list(map(str.strip, list(df_sheet_clean[column_map['Last Name']])))
        # sWards = list(map(str.strip, list(df_sheet_clean['I live in Cleveland Ward #'])))
        # sPhones = list(map(str.strip, list(df_sheet_clean['Phone'])))
        # sEmails = list(map(str.strip, list(df_sheet_clean['Email'])))

        # Add Ward to master_clean 
        if 'Ward' not in df_master_clean:
            df_master_clean['Ward'] = ""
        
        # mWards = list(map(str.strip, list(df_master_clean['Ward'])))

        missing_indices = list()
        update_indices = list()
        data_to_add = list()                # Add missing index
        changes = list()
        for i in range(len(df_sheet_clean)):
            FirstName = df_sheet_clean.iloc[i]["First name"].strip().lower()
            LastName = df_sheet_clean.iloc[i]["Last name"].strip().lower()
            Ward = df_sheet_clean.iloc[i]["Last name"].strip().lower()
            # Searchs for missing data based on first name lastname
            temp = df_master_clean.index[
                                (df_master_clean['First Name'].str.strip().str.lower() == FirstName) & 
                                (df_master_clean['Last Name'].str.strip().str.lower() == LastName)
                            ].tolist()
            
            temp2 = df_master_clean.index[
                                (df_master_clean['First Name'].str.strip().str.lower() == FirstName) & 
                                (df_master_clean['Last Name'].str.strip().str.lower() == LastName) & 
                                (df_master_clean['Ward'].str.strip().str.lower() == LastName)
                            ].tolist()
            
            if len(temp)>0: # Data exists
                change = {"First Name":FirstName, "Last Name":LastName}
                j = temp[0]
                if df_sheet_clean.iloc[i]["Phone"].strip() != df_master_clean.iloc[j]["Phone"].strip():
                    change["Phone Past Value"] = df_master_clean.iloc[j]["Phone"]
                    try:
                        phone = phonenumbers.format_number(phonenumbers.parse(df_sheet_clean.iloc[i]["Phone"], 'US'),'US')
                    except:
                        phone = df_sheet_clean.iloc[i]["Phone"]
                    change["Phone New Value"] = phone
                    df_master_clean.iloc[j]["Phone"] = phone
                    
                if df_sheet_clean.iloc[i]["Email"].strip() != df_master_clean.iloc[j]["Email"]:
                    change["Email Past Value"] = df_master_clean.iloc[j]["Email"]
                    change["Email New Value"] = (parseaddr(df_sheet_clean.iloc[i]["Email"])[1] if parseaddr(df_sheet_clean.iloc[i]["Email"]) else "")
                    df_master_clean.iloc[j]["Email"] = (parseaddr(df_sheet_clean.iloc[i]["Email"])[1] if parseaddr(df_sheet_clean.iloc[i]["Email"]) else "")

                if df_master_clean.iloc[j]["Ward"] != df_sheet_clean.iloc[i]["I live in Cleveland Ward # "]:
                    change["Ward Past Value"] = df_master_clean.iloc[j]["Ward"]
                    change["Ward New Value"] = df_sheet_clean.iloc[i]["I live in Cleveland Ward # "]
                    df_master_clean.iloc[j]["Ward"] = df_sheet_clean.iloc[i]["I live in Cleveland Ward # "]
                
                if "Address" in df_sheet_clean:
                    if df_master_clean.iloc[j]["Address"] != df_sheet_clean.iloc[i]["Address"].strip():
                        df_master_clean.iloc[j]["Address"] = df_sheet_clean.iloc[i]["Address"]
                changes.append(change)
            else:
                data_to_add.append({
                        'First Name':df_sheet_clean.iloc[i][column_map['First Name']],
                        'Last Name':df_sheet_clean.iloc[i][column_map['Last Name']],
                        "Phone":df_sheet_clean.iloc[i]['Phone'],
                        "Email":df_sheet_clean.iloc[i]['Email'],
                        "Ward":df_sheet_clean.iloc[i]["I live in Cleveland Ward # "]
                    })
            
        
        pd.concat([df_master_clean,pd.DataFrame(data_to_add)],ignore_index=True).to_excel('master_clean.xlsx','PB CLE Master List')
        

        pd.DataFrame(changes).to_excel("master_clean_changes.xlsx")
        pd.DataFrame(data_to_add).to_excel("master_clean_data_added.xlsx")
        print('done')
         
        
        
        
            


