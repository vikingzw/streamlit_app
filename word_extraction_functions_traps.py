import docx2txt
# import docx
import pandas as pd
import os
import numpy as np

def retrieve_NO(result):
    for i in range(len(result)-1):
        if (result[i] == 'Sample') & (result[i+1] == 'No.:'):
            k = 0
            while result[i+k] != 'Examiner:':
                k+=1

            return ' '.join(result[(i+2):(i+k)])
    return 'Not Found'


def retrieve_date(result):
    for i in range(len(result)-2):
        if (result[i] == 'Date') & (result[i+2] == 'sampling:'):
            k = 0
            while result[i+k] != 'Sampling':
                k+=1

            return ' '.join(result[(i+3):(i+k)])
    return 'Not Found'


def retrieve_sampling_place(result):
    for i in range(len(result)-1):
        if (result[i] == 'Sampling') & (result[i+1] == 'place:'):
            k = 0
            while result[i+k] != 'Date':
                k+=1

            return ' '.join(result[(i+2):(i+k)])
    return 'Not Found'

def retrieve_contamination_time(result):
    for i in range(len(result)-1):
        if (result[i] == 'Contamination') & (result[i+1] == 'time'):
            k = 0
            while result[i+k] != 'Evaluated':
                k+=1

            return ' '.join(result[(i+3):(i+k)])
    return 'Not Found'

def retrieve_largest_metallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'metallic'):
            k = 0
            while result[i+k] != 'Width':
                k+=1

            return ' '.join(result[(i+5):(i+k)])
    return 'Not Found'

def retrieve_largest_metallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'metallic'):
            k = 1
            while result[i+k] != 'Largest':
                k+=1

            return ' '.join(result[(i+8):(i+k)])
    return 'Not Found'


def retrieve_largest_nonmetallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'nonmetallic'):
            k = 1
            while result[i+k] != 'Width':
                k+=1

            return ' '.join(result[(i+5):(i+k)])
    return 'Not Found'


def retrieve_largest_nonmetallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'nonmetallic'):
            k = 1
            while result[i+k] != 'Fibre':
                k+=1

            return ' '.join(result[(i+8):(i+k)])
    return 'Not Found'

def retrieve_TCV(result):
    for i in range(len(result)):
        if result[i] == 'TCV:':
            k = 0
            while result[i+k] != 'MCV:':
                k+=1

            return ' '.join(result[(i+1):(i+k)])
    return 'Not Found'

def retrieve_MCV(result):
    for i in range(len(result)):
        if result[i] == 'MCV:':
            k = 0
            while result[i+k] != 'FCV:':
                k+=1

            return ' '.join(result[(i+1):(i+k)])
    return 'Not Found'


def retrieve_FCV(result):
    for i in range(len(result)):
        if result[i] == 'FCV:':
            k = 0
            while result[i+k] != 'Remarks:':
                k+=1

            return ' '.join(result[(i+1):(i+k)])
    return 'Not Found'

# def clean_data(df):
#     def fff(x):
#         if x == 'CoatingDate':
#             return 'Coating'
#         if x == 'PressingDate':
#             return 'Pressing'
#         if x == 'StackingDate':
#             return 'Stacking'
#         if x == 'Evaluated':
#             return 'Not Found'
#         if x == 'MCV:':
#             return 'Not Found'
#         if x == 'FCV:':
#             return 'Not Found'
#         if x == 'Remarks:':
#             return 'Not Found'
#         if x == 'Width':
#             return 'Not Found'
#         if x == 'nonmetallic':
#             return 'Not Found'
#         return x

#     df['Sampling Place'] = df['Sampling Place'].apply(lambda x: fff(x))
#     df['Contamination Time'] = df['Contamination Time'].apply(lambda x: fff(x))
#     df['TCV'] = df['TCV'].apply(lambda x: fff(x))
#     df['MCV'] = df['MCV'].apply(lambda x: fff(x))
#     df['FCV'] = df['FCV'].apply(lambda x: fff(x))
#     df['Largest Metallic Particle Length (micro m)'] = df['Largest Metallic Particle Length (micro m)'].apply(
#         lambda x: fff(x))
#     df['Largest Metallic Particle Width (micro m)'] = df['Largest Metallic Particle Width (micro m)'].apply(
#         lambda x: fff(x))
#     return df

def create_alert_col(col):
    alert_col = []
    for elem in col:
        if elem != 'Not Found':
            if elem<250:
                alert_col.append('Green')
            elif 250<=elem<500:
                alert_col.append('Yellow')
            else:
                alert_col.append('Red')
        else:
            alert_col.append('Particle not found')
    return alert_col

def get_traps_data(files):

    df = pd.DataFrame(columns=['Test No.', 'Sampling Date', 'Sampling Place',
                    'Contamination Time', 'TCV', 'MCV', 'FCV'])

    for i in range(len(files)):
        text = docx2txt.process(files[i])
        parsed_text = text.split()
        tmp_df = pd.DataFrame(data={'Test No.': retrieve_NO(parsed_text),
                        'Sampling Date': retrieve_date(parsed_text),
                        'Sampling Place': retrieve_sampling_place(parsed_text),
                        'Contamination Time': retrieve_contamination_time(parsed_text),
                        'TCV': retrieve_TCV(parsed_text),
                        'MCV': retrieve_MCV(parsed_text),
                        'FCV': retrieve_FCV(parsed_text),
                        'Largest Metallic Particle Length (micro m)': retrieve_largest_metallic_particle_length(parsed_text),
                        'Largest Metallic Particle Width (micro m)': retrieve_largest_metallic_particle_width(parsed_text),
                        'Largest Nonmetallic Particle Length (micro m)': retrieve_largest_nonmetallic_particle_length(parsed_text),
                        'Largest Nonmetallic Particle Width (micro m)': retrieve_largest_nonmetallic_particle_width(parsed_text)},index=[0])
        df = pd.concat([df,tmp_df],axis=0,ignore_index=True)

    df['Sampling Date'] = pd.to_datetime(df['Sampling Date'],format='%d/%m/%Y',errors='coerce').dt.date
    df.sort_values(by='Test No.',inplace=True)
    df['Sampling Date'].fillna('Not Found',inplace=True)
    df.iloc[:,3:] = df.iloc[:,3:].replace(to_replace='',value=99999.0) 
    df.iloc[:,3:] = df.iloc[:,3:].astype('float64',copy=True,errors='ignore')
    df.replace(to_replace=99999.0,value='Not Found',inplace=True) 
    df['Alert'] = create_alert_col(df['Largest Metallic Particle Width (micro m)'])
    # df.info()

    return df