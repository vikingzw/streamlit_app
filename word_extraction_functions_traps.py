import docx2txt
# import docx
import pandas as pd
import os
import numpy as np

def retrieve_NO(result):
    index = []
    for i in range(len(result)-1):
        if (result[i] == 'Sample') & (result[i+1] == 'No.:'):
            return result[i+2]
    return 'Not Found'


def retrieve_date(result):
    index = []
    for i in range(len(result)-2):
        if (result[i] == 'Date') & (result[i+2] == 'sampling:'):
            return result[i+3]
    return 'Not Found'


def retrieve_sampling_place(result):
    index = []
    for i in range(len(result)-1):
        if (result[i] == 'Sampling') & (result[i+1] == 'place:'):
            return result[i+2] + result[i+3]
    return 'Not Found'


def retrieve_contamination_time(result):
    for i in range(len(result)-1):
        if (result[i] == 'Contamination') & (result[i+1] == 'time'):
            return result[i+3]
    return 'Not Found'


def retrieve_largest_metallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'metallic'):
            return result[i+5]
    return 'Not Found'


def retrieve_largest_metallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'metallic'):
            return result[i+8]
    return 'Not Found'


def retrieve_largest_nonmetallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'nonmetallic'):
            return result[i+5]
    return 'Not Found'


def retrieve_largest_nonmetallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i] == 'Largest') & (result[i+1] == 'nonmetallic'):
            return result[i+8]
    return 'Not Found'


def retrieve_TCV(result):
    for i in range(len(result)):
        if result[i] == 'TCV:':
            return result[i+1]
    return 'Not Found'


def retrieve_MCV(result):
    index = []
    for i in range(len(result)):
        if result[i] == 'MCV:':
            return result[i+1]
    return 'Not Found'


def retrieve_FCV(result):
    index = []
    for i in range(len(result)):
        if result[i] == 'FCV:':
            return result[i+1]
    return 'Not Found'

def clean_data(df):
    def fff(x):
        if x == 'CoatingDate':
            return 'Coating'
        if x == 'PressingDate':
            return 'Pressing'
        if x == 'StackingDate':
            return 'Stacking'
        if x == 'Evaluated':
            return 'Not Found'
        if x == 'MCV:':
            return 'Not Found'
        if x == 'FCV:':
            return 'Not Found'
        if x == 'Remarks:':
            return 'Not Found'
        if x == 'Width':
            return 'Not Found'
        if x == 'nonmetallic':
            return 'Not Found'
        return x

    df['Sampling Place'] = df['Sampling Place'].apply(lambda x: fff(x))
    df['Contamination Time'] = df['Contamination Time'].apply(lambda x: fff(x))
    df['TCV'] = df['TCV'].apply(lambda x: fff(x))
    df['MCV'] = df['MCV'].apply(lambda x: fff(x))
    df['FCV'] = df['FCV'].apply(lambda x: fff(x))
    df['Largest Metallic Particle Length (micro m)'] = df['Largest Metallic Particle Length (micro m)'].apply(
        lambda x: fff(x))
    df['Largest Metallic Particle Width (micro m)'] = df['Largest Metallic Particle Width (micro m)'].apply(
        lambda x: fff(x))
    return df

def get_traps_data(files):

    df = pd.DataFrame(columns=['Test N', 'Sampling Date', 'Sampling Place',
                    'Contamination Time', 'TCV', 'MCV', 'FCV'])

    for i in range(len(files)):
        text = docx2txt.process(files[i])
        parsed_text = text.split()
        tmp_df = pd.DataFrame(data={'Test N': retrieve_NO(parsed_text),
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

    df['Alert'] = np.where(df['Largest Metallic Particle Width (micro m)'].astype(float) < 250,'Green',
            np.where(df['Largest Metallic Particle Width (micro m)'].astype(float) < 500,'Yellow','Red'))

    
    df = clean_data(df)
    df = df.astype('float64',copy=True,errors='ignore')
    
    df['Sampling Date'] = pd.to_datetime(df['Sampling Date'],format='%d/%m/%Y').dt.date
    df['Sampling Date'] = df['Sampling Date']
    df.sort_values(by='Sampling Date',inplace=True)

    # df.info()

    return df