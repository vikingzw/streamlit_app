import docx2txt
# import docx
import pandas as pd
import os
import numpy as np
from word_extraction_functions_stamps import *

def retrieve_component_type(result):
    for i in range(len(result)-2):
        if (result[i] == 'Component(s)') & (result[i+1] == 'type'):
            k = 0
            while result[i+k] != 'Sample':
                k+=1

            return ' '.join(result[(i+3):(i+k)])
    return 'Not Found'


def retrieve_component_no(result):
    for i in range(len(result)-2):
        if (result[i] == 'Component(s)') & (result[i+1] == 'No.:'):
            k = 0
            while result[i+k] != 'Particle':
                k+=1

            return ' '.join(result[(i+2):(i+k)])
    return 'Not Found'

def retrieve_sample_no(result):
    for i in range(len(result)-2):
        if (result[i] == 'Sample') & (result[i+1] == 'No.:'):
            k = 0
            while result[i+k] != 'Component(s)':
                k+=1

            return ' '.join(result[(i+2):(i+k)])
    return 'Not Found'

def retrieve_tested_area(result):
    for i in range(len(result)-2):
        if (result[i] == 'Tested') & (result[i+1] == 'Area:'):
            k = 0
            while result[i+k] != 'Date':
                k+=1

            return ' '.join(result[(i+2):(i+k)])
    return 'Not Found'

def retrieve_date_of_extraction(result):
    for i in range(len(result)-2):
        if (result[i] == 'Date') & (result[i+1] == 'of') & (result[i+2] == 'Extraction:'):
            k = 0
            while result[i+k] != 'Tested':
                k+=1

            return ' '.join(result[(i+3):(i+k)])
    return 'Not Found'

def retrieve_largest_metallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i:(i+5)] == ['Largest','metallic','particle','Length','[µm]:']):
            k = 0
            while result[i+k] != 'Width1':
                k+=1

            return ' '.join(result[(i+5):(i+k)])
    return 'Not Found'

def retrieve_largest_metallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i:(i+5)] == ['Largest','metallic','particle','Length','[µm]:']):
            k = 1
            while result[i+k] != 'Largest':
                k+=1
            return ' '.join(result[(i+8):(i+k)])
    return 'Not Found'

def retrieve_largest_nonmetallic_particle_length(result):
    for i in range(len(result)-2):
        if (result[i:(i+5)] == ['Largest','nonmetallic','particle2','Length','[µm]:']):
            k = 0
            while result[i+k] != 'Width1':
                k+=1

            return ' '.join(result[(i+5):(i+k)])
    return 'Not Found'

def retrieve_largest_nonmetallic_particle_width(result):
    for i in range(len(result)-2):
        if (result[i:(i+5)] == ['Largest','nonmetallic','particle2','Length','[µm]:']):
            k = 1
            while result[i+k] != 'Stretched':
                k+=1
            return ' '.join(result[(i+8):(i+k)])
    return 'Not Found'

def retrieve_particle_count_N(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'N':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '2000...3000':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_M(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'M':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '1500...2000':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_L(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'L':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '1000...1500':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_K(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'KL':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '600...1000':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_J(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'J':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '400…600':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   


def retrieve_particle_count_I(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'I':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '200…400':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_H(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'H':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '150…200':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_G(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'G':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '100...150':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   


def retrieve_particle_count_F(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'F':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '50...100':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_E(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'E':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '25...50':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_D(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'D':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '15...25':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_C(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'C':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '15...25':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'   

def retrieve_particle_count_C(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'C':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != '5...15':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'  
    
def retrieve_particle_count_B(result,type):

    for i in range(len(result)-2):
        if (result[i:(i+2)] == ['Detailed', 'results:']):
            k = 1
            #Find distance to correct category
            while result[i+k] != 'B':
                k+=1
            category_pos = i+k

            j = 0
            while result[category_pos+j] != 'CCC':
                j+=1
            return ' '.join(result[(category_pos+type):(category_pos+j-4+type)])
    return 'Not Found'    

def get_stamps_data(files):
    
    df = pd.DataFrame(columns=['Component type', 'Component No.', 'Sample No.','Tested Area', 'Date of extraction', 
                        'Largest metallic particle length [micro m]', 'Largest metallic particle width [micro m]',
                        'Largest nonmetallic particle length [micro m]','Largest nonmetallic particle width [micro m]',
                        'Metallic particle count N','Non-Metallic particle count N','Metallic particle count M','Non-Metallic particle count M',
                        'Metallic particle count L','Non-Metallic particle count L','Metallic particle count K','Non-Metallic particle count K',
                        'Metallic particle count J','Non-Metallic particle count J','Metallic particle count I','Non-Metallic particle count I',
                        'Metallic particle count H','Non-Metallic particle count H','Metallic particle count G','Non-Metallic particle count G',
                        'Metallic particle count F','Non-Metallic particle count F','Metallic particle count E','Non-Metallic particle count E',
                        'Metallic particle count D','Non-Metallic particle count D','Metallic particle count C','Non-Metallic particle count C',
                        'Metallic particle count B','Non-Metallic particle count B'])

    for i in range(len(files)):
        text = docx2txt.process(files[i])
        parsed_text = text.split()
        tmp_df = pd.DataFrame(data={'Component type': retrieve_component_type(parsed_text),
                        'Component No.': retrieve_component_no(parsed_text),
                        'Sample No.': retrieve_sample_no(parsed_text),
                        'Tested Area': retrieve_tested_area(parsed_text),
                        'Date of extraction': retrieve_date_of_extraction(parsed_text),
                        'Largest metallic particle length [micro m]': retrieve_largest_metallic_particle_length(parsed_text),
                        'Largest metallic particle width [micro m]': retrieve_largest_metallic_particle_width(parsed_text),
                        'Largest nonmetallic particle length [micro m]': retrieve_largest_nonmetallic_particle_length(parsed_text),
                        'Largest nonmetallic particle width [micro m]': retrieve_largest_nonmetallic_particle_width(parsed_text),
                        'Metallic particle count N': retrieve_particle_count_N(parsed_text,1),
                        'Non-Metallic particle count N': retrieve_particle_count_N(parsed_text,2),
                        'Metallic particle count M': retrieve_particle_count_M(parsed_text,1),
                        'Non-Metallic particle count M': retrieve_particle_count_M(parsed_text,2),
                        'Metallic particle count L': retrieve_particle_count_L(parsed_text,1),
                        'Non-Metallic particle count L': retrieve_particle_count_L(parsed_text,2),
                        'Metallic particle count K': retrieve_particle_count_K(parsed_text,1),
                        'Non-Metallic particle count K': retrieve_particle_count_K(parsed_text,2),
                        'Metallic particle count J': retrieve_particle_count_J(parsed_text,1),
                        'Non-Metallic particle count J': retrieve_particle_count_J(parsed_text,2),
                        'Metallic particle count I': retrieve_particle_count_I(parsed_text,1),
                        'Non-Metallic particle count I': retrieve_particle_count_I(parsed_text,2),
                        'Metallic particle count H': retrieve_particle_count_H(parsed_text,1),
                        'Non-Metallic particle count H': retrieve_particle_count_H(parsed_text,2),
                        'Metallic particle count G': retrieve_particle_count_G(parsed_text,1),
                        'Non-Metallic particle count G': retrieve_particle_count_G(parsed_text,2),
                        'Metallic particle count F': retrieve_particle_count_F(parsed_text,1),
                        'Non-Metallic particle count F': retrieve_particle_count_F(parsed_text,2),
                        'Metallic particle count E': retrieve_particle_count_E(parsed_text,1),
                        'Non-Metallic particle count E': retrieve_particle_count_E(parsed_text,2),
                        'Metallic particle count D': retrieve_particle_count_D(parsed_text,1),
                        'Non-Metallic particle count D': retrieve_particle_count_D(parsed_text,2),
                        'Metallic particle count C': retrieve_particle_count_C(parsed_text,1),
                        'Non-Metallic particle count C': retrieve_particle_count_C(parsed_text,2),
                        'Metallic particle count B': retrieve_particle_count_B(parsed_text,1),
                        'Non-Metallic particle count B': retrieve_particle_count_B(parsed_text,2)},index=[0])
        df = pd.concat([df,tmp_df],axis=0,ignore_index=True)

    df['Alert'] = np.where(df['Largest metallic particle width [micro m]'].astype(float) < 250,'Green',
        np.where(df['Largest metallic particle width [micro m]'].astype(float) < 500,'Yellow','Red'))

    df.replace(to_replace='',value=99999.0,inplace=True) 
    # for i,col in enumerate(df.columns):
    #     if i > 4:
    #         df.astype({col:'float64'})
    df = df.astype('float64',copy=True,errors='ignore')
    df['Date of extraction'] = pd.to_datetime(df['Date of extraction'],format='%d/%m/%Y').dt.date
    df.sort_values(by='Date of extraction',inplace=True)
    df.replace(to_replace=99999.0,value='Not Found',inplace=True) 

    return df