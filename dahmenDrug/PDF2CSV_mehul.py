"""
Spyder Editor
Created on Mon Feb 15 11:02:34 2023
@author: M03593 (Mehul Bafna)
"""

#Imported Modules
from tabula.io import read_pdf
import pandas as pd
import os
import time
import math
# For each page the table can be read with the following code

class extractpdf:
    
    def __init__(self,filepath,filename,columns):
        self.filepath = filepath
        self.filename = filename
        self.columns = columns
        
    @staticmethod 
    def encrypt(n):
        s=''
        for i in str(n):
            s+=chr(pow(int(i),2,65)+65)
        return s
    
    @staticmethod 
    def decrypt(s):
        n=''
        for i in s:
            n+=str(int(math.sqrt(ord(i)-65)))
        return int(n)
        
    def pdf2csv(self):
        
        #Creating dataframes of the pdf file for all pages
        dfs = read_pdf(self.filepath+self.filename, guess=False, pages = 'all',stream=True , encoding="latin", columns = self.columns)
        #To append dataframes from multiple pages
        df = pd.DataFrame(columns=[l+1 for l in range(17)])
        
        L = []
        #Renaming the columns
        for i in dfs:
            i.columns = [k+1 for k in range(len(i.columns))]
        
        #To obtain date and time information
        f = dfs[0][1][dfs[0][1]=='Untersuchung'].index.values[0]
        D = []
        T = []
        F = ['Laboratory Parameters','','Unit','Reference Scale','Response 1','Response 2','Response 3','Response 4','Response 5','Response 6']
        for g in range(4):
            if g==0:
                D.append(str(dfs[0].iloc[f-4,7])+str(dfs[0].iloc[f-4,8])+extractpdf.encrypt(str(dfs[0].iloc[f-4,10])+str(dfs[0].iloc[f-4,11])+str(dfs[0].iloc[f-4,12])))
                T.append('')
            else:
                D.append('')
                T.append('')
            
        if not pd.isna(dfs[0].iloc[f+2,4]) and not pd.isna(dfs[0].iloc[f+2,5]) and not pd.isna(dfs[0].iloc[f+3,4]):
            D.append(str(dfs[0].iloc[f+2,4])+str(dfs[0].iloc[f+2,5]))
            T.append(str(dfs[0].iloc[f+3,4]))
        else:
            D.append('')
            T.append('')
        if not pd.isna(dfs[0].iloc[f+2,6]) and not pd.isna(dfs[0].iloc[f+2,7]) and not pd.isna(dfs[0].iloc[f+3,6]) and not pd.isna(dfs[0].iloc[f+3,7]):
            D.append(str(dfs[0].iloc[f+2,6])+str(dfs[0].iloc[f+2,7]))
            T.append(str(dfs[0].iloc[f+3,6])+str(dfs[0].iloc[f+3,7]))
        else:
            D.append('')
            T.append('')
            L.append('N.A.')
        if not pd.isna(dfs[0].iloc[f+2,8]) and not pd.isna(dfs[0].iloc[f+2,9]) and not pd.isna(dfs[0].iloc[f+2,10]) and not pd.isna(dfs[0].iloc[f+3,8]):    
            D.append(str(dfs[0].iloc[f+2,8])+str(dfs[0].iloc[f+2,9])+str(dfs[0].iloc[f+2,10])[0:3])
            T.append(str(dfs[0].iloc[f+3,8]))
        else:
            D.append('')
            T.append('')
        if not pd.isna(dfs[0].iloc[f+2,10]) and not pd.isna(dfs[0].iloc[f+2,11]) and not pd.isna(dfs[0].iloc[f+2,12]) and not pd.isna(dfs[0].iloc[f+3,10]) and not pd.isna(dfs[0].iloc[f+3,11]):    
            D.append(str(dfs[0].iloc[f+2,10])[4]+str(dfs[0].iloc[f+2,11])+str(dfs[0].iloc[f+2,12]))
            T.append(str(dfs[0].iloc[f+3,10])+str(dfs[0].iloc[f+3,11]))
        else:
            D.append('')
            T.append('')
        if not pd.isna(dfs[0].iloc[f+2,13]) and not pd.isna(dfs[0].iloc[f+2,14]):
            D.append(str(dfs[0].iloc[f+2,13])+str(dfs[0].iloc[f+2,14]))
            T.append(str(dfs[0].iloc[f+3,13]))
        else:
            D.append('')
            T.append('')
        if not pd.isna(dfs[0].iloc[f+2,15]) and not pd.isna(dfs[0].iloc[f+2,16]):
            D.append(str(dfs[0].iloc[f+2,15])+str(dfs[0].iloc[f+2,16]))
            T.append(str(dfs[0].iloc[f+3,15])+str(dfs[0].iloc[f+3,16]))
        else:
            D.append('')
            T.append('')
        
        #To append dataframes from multiple pages
        for j in range(0,len(dfs)):
            cdf = pd.concat([df,dfs[j]],ignore_index=True,join='outer')
            df = cdf
        
        #Drop irrelevant columns
        for l in range(6,18):
            K = [7,9,12,14,16]
            if l not in K:
                cdf.drop(labels=l,axis=1,inplace=True)
        
        #Data preprocessing          
        cdf= cdf.reset_index(drop=True)
        cdf.columns= [k+1 for k in range(len(cdf.columns))]
        
        pat = ')'
        
        df1 = cdf.loc[cdf[1].str.endswith(pat,na=False)]
        df2 = df1.loc[df1[1].str.startswith('(',na=False)]
        df1 = df1[~df1.isin(df2)]
        df1 = df1.loc[df1[1].str.endswith(pat,na=False)]
        df1 = df1.fillna('N.A.')
        df1.loc[0] = D
        df1.loc[1] = T
        df1.loc[1.5] = F
        df1.sort_index(axis = 0,inplace=True)
        df1.reset_index(drop=True,inplace=True)
        
        for i in df1.columns[4:len(df1.columns)-1]:
            if str(df1.iloc[0,i])!='':
                pass
            else:
                df1[i+1].replace('N.A.','',inplace=True)

        #Returning the Exporting results to csv
        return df1.to_csv(filepath+'results.csv',encoding='latin',index=False,header=False) 

if __name__=='__main__':
    #Fetches the current username
    stime = time.time()
    filepath = f'C:\\Users\\{os.getlogin()}\\.spyder-py3\\'
    filename = 'ALTMP_C11_2.pdf'
    columns = [169,195,230,285,315,334,353,380,404,412,432,452,478,498,524,544]
    #columns = (150,200,250,300,350,400)
    obj = extractpdf(filepath,filename,columns)
    obj.pdf2csv()
    print("--- Python script takes %s seconds ---" % (time.time() - stime))