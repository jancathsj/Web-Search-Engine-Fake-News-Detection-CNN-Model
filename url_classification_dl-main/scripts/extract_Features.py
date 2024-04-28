from sys import platform
import data_creation_v3
import datetime
import math
import pandas as pd
import numpy as np
import whois
import stopit
from tqdm import tqdm
from interruptingcow import timeout
import os
print(os.getcwd())
os.chdir('../../../Data_Collection_03/02_Preprocessing/02_URL')

l = ['02_URL_Real.csv','02_URL_Fake.csv']
#l = ['DefacementSitesURLFiltered.csv','phishing_dataset.csv','Malware_dataset.csv','spam_dataset.csv','Benign_list_big_final.csv']
#l = ['phishing_dataset.csv']


emp = data_creation_v3.UrlFeaturizer("").run().keys()
A = pd.DataFrame(columns = emp)
t=[]
for j in l:
    print(j)
    d=pd.read_csv(j,header=None).to_numpy().flatten()
    for i in tqdm(d):
        
        try: 
            #with stopit.ThreadingTimeout(30) as to_ctx_mgr:
            #with stopit.Timeout(30.0, swallow_exc=False) as timeout_ctx:
            #with timeout(30, exception = RuntimeError):  
                #assert to_ctx_mgr.state == to_ctx_mgr.EXECUTING
                temp=data_creation_v3.UrlFeaturizer(i).run()
                temp["File"]=j.split(".")[0]
                t.append(temp)
                
                #timeout_ctx.state == timeout_ctx.EXECUTED
        #except TimeoutException:
        except RuntimeError: 
            pass 
A=A.append(t)
#os.chdir('../')
A.to_csv("03_features.csv")
