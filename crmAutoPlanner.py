#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings("ignore")
pd.set_option('display.max_rows', None)
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import datetime as dt
from scipy import stats
import openpyxl
import calendar


# In[2]:

print("Request in process......")


# find difference between 2 dates in months

def month_diff(x,y):
    end = x.dt.to_period('M').view(dtype='int64')
    start = y.dt.to_period('M').view(dtype='int64')
    return end - start


# In[3]:


# read transfer order
path = r"C:\Users\ChanStephen\OneDrive - GC International AG\crm\\"
df = pd.read_excel(path+'Transfer Order by Product5_31_2022 5 54 35 PM.xlsx', header=4)
df.columns = [i.lower() for i in df.columns]
df.columns = df.columns.str.replace(' ','_')
df.rename(columns={'value_in_base_currency':'sales','sales_person_name':'sales_rep',
                   'sales_order_date':'date'}, inplace=True)
df.date = pd.to_datetime(df.date, format='%d-%m-%Y')
df.account= [' '.join(i.split()) for i in df.account]
df.product_code=[' '.join(i.split()) for i in df.product_code]

df.rename(columns={'class':'cltv'},inplace=True)

df = df[['date','account_code','account','product_code','product_name','sales_rep','sales','cltv','country']]


# ## clustering address in Sg account list

# In[4]:


# read account list

ac = pd.read_excel(path+'sea_account.xlsx')
ac.columns = ac.columns.str.lower()
ac.columns = ac.columns.str.replace(' ','_')
ac.rename(columns={'class':'cltv','sales_executive':'sales_rep','region':'country'}, inplace=True)
ac.postal_code= ac.postal_code.replace(' ',np.nan, regex=True)
ac1=ac[['account_code','account_name','country', 'city','postal_code','billing_address','cltv']]
ac1.country = [i.lower() for i in ac1.country]

sg = ac1[ac1.country=='singapore']
sg_pc_nan = sg[sg.postal_code.isna()]
sg = sg.dropna(subset=['postal_code'])


# In[5]:


# extract postal code from billing address

import re

for i, n in zip(sg_pc_nan.index, sg_pc_nan.billing_address):
    
    # initializing K 
    K = 6
    # using regex() to solve problem
    temp = re.search('\d{% s}'% K, n)
    res = (temp.group(0) if temp else '')
    sg_pc_nan.loc[i,'postal_code']=res


# In[6]:


sg_pc = sg_pc_nan
sg = sg.append(sg_pc)


# In[7]:


# In[8]:


dfr = pd.read_excel(path+'sg_account_no_addresses_6 June.xlsx')


# In[9]:


# dfr.head(2)


# In[10]:


# extract postal code from billing address

import re

for i, n in zip(dfr.index, dfr.postal_sec):
    
    # initializing K 
    K = 6
    # using regex() to solve problem
    temp = re.search('\d{% s}'% K, n)
    res = (temp.group(0) if temp else '')
    dfr.loc[i,'postal_code']=res


# ## combining 2 files

# In[11]:


for i, m in zip(sg.index, sg.account_code):
    for mm, gg in zip(dfr.account_code, dfr.postal_code):
        if m==mm:
            sg.loc[i,'postal_code']=gg


# In[12]:


# sg


# In[ ]:





# In[ ]:





# In[13]:


def classify(cty):
    # compute number of days from last purchase
    _ = df[df.country==cty]
    _['current_date']=pd.to_datetime('today')
    _['since_last']=month_diff(_.current_date,_.date)
    
    req = (_.since_last >= 6) & (_.since_last <= 24)
    _=_[req]
#     _=_[_.since_last >= 6]
    class_a = _[_.cltv=='A']
    class_b = _[_.cltv=='B']
    class_c = _[_.cltv=='C']
    return class_a, class_b, class_c


# ## Sg account list to categorise accounts into clusters

# In[14]:


sg.columns


# In[15]:


sg.postal_code = sg.postal_code.astype(str)

# create postal sector to be clustered

sg['postal_sec']=np.nan

for i, p in zip(sg.index,sg.postal_code):
    x = p[0:2]
    sg.loc[i,'postal_sec']=x


# In[16]:


# categorizing into districts


s1=['01','02','03','04','05','06']
s2=['07','08']
s3=['14','15','16']
s4=['09','10']
s5=['11','12','13']
s6=['17']
s7=['18','19']
s8=['20','21']
s9=['22','23']
s10=['24','25','26','27']
s11=['28','29','30']
s12=['31','32','33']
s13=['34','35','36','37']
s14=['38','39','40','41']
s15=['42','43','44','45']
s16=['46','47','48']
s17=['49','50','81']
s18=['51','52']
s19=['53','54','55','82']
s20=['56','57']
s21=['58','59']
s22=['60','61','62','63','64']
s23=['65','66','67','68']
s24=['69','70','71']
s25=['72','73']
s26=['77','78']
s27=['75','76']
s28=['79','80']


# In[17]:


# generate numbers for district
u=1
sn=[]
while u <=28:
    j='d'+ str(u)
    sn.append(j)
    u=int(u)
    u+=1


# In[18]:


s=[s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15,s16,s17,s18,s19,s20,s21,s22,s23,s24,s25,s26,s28,s28]


# In[19]:


# len(sn), len(s)


# In[20]:


# create new cluster column

sg['cluster']=np.nan

# clustering Sg account list

def cluster(s,sn):
    for i,n in zip(sg.index, sg.postal_sec):
        if n in s:
            sg.loc[i,'cluster']=sn
        else:
            continue
    return sg


# In[21]:


for i,j in zip(s,sn):
    sg1=cluster(i,j)


# In[22]:


# sg1.to_csv('sg_account_no_addresses.csv',index=False)


# In[23]:


# sg1.loc[17801,'billing_address']


# In[24]:


# sg1[sg1.billing_address=='Singapore, '].to_csv('sg_account_no_addresses.csv',index=False)


# In[25]:


sgfinal = sg1.dropna(subset=['cluster'])


# In[26]:


# sg account list

# sgfinal.info()


# In[27]:


# extract class A & B customers

cty='Singapore'
a,b,c=classify(cty)

ab = a.append(b)
abc = ab.append(c)


# In[28]:


abc.head()


# ## CRM sales opportunity

# In[29]:


abc_ = abc[['date','cltv','account_code','account','product_code','product_name','sales','sales_rep',]]
abc_=abc_.rename(columns={'date':'Date Of Last Purchase','product_name':'Product Name','sales_rep':'Sales Person Name',
                    'cltv':'Class','account':'Account','sales':'Sales Amount','product_code':'Product Code'})
abc_.to_csv('crm_sales_oppy.csv',index=False)


# In[30]:


ab.sales_rep.unique()


# In[31]:


ab.to_csv('potential_cust_list.csv',index=False)


# ## Find out all the weekdays of the month

# In[32]:


# yr=2022
def get_weekdays(mth, yr=2022):
    cal = calendar.monthcalendar(yr,mth)
    tot=[]
    for n in range(len(cal)):
        for i in range(0,5):
            a = cal[n][i]
            if a!=0:
                tot.append(a)
            else:
                continue
    return tot


# In[33]:


mth=8
yr=2022
calendar.monthcalendar(yr,mth)


# In[34]:


mth=8
wkdays = get_weekdays(mth,yr)


# In[35]:


# wkdays


# In[36]:


calendar.weekday(2022,8,21)


# In[37]:


class crm():
    def __init__(self,name, meeting, month, date, year, stime, etime,dayname, opp, desc, cocaller, cl, sales_rep):
        self.name = name
        self.meeting = meeting
        self.month = month
        self.date = date
        self.year = year
        self.stime = stime
        self.etime = etime
        self.dayname = dayname
        self.opp = opp
        self.desc = desc
        self.cocaller = cocaller
        self.cl=cl
        self.sales_rep = sales_rep
    
        
    def fdate(self):
        return "{}/{}/{}".format(self.month,self.date,  self.year)


# In[38]:


t1="10:00 AM"
t2="11:30 AM"
t3='01:00 PM'
t4="02:30 PM"
t5="04:00 PM"
t6="05:30 PM"
t7="07:00 PM"

timelist=[t1,t2,t3,t4,t5,t6,t7]


# In[39]:


# groupby account code fo transfer order

a1 = ab.groupby(['account_code','account','sales_rep']).agg(no_pdts=('product_name','count')).reset_index()
a1.no_pdts = [str(i)+" "+ "products" for i in a1.no_pdts]


# In[40]:


# attached cluster number to transfer order

a1['cluster']=np.nan
for i,a in zip(a1.index,a1.account_code):
    for aa, c in zip(sgfinal.account_code, sgfinal.cluster):
        if a==aa:
            a1.loc[i,'cluster']=c


# In[41]:


a2 = a1.dropna(subset=['cluster'])


# In[42]:


# a2


# ## sales leads template

# In[43]:


ab.groupby(['date','account','product_name','sales_rep']).agg(sales=('sales','sum')).reset_index()


# In[ ]:





# In[ ]:





# In[ ]:





# ## Iterate over days of the month

# In[44]:


# create newly improved function

def plan1(tt,rep, mth,cl, s=0):
    
    wkdays = get_weekdays(mth,yr)
    
    rep = a2[a2.sales_rep==rep]
    m = rep[rep.cluster==cl].reset_index(drop=True)
    if m.shape[0]<tt: #------------------------------------------------------------------------------
        tt=m.shape[0]
    k = m.shape[0]//tt #----------------------------------------------------------------------------
    
    account=[]
    meeting=[]
    date=[]
    start=[]
    end=[]
    dayname=[]
    oppy_pdt=[]
    description=[]
    cocaller=[]
    city=[]
    rep=[]
    e = s+k
    for n in wkdays[s:e]:
        for i in range(tt): #---------------------------------------------------------------
            c1=crm(m.loc[i,'account'], 'telemarketing', mth,n,2022,timelist[i],timelist[i+1],calendar.weekday(2022,mth,n),'Y', m.loc[i,'no_pdts'],'N',cl, 
                   m.loc[i,'sales_rep'])
            account.append(c1.name)
            meeting.append(c1.meeting)
            date.append(c1.fdate())
            start.append(c1.stime)
            end.append(c1.etime)
            dayname.append(c1.dayname)
            oppy_pdt.append(c1.opp)
            description.append(c1.desc)
            cocaller.append(c1.cocaller)
            city.append(c1.cl)
            rep.append(c1.sales_rep)
            
            if i==tt-1: #------------------------------------------------------------------------
                m=m.loc[tt::].reset_index(drop=True) #--------------------------------------------
            else:
                continue
            
            
    df=pd.DataFrame({'account':account,'meeting':meeting,'date':date,'start':start,'end':end,'day':dayname,'oppy_pdt?':oppy_pdt,
               'description':description,'Co-caller?':cocaller,'city':cl,'sales_rep':rep})
    
    return df, e


def cluster(rep):
    _ = a2[a2.sales_rep==rep]
    return _.cluster.unique()


def index2int(dff,x):
    i = dff[dff.account==x].index.tolist()
    return int((" ".join(map(str, i))))


def index2str(i):
    return str((" ".join(map(str, i))))

def firstrow(p):
    _=p.iloc[[0]].T
    return _.T


# ## run solution

# In[45]:


# rep='Jade Li' #-----------------------------------------------------------------------------------
# tt=6
# s=0
# mth=6
# tot=pd.DataFrame()
# for i in cluster(rep):
#     df, s = plan1(tt, rep,mth,i,s)
#     tot=tot.append(df)
# tot1=tot.reset_index(drop=True)
# tot2=tot1.copy()
# tot=tot.reset_index(drop=True)


# # In[46]:


# tot


# # In[47]:


# tot5=tot.set_index('account') #--------------------------------------------------------------------------


# ## transform into google calendar format

# In[48]:


def csv2google(dfff):
    dfff=dfff.reset_index()
    dfff.rename(columns={'account':'Subject','date':'Start Date','start':'Start Time','end':'End Time'},inplace=True)
    dfff=dfff[['Subject','Start Date','Start Time','End Time']]
    return dfff
    


# ## bug fixing

# In[49]:


# tot5=tot.set_index('account')


# In[50]:


# pref = pd.DataFrame({'account':['Tooth Matters Dental Surgery (Toa Payoh)','Dental Werks@Farrer','Asia Healthcare Dental Centre','Synergy Dental Group Pte Ltd'], 
#                      'pref_time': [t3,t2,t4,t1],'pref_day':[2,4,1,2],'city':['d12','d10','d9','d6']})

# # pref = pd.DataFrame({'account':['Tooth Matters Dental Surgery (Toa Payoh)'], 
# #                      'pref_time': [t3],'pref_day':[2],'city':['d12']}, index=[0])

# pref = pref.set_index('account')


# ## preference allocator

# In[51]:

# managing preferences
def crmplanner(pref, df):

    tot5=df.set_index('account')

    # check if district is in the list
    tot2=pd.DataFrame()
    for c in df.city.unique():
        for ct in pref.city.unique():
            if c==ct:
                aa = df[df.city==ct] # extract the dataframe with selected city district
                tot2=tot2.append(aa)

            else:
                continue

    aa = tot2.copy()


    # check if the account relating preferred time slot is in the original dataframe list of accounts
    tot3=[]
    for acc in pref.index.unique():
        for ac in aa.account.unique():
            if acc==ac:
                tot3.append(acc) # returns a list of accounts which tally with accounts with preferred time slot 
                len_list = len(tot3)

            else:
                continue


    for i in tot3:
        # extract customer and preferred T/S (df)
        a_pref = pref.loc[[i]].T
        a_pref = a_pref.T

        # extract name of customer from preferred list
        a_pref_name = index2str(a_pref.index) # a_pref.index is based on iterator 'i'
        a_pref_city = a_pref.city[0]
        a_pref_time = a_pref.pref_time[0]
        a_pref_day = a_pref.pref_day[0]

        # check to see whether a_pref_name, a_pref_city, a_pref_time are satisfied in a single customer
        # extract customer, city and start time from original list

        for s in tot5.index: # iterate over account names in original list
            if a_pref_name == s: # check whether pref account name is found in the original list
                a_orig = tot5.loc[[i]].T
                a_orig = a_orig.T
                a_orig_name = index2str(a_orig.index)
                a_orig_city = a_orig.city[0]
                a_orig_time = a_orig.start[0]
                a_orig_day = a_orig.day[0]



            else:
                continue

                # find out whether all conditions are satisfied for account name, city and preferred time
                # conditions satisfied
            if a_pref_name==a_orig_name and a_pref_city==a_orig_city and a_pref_time==a_orig_time and a_pref_day==a_orig_day:
                print(f"{a_pref_name}: Preferred day and time satisfied")
                print()
                continue # exit if all conditions are satisfied

            # conditions satisfied
            elif a_pref_name==a_orig_name and a_pref_city==a_orig_city and a_pref_time==a_orig_time:
                print(f"{a_pref_name}: Preferred time satisfied")
                print()
                continue # exit if all conditions are satisfied

            elif a_pref_name==a_orig_name and a_pref_city==a_orig_city:


                # extract original df with a_pref_city
                a_orig_df_city = tot5[tot5.city==a_orig_city]

                # check whether preferred day is in original crm list with selected district
                if a_pref_day in a_orig_df_city.day.tolist():


                    # extract df based on preferred day, preferred day available
                    a_pref_day_df = a_orig_df_city[a_orig_df_city.day==a_pref_day]

                    num_of_appoints = a_orig_df_city[a_orig_df_city.day==a_pref_day]


                    # iterate over account name and start time to find a match with preferred time
                    for i, d in zip(a_pref_day_df.index,a_pref_day_df.start):

                        if d==a_pref_time: # preferred day available

                            print()
                            target_acct_with_pref_time = a_pref_day_df.loc[[i]].T
                            target_acct_with_pref_time = target_acct_with_pref_time.T

                            # extract name of target acct with preferred time
                            target_acct_with_pref_time_name = index2str(target_acct_with_pref_time.index)

                            # switch account names between 'target_acct_with_pref_time_name' and 'a_pref_name'
                            b_target_acct_with_pref_time=target_acct_with_pref_time.rename                            (index={target_acct_with_pref_time_name:a_pref_name})
                            b_orig = a_orig.rename(index={a_pref_name:target_acct_with_pref_time_name})

                            # delete these 2 names from original crm list "tot5"
                            tot5 = tot5.drop([target_acct_with_pref_time_name,a_pref_name])

                            tot5=tot5.append(b_target_acct_with_pref_time)
                            tot5=tot5.append(b_orig)

                            print(f"{a_pref_name}: Preferred day and time satisfied")
                            print()
                        else:
                            a_orig.loc[a_pref_name,'start']=a_pref_time
                            tot5=tot5.drop(a_pref_name)
                            tot5=tot5.append(a_orig)


                    # when the preferred time is in the original list, just switch place with another account        
                elif a_pref_time in a_orig_df_city.start.tolist():

                    target_acct_with_pref_time = a_orig_df_city[a_orig_df_city.start==a_pref_time]
                    target_acct_with_pref_time = firstrow(target_acct_with_pref_time)

                    # extract name of target acct with preferred time
                    target_acct_with_pref_time_name = index2str(target_acct_with_pref_time.index)

                    # switch account names between 'target_acct_with_pref_time_name' and 'a_pref_name'
                    b_target_acct_with_pref_time=target_acct_with_pref_time.rename                    (index={target_acct_with_pref_time_name:a_pref_name})
                    b_orig = a_orig.rename(index={a_pref_name:target_acct_with_pref_time_name})

                    # delete these 2 names from original crm list "tot5"
                    tot5 = tot5.drop([target_acct_with_pref_time_name,a_pref_name])

                    tot5=tot5.append(b_target_acct_with_pref_time)
                    tot5=tot5.append(b_orig)

                    print(f"{a_pref_name}: Preferred time satisfied")
                    print()


                    # when preferred time is not in the original list, just insert the time
                else:
                    a_orig.loc[a_pref_name,'start']=a_pref_time
                    tot5=tot5.drop(a_pref_name)
                    tot5=tot5.append(a_orig)
                    
    return tot5


# In[ ]:





# In[52]:


# rep='Patrick Teo'
# tt=6
# s=0
# mth=7
# tot=pd.DataFrame()
# for i in cluster(rep):
#     df, s = plan1(tt, rep,mth,i,s)
#     tot=tot.append(df)
# tot=tot.reset_index(drop=True)


# In[53]:


# tot.to_csv('crm_calendar_v1.csv', index=False)


# In[54]:


pref = pd.DataFrame({'account':['Royce Dental Surgery (Bidadari)','CITIZEN DENTAL SURGERY','Implant Dental Centre'], 
                     'pref_time': [t1,t6,t5],'pref_day':[4,1,3],'city':['d13','d23','d10']})

# pref = pd.DataFrame({'account':['Tooth Matters Dental Surgery (Toa Payoh)'], 
#                      'pref_time': [t3],'pref_day':[2],'city':['d12']}, index=[0])

pref = pref.set_index('account')


# In[55]:


# pref


# In[56]:


# pref=pref
# df=tot
# temp=crmplanner(pref,df)#-----------------------------------------------


# In[57]:


# temp=temp.reset_index()#-----------------------------------------


# In[58]:


# temp.to_csv('crm_calendar_v1.csv', index=False)#------------------------------------------------


# In[59]:


pref.reset_index()
pref.to_csv('pref.csv', index=False)


# In[60]:


# csv2google(temp).to_csv('crm_calendar.csv',index=False) #-------------------------------------------------------


# In[61]:


rep='Patrick Teo'
tt=6
s=0
mth=7
tot=pd.DataFrame()
for i in cluster(rep):
    df, s = plan1(tt, rep,mth,i,s)
    tot=tot.append(df)
tot=tot.reset_index(drop=True)


# In[62]:


tot


# In[63]:


pref = pd.DataFrame({'account':['Royce Dental Surgery (Bidadari)','CITIZEN DENTAL SURGERY','Implant Dental Centre'], 
                     'pref_time': [t1,t6,t4],'pref_day':[2,4,1],'city':['d13','d23','d10']})

# pref = pd.DataFrame({'account':['Tooth Matters Dental Surgery (Toa Payoh)'], 
#                      'pref_time': [t3],'pref_day':[2],'city':['d12']}, index=[0])

pref = pref.set_index('account')


# In[64]:


pref


# In[65]:


pref=pref
df=tot
temp1 = crmplanner(pref,df)
temp1 = temp1.reset_index()
temp1.to_csv(path+'crm_calendar_v1.csv',index=False)

csv2google(temp1).to_csv(path+'crm_calendar.csv',index=False)


# In[66]:


print("Process completed.")


# In[ ]:




