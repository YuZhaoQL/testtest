import pandas as pd
import  numpy as np
import os
title=['name','ks','kq','zy','sy','ps']

dat1=pd.read_excel('2ks.xlsx')
dat2=pd.read_excel('2kq.xlsx')
dat3=pd.read_excel('2zy.xlsx')
dat4=pd.read_excel('2sy.xlsx')


# print(dat1.shape)
# print(dat2.shape)
# print(dat3.shape)
# print(dat4.shape)
print(dat1.columns.values)
print(dat2.columns.values)
print(dat3.columns.values)
print(dat4.columns.values)
dat2N=dat2.loc[:,['name','w1']]
# print(dat2N.shape)
dat3N=dat3.loc[:,['name','w2']]
dat4N=dat4.loc[:,['name','w3']]
# print(dat3N)
# print(dat4N)


finallist=[]

for row in dat1.iterrows():
    name=row[1][1]
    vks=row[1][2]
    # print(name)
    # print(ks)
    result2=dat2N[dat2N['name']==name]
    result3=dat3N[dat3N['name']==name]
    result4=dat4N[dat4N['name']==name]

    if result2.shape[0]!=1:
        exit(11)
    else:
        # print(result2)
        vkq=float(result2.iloc[0,1])
        # print(vkq)

    if result3.shape[0] != 1:
        exit(11)
    else:
        vzy=float(result3.iloc[0,1])
        # print(vzy)

    if result4.shape[0] != 1:
        exit(11)
    else:
        vsy=float(result4.iloc[0,1])
        # print(vsy)
    vps=vkq*0.25+vzy*0.375+vsy*0.375-5
    tmplist=[]
    tmplist.append(name)
    tmplist.append(vks)
    tmplist.append(vkq)
    tmplist.append(vzy)
    tmplist.append(vsy)
    tmplist.append(vps)
    finallist.append(tmplist)


finaldat=pd.DataFrame(data=np.array(finallist),columns=title)
print(finaldat.shape)
finaldat.to_excel('summary2.xlsx')

