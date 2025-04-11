# -*- coding: utf-8 -*-
"""
Created on Thu Apr 10 23:05:29 2025

@author: kim.arne
testet at det virker i * Spyder version: 5.5.1  (conda)
* Python version: 3.12.4 64-bit
* Qt version: 5.15.2
* PyQt5 version: 5.15.10
* Operating System: Windows-10-10.0.19045-SP0
Betinger pandas installert
"""

#Hovedoppgave. Bruker standardoppgave.
#bruker Pandas til å behandle excel.

import pandas as pd
#df = data frame. Bruker konvensjonelt navngiving
import numpy as np
import matplotlib.pyplot as plt
import datetime #For del E

#Leser excel fil. Vurder feilhåndtering ved feil på filformat etter kvart
try:
 df = pd.read_excel('support_uke_24.xlsx')
except:
     print("Feil ved innlesing av data fra rekneark")

try:
 u_dag = df['Ukedag'].to_numpy()
 kl_slett = df['Klokkeslett'].to_numpy()
 varighet =df['Varighet'].to_numpy()
 score = df['Tilfredshet'].to_numpy()
except:
     print("Feil ved filformat")
#vurder om pythons innebyggede feilindikering gir mer utfyllende informasjon    
     
#delmål A oppfylt.

#Del B Antall hendvendelser for kvar av 5 ukedager. Visualiser med søylediagram
Antall_Mandag = (u_dag == 'Mandag').sum()
Antall_Tirsdag = (u_dag == 'Tirsdag').sum()
Antall_Onsdag = (u_dag == 'Onsdag').sum()
Antall_Torsdag = (u_dag == 'Torsdag').sum()
Antall_Fredag  = (u_dag == 'Fredag').sum()

#Plotter med matplotlib, Vurder bedre løsning etter hvert.
x = np.array(["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"])
y = np.array([Antall_Mandag , Antall_Tirsdag , Antall_Onsdag , Antall_Torsdag,Antall_Fredag])
plt.bar(x,y)
plt.show()

#Del C finn kortetste og lengste samtaletid logget i uke 24

#Bruker Pandas series max og min. Vurder om Pandas aggregate kunne vært mer hensiktsmessig men dette virker for nå
Kortest_samtale = df['Varighet'].max()
Kortest_samtale_Index = df['Varighet'].idxmax() #array indeks for maksverdi
Lengste_samtale = df['Varighet'].min()
Lengste_samtale_index =  df['Varighet'].idxmin() #array indeks for minverdi

#Finnes mer smartere og lettere måte å utføre på men dette virker ryddigst.
print ("Korteste samtale var:",u_dag[Kortest_samtale_Index],"kl", kl_slett[Kortest_samtale_Index],"varighet:",Kortest_samtale)
print ("Lengste samtale var:",u_dag[Lengste_samtale_index],"kl", kl_slett[Lengste_samtale_index],"varighet:", Lengste_samtale)

#del D Regn ut gjennomsnitt samtaletid alle hendvendesle

#totalt antall hendvendelser
antal_hendvendelser = len(u_dag)


#bruker pd mean for enkelt gjennomsnitt.
df["Varighet"] = pd.to_timedelta(df["Varighet"])
Varighet_snitt =df["Varighet"].mean()
tot_sek = Varighet_snitt.total_seconds()
timer = int(tot_sek // 3600)
minuter = int((tot_sek % 3600) // 60)
sek = int(tot_sek % 60)

print("Totalt antall hendvendelser:",antal_hendvendelser,"stk")    
print(f"Gjennomsnitt tid (tt:mm:ss): {timer:02}:{minuter:02}:{sek:02}") #filtrer tid til 2 heltall
    
#del e antall hendvendelser i blokker på 2 timer

#Vurder å gjør samenligningsfiltrering i egen funksjon
# ny formatering av dataramme til tid
df['Klokkeslett'] = pd.to_datetime(df['Klokkeslett'], format='%H:%M:%S').dt.time

# kl 08 til kl 10
start_tid = pd.to_datetime('08:00:00', format='%H:%M:%S').time()
slutt_tid = pd.to_datetime('10:00:00', format='%H:%M:%S').time()

#filtrer og teller på aktuel klokke
filtermaske = df[(df['Klokkeslett'] >= start_tid) & (df['Klokkeslett'] <= slutt_tid)].shape[0]
kl_08_10 = filtermaske

# kl 10 til kl 12
start_tid = pd.to_datetime('10:00:00', format='%H:%M:%S').time()
slutt_tid = pd.to_datetime('12:00:00', format='%H:%M:%S').time()
filtermaske = df[(df['Klokkeslett'] >= start_tid) & (df['Klokkeslett'] <= slutt_tid)].shape[0]
kl_10_12 = filtermaske

# kl 12 til kl 14
start_tid = pd.to_datetime('12:00:00', format='%H:%M:%S').time()
slutt_tid = pd.to_datetime('14:00:00', format='%H:%M:%S').time()
filtermaske = df[(df['Klokkeslett'] >= start_tid) & (df['Klokkeslett'] <= slutt_tid)].shape[0]
kl_12_14 = filtermaske

# kl 12 til kl 14
start_tid = pd.to_datetime('12:00:00', format='%H:%M:%S').time()
slutt_tid = pd.to_datetime('14:00:00', format='%H:%M:%S').time()
filtermaske = df[(df['Klokkeslett'] >= start_tid) & (df['Klokkeslett'] <= slutt_tid)].shape[0]
kl_12_14 = filtermaske

# kl 14 til kl 16
start_tid = pd.to_datetime('14:00:00', format='%H:%M:%S').time()
slutt_tid = pd.to_datetime('16:00:00', format='%H:%M:%S').time()
filtermaske = df[(df['Klokkeslett'] >= start_tid) & (df['Klokkeslett'] <= slutt_tid)].shape[0]
kl_14_16 = filtermaske


#Sekordiagram over data
#Plotter med matplotlib, Vurder bedre løsning etter hvert.
labels = '08:00-10:00', '10:00-12:00', '12:00-14:00', '14:00-16:00'
#y = np.array([kl_08_10 , kl_10_12 , kl_12_14 , kl_14_16])
størrelser = [kl_08_10 , kl_10_12 , kl_12_14 , kl_14_16]

fig, ax = plt.subplots()
#bruker standard farger inntil videre
ax.set_title("Henvendelser pr time uke 24")
ax.pie(størrelser, labels=labels, autopct='%1.1f%%')



"""del f beregne NPS 
1-6 = negtiv vil ikke anbefale
, 7-8 = nøytralt og 9-10 positivt ; vil trolig anbefale til andre

Beregn kvar gruppe i prosent av registrerte tilbakemeldinger. Trekk negative fra positive og multipliser med 100"""
Filt_score = df["Tilfredshet"].dropna() #Filtrerer bort tomme felter

n = len(Filt_score) #Totalt antall tilbakemeldinger

#definering av variblaer brukt i for loop. Vurder om global kan brukes. 
negativ = 0
nøytral = 0
positiv = 0

#Iterasjon av tilbakemeldinger
for i in Filt_score:
    #print(n)#testing
    if i >= 1 and i <= 6:
        negativ = negativ + 1
    if i >=8 and i <= 8:
        nøytral = nøytral + 1
    if i >=9 and i <= 10:
        positiv = positiv + 1

promoterere = (positiv / n)*100 #prosentvis andel
pasive = (nøytral / n)*100 #prosentvis andel
detraktører = (negativ / n)*100 #prosentvis andel

#print (promoterere,pasive,detraktører)

NPS = promoterere - detraktører
print ("Net Promoter Score er beregnet til:",int(NPS))


#Sekordiagram over data
#Plotter med matplotlib, Vurder bedre løsning etter hvert.
labels = 'Promotører', 'passive', 'detraktører'
størrelser = [promoterere , pasive , detraktører]

fig, ax = plt.subplots()
#bruker standard farger inntil videre
ax.set_title("Net Promoter Score uke 24")
ax.pie(størrelser, labels=labels, autopct='%1.1f%%')

