# -*- coding: utf-8 -*-
"""
Created on Thu Apr 10 23:05:29 2025

@author: kim.arne
"""

#Hovedoppgave. Bruker standardoppgave.
#bruker Pandas til å behandle excel.

import pandas as pd
#df = data frame. Bruker konvensjonelt navngiving
import numpy as np
import matplotlib.pyplot as plt

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
#print ("Korteste samtale registrert er:", Kortest_samtale," Lengste samtale:", Lengste_samtale)
print ("Korteste samtale var:",u_dag[Kortest_samtale_Index],"kl", kl_slett[Kortest_samtale_Index],"varighet:",Kortest_samtale)
print ("Lengste samtale var:",u_dag[Lengste_samtale_index],"kl", kl_slett[Lengste_samtale_index],"varighet:", Lengste_samtale)

#del D Regn ut gjennomsnitt samtaletid alle hendvendesle
#Vurder iterativ løkker

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
print(f"Gjennomsnitt tid (hh:mm:ss): {timer:02}:{minuter:02}:{sek:02}")
    
#del e


#kl 08-10




#e) Supportvaktene i MORSE er delt inn i 2-timers bolker: kl 08-10, kl 10-12, kl 12-14 og kl
#14-16. Skriv et program som finner det totale antall henvendelser supportavdelingen mottok
#for hver av tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24. Resultatet visualiseres ved
#bruk av et sektordiagram (kakediagram).








#del f





