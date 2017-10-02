# -*- coding: utf-8 -*-
"""
Tekee tilaston aineiston virheistä.
Created on Mon Sep 25 14:36:25 2017

@author: tsokkonen
"""

import os
import openpyxl

SARAKE = "D"
ENSIMMAINEN_RIVI = 5
NIMEN_LOPPUOSA = "_kk_korjaus.xlsx"
TAULUKON_NIMI = "korjaus1"
JUURI = "O:\\hakemisto\\polku\\"

def hae_taulukko(vuosi, kuukausi):
    vaihda_hakemistoa(vuosi, kuukausi)
    nimi = rakenna_tyokirjan_nimi(vuosi, kuukausi)
    return avaa_kasiteltava_taulukko(nimi)

def vaihda_hakemistoa(vuosi, kuukausi):
    os.chdir(JUURI + vuosi + "\\" + vuosi + "_" + kuukausi)

def rakenna_tyokirjan_nimi(vuosi, kuukausi):
    return vuosi + "_" + kuukausi + NIMEN_LOPPUOSA

def avaa_kasiteltava_taulukko(nimi):
    try:
        tyokirja = openpyxl.load_workbook(nimi)
        return tyokirja.get_sheet_by_name(TAULUKON_NIMI)
    except IOError as e:
        print "I/O virhe({0}): {1}".format(e.errno, e.strerror) + " : " + nimi

def poimi_virheilmoitukset_taulukosta(taulukko):
    virheet = []
    viimeinen_rivi = taulukko.max_row
    taman_taulukon_rivit = range(ENSIMMAINEN_RIVI, viimeinen_rivi + 1)
    for rivi in taman_taulukon_rivit:
        teksti = poimi_teksti_solusta(rivi, taulukko)
        try:
            if teksti != None and "tehty" not in teksti:
                virheet.append(teksti)
        except TypeError:
            print "Tyyppivirhe"
    return virheet

def poimi_teksti_solusta(rivi, taulukko):
    return taulukko[SARAKE + str(rivi)].value

def laske_virheilmoitusten_frekvenssit(virheilmoitukset, frekvenssit):
    for k in virheilmoitukset:
        if frekvenssit.has_key(k):
            frekvenssit[k] += 1
        else:
            frekvenssit[k] = 1
    return frekvenssit

def laske_virheilmoitusten_lukumaarat(virheilmoitukset, lukumaarat):
    return lukumaarat.append(len(virheilmoitukset));

def main():
    # tilastoi virheet näille vuosille
    vuodet = ["2014", "2015", "2016"]
    
    # tee lista kuukausista
    kuukaudet = [str(x + 1).zfill(2) for x in range(12)]
    
    # rakenna lista taulukko-olioita iteroimalla vuosien ja kuukausien yli
    taulukot = [hae_taulukko(v, k) for v in vuodet for k in kuukaudet]
    
    # poimi virheilmoitukset taulukoista ja laske niiden frekvenssit
    frekvenssit = {}
    lukumaarat = []
    for taulukko in taulukot:
        virheilmoitukset = poimi_virheilmoitukset_taulukosta(taulukko)
        laske_virheilmoitusten_frekvenssit(virheilmoitukset, frekvenssit)
        laske_virheilmoitusten_lukumaarat(virheilmoitukset, lukumaarat)

    # tulosta frekvenssit
    print u'Virheilmoitus#Lukumäärä'
    for k,v in frekvenssit.items():
        print k, "#", v

    # tulosta lukumaarat
    print u'Virheilmoituksia kuukaudessa'
    for k in lukumaarat:
        print k

if __name__ == '__main__':
    main()
