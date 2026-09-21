# NTNUI-klubbklar
kode for å generere bestilling og faktura fra resultatene fra en google forms.

## Hvordan Fungerer det?
Programmet tar inn svarene fra google formsen i form av en excel fil. Tar så utgangspukt i `mal.xlsx` for å bestemme navnet på de ulike varene. og `prisliste.xlsx` for prisen på de ulike produktene. Oppdater disse med nyeste verdier om det skulle forekomme endringer.

## Hvordan kjøre:
1. Skriv inn navn på excel filen du får fra formsen i `main.py` i variablene `answer_filename`. 
2. Anngi navn/filplassering av faktura og bestillings excel filen. 
3. Skriv inn start og sluttdato for neste bestilling. 
4. Kjør så `main.py`
5. De to excel filene bør bli generet/oppdatert med alt du trenger.

## Opdatere programmet
Dersom formsen blir oppdatert er det nødvendig å gå gjennom `person_class.py` og oppdatere spørsmålene der. Om nye produkter blir lagt til. krever det  at det bli lagt til i både `person_class.py`, `mal.xlsx` og `prisliste.xlsx`.

## Forbedringer:
- Fjerne `prisliste.py`
- Skrive  `t-skorte` og ikke `t-skjorte (herre)`
- Legge til en excel fil med alle spørsmålene og hva de betyr for å gjøre det enklere å ekspandere/endre bestillingsskjema.
