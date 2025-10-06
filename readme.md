# Vad scriptet gör:

• Tar en infil, med spelare på rader, och matcher på kolumner, samt två kolumner för villig att spela mer matcher, samt att vara målvakt.  
• Strikt max 1 målvakt per spelare <code>--gk-cap 1</code>  
• Standardkallelser: max 2 hemma och max 2 borta per spelare <code>--max-home-base</code> och <code>--max-away-base</code>  
• Kedja 3 (RESERV) skapas endast om exakt 4 spelare finns <code>--require-exact-reserve-four</code> som är på default. Du kan släppa detta krav med <code>--no-require-exact-reserve-four</code>  
• En extra infosträng på varje matchflik räknas ut: MÖJLIGA RESERVER (alla övriga tillgängliga spelare). Denna påverkar inte reservkallelser. Men är bra för att se vilka som är tillgängliga som reserver.    
• Huvudfliken får översiktskolumner: **Kallelser Hemma**, **Kallelser Borta**, **Kallelser Totalt**, **Reservkallelser**, samt **Målvaktsgånger**. För att se över att allt stämmer.   
• Målvaktsval prioriterar de som angett målvakt (kan stängas av med <code>--no-prefer-gk-volunteers</code>).  
• Matchkolumner identifieras automatiskt via titlar som innehåller (Hemma) eller (Borta).  
• “Bortavillig” spelare härleds från antalet svar på bortamatcher om den finns.

## Exempel på fler körningar
Med egen utdatafil och utan krav på exakt 4 i reservkedjan:  
<code>python p16_allocation.py P16-svar_blank.xlsx P16_resultat.xlsx --no-require-exact-reserve-four</code>

Sätt andra tak:  
<code>python p16_allocation.py P16-svar_blank.xlsx P16_resultat.xlsx --max-home-base 2 --max-away-base 2 --gk-cap 1</code>
