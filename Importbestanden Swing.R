# R Script voor het aanmaken van Swing importbestanden in 'long' format (platte data)

# Dit script genereert twee bestanden.
# Bestand 1 bevat alle resultaten (waaronder berekende gemiddelden, N, N1, N0)
# Bestand 2 bevat een swingimportbestand met twee tabbladen, waarbij:
  # Tabblad 1 de platte Swing data bevat
  # Tabblad 2 de metadata bevat


# 0. Voorbereiding --------------------------------------------------------

# Leegmaken Global Environment
rm(list=ls())


# 1. Laden benodigde packages --------------------------------------------

# Laden van packages
library(tidyverse) # tidyverse package biedt basis functionaliteit
library(readxl) # voor read_excel()
library(openxlsx) # voor write.xlsx()
library(haven) # voor read_spss(), zap_label() en zap_labels()
library(labelled) # voor to_character()
library(fastDummies) # voor dummy_cols()


# 2. Definieren minimum aantallen -----------------------------------------

# Is deze groep kleiner dan het aangegeven aantal dan wordt deze waarde op missing (-99996) gezet.
# De ingevulde waarden zijn gebaseerd op de landelijke afspraken. Indien gewenst kunnen de aantallen
# worden aangepast. Zo betekent een mininum N per cel van 0 dat er percentages uit kunnen komen die 
# 0% of 100% zijn binnen een bepaalde groep.

# Minimum aantal respondenten per variabele
minimum_N <- 50

# Minimum aantal respondenten per cel/percentage
minimum_N_cel <- 0


# 3. Definieren regionaam en -code ----------------------------------------

# Pas onderstaande naam en code aan naar de naam en code van je eigen regio.
regionaam <- 'GGD Limburg-Noord'
regiocode <- 23


# 4. Doelmap aanmaken ----------------------------------------

# Stel de working directory in op de map waar je de resultaten wilt opslaan. Het is het handigst als dit de map is waar ook het indicatorenoverzicht staat.
# Pas het deel tussen '' aan.
setwd('Q:/vul/hier/de/juiste/padnaam/in')

# Een (zelfgekozen) mapnaam die wordt aangemaakt in de hierboven opgegeven working directory. Hier worden de resultaten opgeslagen.
# Pas de naam tussen '' aan als je de map een andere naam wilt geven.
mapnaam <- 'Output Resultaten'

# Aanmaken map 
dir.create(mapnaam)


# 5. Bronnaam geven ----------------------------------------

# Definieren van de bron die wordt weergegeven in de swing metadata
# Pas de naam tussen '' aan naar de naam die de bron heeft in je eigen SWING dashboard/databank
bron <- 'bron swing'

# 6. Indicatorenoverzicht en data inladen ---------------------------------

# Sheets uit indicatorenoverzicht inladen
# Pas de naam tussen '' aan als je het indicatorenoverzicht een andere naam hebt gegeven.
indicatorenoverzicht <- 'Indicatorenoverzicht VO 2022.xlsx' %>%
  excel_sheets() %>%
  set_names() %>%
  map(read_xlsx, path = 'Indicatorenoverzicht VO 2022.xlsx')

# Indicatoren tabblad
indicatoren <- indicatorenoverzicht[['Indicatoren']] %>%
  filter(filter == 0)

# Uitsplitsingen tabblad
uitsplitsingen <- indicatorenoverzicht[['Uitsplitsingen']]

# Data laden
# Verzet de working directory zo nodig naar de map met de data. Pas het deel tussen '' aan.
setwd('Q:/vul/hier/de/juiste/padnaam/in')
# Bij het laden van de data worden met het col_select argument alleen de benodigde kolommen
# uit de data ingeladen. De gebiedsniveauvariabelen nederlandvar, regiovar en gemeentenvar
# worden in de stap hierna aangemaakt met behulp van de opgegeven regiocode, regionaam en de
# variabelen MIREB201 en Gemeentecode.
# Pas de naam tussen '' aan naar de naam van jullie landelijke databestand.
data <- read_spss('naamdatabestand.sav', 
                  col_select = c('MIREB201', 'Gemeentecode',
                                 indicatoren$indicator[indicatoren$dichotomiseren == 0],
                                 indicatoren$indicator[indicatoren$dichotomiseren == 1] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                 indicatoren$uitsplitsing %>% .[is.na(.) == F] %>% str_split(', ') %>% unlist() %>% unique() %>% setdiff('geen'),
                                 indicatoren$weegfactor %>% unlist() %>% unique()))


# 7. Databewerkingen uitvoeren --------------------------------------------

# Pas nederlandvar, regiovar en gemeentenvar aan naar de gebiedsniveaus die in je eigen SWING dashboard/databank staan
# Zorg ervoor dat de namen voor de gebiedsniveauvariabelen overeen komen met die in het indicatorenoverzicht.
data <-  data %>%
  mutate(totaal = 'totaal', # variabele aanmaken om het totaalgemiddelde te kunnen berekenen
         nederlandvar = 'Nederland',
         regiovar = ifelse(MIREB201 == regiocode, regionaam, NA), # variabele voor de regio aanmaken op basis van de eerder opgegeven regiocode en regionaam
         gemeentenvar = ifelse(MIREB201 == regiocode, to_character(Gemeentecode), NA))

# 8. Dichotomiseren -------------------------------------------------------

# Variabelen dichotomiseren zoals aangegeven in het indicatorenoverzicht
if(any(indicatoren$dichotomiseren == 1)){
  data <- dummy_cols(data, 
                     select_columns = unique(str_extract(indicatoren$indicator[indicatoren$dichotomiseren == 1], '.+(?=_[0-9]+$)')),
                     ignore_na = T)
}


# 9. Hercoderen -----------------------------------------------------------

# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep.
# Zijn je variabelen al correct gehercodeerd dan kun je deze stap overslaan.
# Dit stukje code geeft een warning die kan worden genegeerd.
data <- data %>%
  mutate_at(c(data %>%
                select(indicatoren$indicator) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                indicatoren$indicator[.]), list(~recode(., `8`= 0)))


# 10. Unnesten -----------------------------------------------------------

# Unnesten van indicatorenoverzicht (uitsplitsingen en gebiedsniveaus)
analyseschema <- indicatoren %>%
  filter(filter == 0) %>%
  select(-filter) %>%
  mutate_at(.vars = c('uitsplitsing', 'gebiedsniveau'), str_split, ', ') %>%
  unnest(uitsplitsing) %>%
  unnest(gebiedsniveau) %>%
  mutate(uitsplitsing = ifelse(uitsplitsing == 'geen', NA, uitsplitsing))


# 11. Gemiddeldes berekenen en opslaan -----------------------------------------------

# Functie compute_mean aanmaken om gemiddelden te berekenen
compute_mean <- function(data, indicator, naam = NA, omschrijving = NA, swingnaam = NA, uitsplitsing = NA, 
                         gebiedsniveau, weegfactor, continue = 0, jaar = NA, afronding = 1, informatie = NA){
  
  variabelen <- setdiff(c(uitsplitsing, gebiedsniveau), NA)
  
  data %>%
    select(all_of(c(indicator, variabelen, weegfactor))) %>%
    group_by(., across(all_of(variabelen))) %>%
    summarise_at(indicator, list(mean = ~mean(., na.rm = T), # berekenen gemiddelde
                                 weighted_mean = ~weighted.mean(x = ., w = .data[[weegfactor]], na.rm = T), # berekenen gewogen gemiddelde
                                 N = ~sum(!is.na(.)), # berekenen van de N
                                 N0 = ~length(which(.==0)), # berekenen van de N0 
                                 N1 = ~length(which(.==1)))) %>% # berekenen van de N1
    drop_na(any_of(variabelen)) %>%
    ungroup() %>%
    {if(is.na(uitsplitsing)) rename(., gebiedsniveau_code = 1) %>% mutate(., uitsplitsing_code = NA)
      else rename(., uitsplitsing_code = 1, gebiedsniveau_code = 2)} %>%
    filter(is.na(uitsplitsing_code) | uitsplitsing_code != 0) %>%
    mutate(gebiedsniveau_label = to_character(gebiedsniveau_code),
           uitsplitsing_label = to_character(uitsplitsing_code),
           indicator = indicator,
           naam = naam,
           omschrijving = omschrijving,
           swingnaam = swingnaam,
           uitsplitsing_indicator = uitsplitsing,
           gebiedsniveau_indicator = gebiedsniveau,
           weegfactor = weegfactor,
           continue = continue,
           jaar = jaar,
           afronding = afronding,
           informatie = informatie) %>%
    select(indicator, naam, omschrijving, swingnaam,
           uitsplitsing_indicator, uitsplitsing_code, uitsplitsing_label,
           gebiedsniveau_indicator, gebiedsniveau_code, gebiedsniveau_label, jaar, everything()) %>%
    zap_label() %>%
    zap_labels()
}

# Het volgende codeblokje hoeft niet te worden gerund, maar kan handig zijn. Zie de onderstaande uitleg.
# De compute_mean() functie die aangemaakt wordt in de vorige stap kun je gebruiken om de cijfers voor
# een combinatie van indicator, uitsplitsing en gebiedsniveau te berekenen. Mocht je bij het berekenen
# van alle cijfers later in het script vastlopen dan kun je functie gebruiken om erachter te komen waar 
# de fout zit. Stel je wilt de cijfers berekenen voor overgewicht (AGGWS204), uitgesplitst naar geslacht (AGGSA202) 
# voor elke GGD regio (MIREB201) dan vul je als volgt de compute_mean() functie in:
compute_mean(data = data,
             naam = 'Overgewicht',
             indicator = 'AGGWS204', 
             uitsplitsing = 'AGGSA202', # Wil je niet uitsplitsen laat deze regel dan weg.
             gebiedsniveau = 'MIREB201', 
             weegfactor = 'ewCBSGGD') %>%
  View()


# Berekeningen uitvoeren op basis van het analyseschema en alle output combineren
output <- analyseschema %>%
  select(-dichotomiseren) %>%
  pmap(compute_mean, data = data, .progress = T) %>% # het .progress = T argument kan een error geven, verwijder dit argument inclusief de komma die eraan voorafgaat
  bind_rows()


# Swingnamen en omschrijvingen aanpassen aan de hand van de opgegeven extensies in het tabblad uitsplitsingen
# en de resultaten wegschrijven naar .xlsx
resultaten <- output %>%
  mutate(swingnaam = paste0(swingnaam, 
                            replace_na(uitsplitsingen$swingnaam_extensie[match(paste(output$uitsplitsing_indicator, output$uitsplitsing_code),
                                                                               paste(uitsplitsingen$uitsplitsing_indicator, uitsplitsingen$uitsplitsing_code))],'')),
         naam = paste(naam, 
                              replace_na(uitsplitsingen$omschrijving_extensie[match(paste(output$uitsplitsing_indicator, output$uitsplitsing_code),
                                                                                    paste(uitsplitsingen$uitsplitsing_indicator, uitsplitsingen$uitsplitsing_code))],'')),
         omschrijving = paste(omschrijving, 
                              replace_na(uitsplitsingen$omschrijving_extensie[match(paste(output$uitsplitsing_indicator, output$uitsplitsing_code),
                                                                                    paste(uitsplitsingen$uitsplitsing_indicator, uitsplitsingen$uitsplitsing_code))],'')))


# Verzet de working directory zo nodig naar de map waar je de nieuwe map hebt aangemaakt (zie begin van r-script). Pas het deel tussen '' aan.
setwd('Q:/vul/hier/de/juiste/padnaam/in')
# Resultaten opslaan als .xlsx bestand. Dit bestand bevat alle berekende output (gemiddelden, N, N0 en N1).
write.xlsx(x = resultaten, file = paste0(mapnaam, '/', 'Resultaten.xlsx'))

# Aanmaken van swing databestand
swing_data <- resultaten %>%
  mutate(WAARDE = ifelse(N == 0 | N < minimum_N | N1 < minimum_N_cel | N0 < minimum_N_cel , -99996, weighted_mean),
         WAARDE = if_else(continue == 0 & WAARDE != -99996, WAARDE * 100, WAARDE)) %>%
  select(PERIODE = jaar, 
         GEBIEDSNIVEAU = gebiedsniveau_indicator, 
         GEBIED = gebiedsniveau_code, 
         NAAM = swingnaam, 
         WAARDE)

# Aanmaken van swing metadata
swing_metadata <- resultaten %>%
  select(Name = naam, Description = omschrijving, `Indicator code` = swingnaam, RoundOff = afronding, `More information` = informatie, continue) %>%
  distinct(Name, .keep_all = T) %>%
  mutate(Unit = ifelse(continue == 0,'p', 'g'),
         Source = bron,
         `Data type` = ifelse(continue == 0,'Percentage', 'Mean')) %>%
  select(-continue)

# Het opslaan van het swing import bestand als .xlsx bestand in de aangemaakte map (directory)
write.xlsx(list('data' = swing_data,
                'metadata' = swing_metadata), paste0(mapnaam, '/', "Swing import.xlsx"), rowNames = FALSE)


