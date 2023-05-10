# R Script voor het aanmaken van Swing importbestanden in 'long' format (platte data)
# Tabblad 1 van de output bevat de data voor Swing
# Tabblad 2 van de output bevat de metadata voor Swing

# Leegmaken Global Environment
rm(list=ls())

# Laden van packages
library(tidyverse) # tidyverse package biedt basis functionaliteit
library(readxl) # voor read_excel()
library(openxlsx) # voor write.xlsx()
library(haven) # voor read_spss(), zap_label() en zap_labels()
library(labelled) # voor to_character()
library(fastDummies) # voor dummy_cols()

# Is deze groep kleiner dan het aangegeven aantal dan wordt deze waarde op missing (-99996) gezet.
# Minimum aantal respondenten per variabele
minimum_N <- 50

# Minimum aantal respondenten per cel/percentage
minimum_N_cel <- 0

# Een (zelfgekozen) mapnaam die wordt aangemaakt in de huidige working directory. Hier worden de resultaten opgeslagen.
mapnaam <- 'Output Resultaten'

# Aanmaken map 
dir.create(mapnaam)

# Definiëren van de bron die wordt weergegeven in de swing metadata
bron <- 'gmvo'

# Sheets uit indicatorenoverzicht inladen
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
data <- read_spss(file = file.choose(), 
                  col_select = c('Gemeentecode',
                                 indicatoren$indicator[indicatoren$dichotomiseren == 0],
                                 indicatoren$indicator[indicatoren$dichotomiseren == 1] %>% str_match('.+(?=_[0-9]+$)|.+') %>% as.vector() %>% unique(),
                                 indicatoren$uitsplitsing %>% .[is.na(.) == F] %>% str_split(', ') %>% unlist() %>% unique() %>% setdiff('geen'),
                                 indicatoren$gebiedsniveau %>% .[is.na(.) == F] %>% str_split(', ') %>% unlist() %>% unique(),
                                 indicatoren$weegfactor %>% unlist() %>% unique()))


# Variabelen dichotomiseren zoals aangegeven in het indicatorenoverzicht
if(any(indicatoren$dichotomiseren == 1)){
  data <- dummy_cols(data, 
                     select_columns = unique(str_extract(indicatoren$indicator[indicatoren$dichotomiseren == 1], '.+(?=_[0-9]+$)')),
                     ignore_na = T)
}


# Hercoderen van variabele met 8 = 'nvt' naar 0 zodat de percentages een weergave zijn van de totale groep.
# Zijn je variabelen al correct gehercodeerd dan kun je deze stap overslaan.
# Dit stukje code geeft een warning die kan worden genegeerd.
data <- data %>%
  mutate_at(c(data %>%
                select(indicatoren$indicator) %>%
                val_labels() %>%
                str_detect('[Nn][\\.]?[Vv][\\.]?[Tt]') %>%
                indicatoren$indicator[.]), list(~recode(., `8`= 0)))

# Unnesten van indicatorenoverzicht (uitsplitsingen en gebiedsniveaus)
analyseschema <- indicatoren %>%
  filter(filter == 0) %>%
  select(-filter) %>%
  mutate_at(.vars = c('uitsplitsing', 'gebiedsniveau'), str_split, ', ') %>%
  unnest(uitsplitsing) %>%
  unnest(gebiedsniveau) %>%
  mutate(uitsplitsing = ifelse(uitsplitsing == 'geen', NA, uitsplitsing))

# Functie aanmaken om gemiddelden te berekenen
compute_mean <- function(data, indicator, naam, omschrijving, swingnaam, uitsplitsing, gebiedsniveau, weegfactor, continue, jaar, afronding, informatie){
  
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

# Functie kan ook worden gebruikt om cijfers voor een enkele indicator te berekenen
compute_mean(data = data, 
     omschrijving = 'test',
     naam = 'var',
     swingnaam = 'swingnaam', 
     indicator = 'KLGGA207_1', 
     uitsplitsing = 'Totaal18_64', 
     gebiedsniveau = 'MIREB201', 
     weegfactor = 'ewCBSGGD', 
     continue = 0, 
     jaar = 2022,
     afronding = 1,
     informatie = 'pdf.pdf') %>%
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


# Resultaten opslaan. Deze file bevat alle berekende output(gemiddelden, N, N0 en N1).
write.xlsx(x = resultaten, file = paste0(mapnaam, '/', 'Resultaten.xlsx'))

# Aanmaken van swing databestand
swing_data <- resultaten %>%
  mutate(WAARDE = ifelse(N < minimum_N | N1 < minimum_N_cel | N0 < minimum_N_cel , -99996, weighted_mean),
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


