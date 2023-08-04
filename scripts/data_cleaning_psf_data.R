################################################################################
################################################################################

# Titre : Script de traitement des donnees des PSF
# Centre de Coordination des Operations d'urgence de sante Publique
# Auteur: Aureol Ngako, Epidemiologiste (GeorgeTown University)


################################################################################
################################################################################

# Chargement des packages utiles pour le traitement de la  base de donnee ----

    pacman::p_load( rio,        # Pour l'import et l'export des donnees
                    here,       # Pour la localisation des fichiers
                    tidyverse,  # Pour la gestion des donnees
                    writexl,    # Ecrire les fichiers excel avec plusieurs feuilles
                    readxl,     # Importer les fichiers excel avec plusieurs feuilles
                    lubridate   # Pour la gestion des dates
                    )

# Extraire les noms des  feuilles et  les  sauvegarder

  sheet_names  <- readxl::excel_sheets(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx")) 
  
## Selectionner les feuilles d'interet et ....
  
  sheet_names  <-  ifelse(str_detect(sheet_names,"S EPIDE") == "FALSE", NA_character_,sheet_names)
  
## Supprimer les feuilles inutiles 
  
  sheet_names <- na.omit(sheet_names) 
  
## Sites maritimes  ---------------------------------
  
  maritim <- c("YAGOUA",
               "PAK",
               "CAMPO",
               "PAL",
               "PAD",
               "IDENAU",
               "SOCAMBO",
               "TIKO",
               "EKONDO TITI",
               "GBITI",
               "ABOULOU",
               "LELE",
               "MANOKA")
  
## Ecrire une fonction qui va faire le nettoyage de chacune des bases de donnees ------
  
  clean_data <- function(data){
    
    data$jour <- str_sub(names(data)[1],12,13)
    data$annee <- str_extract(names(data)[1],
                              "2020|2021|2022|2023")
    data$mois_chaine <- str_extract(names(data)[1],
                                    "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
    
    data <-  data %>% 
      mutate(mois_num = case_when(
        mois_chaine == "JANVIER"  ~ 1,
        mois_chaine == "FEVRIER"  ~ 2,
        mois_chaine == "Février"  ~ 2,
        mois_chaine == "MARS"     ~ 3,
        mois_chaine == "AVRIL"    ~ 4,
        mois_chaine == "Avril"    ~ 4,
        mois_chaine == "MAI"      ~ 5,
        mois_chaine == "JUIN"     ~ 6,
        mois_chaine == "JUILLET"  ~ 7,
        mois_chaine == "AOUT"     ~ 8,
        mois_chaine == "SEPTEMBRE"~ 9,
        mois_chaine == "OCTOBRE"  ~ 10,
        mois_chaine == "NOVEMBRE" ~ 11,
        mois_chaine == "DECEMBRE" ~ 12),
        date = ymd(paste(annee,mois_num,jour,sep = "-")),
        type_poe    =  "Terrestre") %>% 
      # Selectionner et recoder les variables d'interet 
      
      select(date,
             type_poe,
             PSF                 = ...2 ,
             Promp               = ...3,
             Comp                = ...4,
             `Nb. Debarques`     = ...5,
             `Nb. Tests valides` = ...6,
             `Nb. Testes TDR`    = ...7,
             `Nb. Testes Positifs`= ...8,
             `Nb. Disposant d'un carnet de vaccination` = ...9) %>% 
      ## Supprimer les lignes inutiles
      drop_na(PSF) %>% 
      ## Convertir les variables d'interet en type numerique
      mutate(across(.cols = 4:10,.fns = as.numeric)) %>% 
      ## Replacer les variables numeriques vides par zero
      mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
      filter(!PSF %in% maritim) 
    
    
    data
  }
  
### Fonction pour les autres fichiers (Terrestre) ---------------------------------

clean_data2 <- function(data){
    
    data$jour <- str_sub(names(data)[1],12,13)
    data$annee <- str_extract(names(data)[1],
                              "2020|2021|2022|2023")
    data$mois_chaine <- str_extract(names(data)[1],
                                    "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
    
    data <-  data %>% 
      mutate(mois_num = case_when(
        mois_chaine == "JANVIER"  ~ 1,
        mois_chaine == "FEVRIER"  ~ 2,
        mois_chaine == "Février"  ~ 2,
        mois_chaine == "MARS"     ~ 3,
        mois_chaine == "AVRIL"    ~ 4,
        mois_chaine == "Avril"    ~ 4,
        mois_chaine == "MAI"      ~ 5,
        mois_chaine == "JUIN"     ~ 6,
        mois_chaine == "JUILLET"  ~ 7,
        mois_chaine == "AOUT"     ~ 8,
        mois_chaine == "SEPTEMBRE"~ 9,
        mois_chaine == "OCTOBRE"  ~ 10,
        mois_chaine == "NOVEMBRE" ~ 11,
        mois_chaine == "DECEMBRE" ~ 12),
        date = ymd(paste(annee,mois_num,jour,sep = "-")),
        type_poe    =  "Terrestre") %>% 
      # Selectionner et recoder les variables d'interet 
      
      select(date,
             type_poe,
             PSF                 = contains("DONNEES"),
             Promp               = ...2,
             Comp                = ...3,
             `Nb. Debarques`     = ...4,
             `Nb. Tests valides` = ...5,
             `Nb. Testes TDR`    = ...6,
             `Nb. Testes Positifs`= ...7,
             `Nb. Disposant d'un carnet de vaccination` = ...8) %>% 
      ## Supprimer les lignes inutiles
      drop_na(PSF) %>% 
      ## Convertir les variables d'interet en type numerique
      mutate(across(.cols = Promp:`Nb. Disposant d'un carnet de vaccination`,.fns = as.numeric)) %>% 
      ## Replacer les variables numeriques vides par zero
      mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
      # Enlever la lignes avec PSF
      filter(!PSF %in% "PSF") %>% 
      filter(!PSF %in% maritim)
    
    
    data 
  }
  
# Fichier des donnees terrestres  ----------------------------------------------
## Fichier 1 -----------------------------------------
combined_terrestre <- sheet_names %>% 
             purrr::set_names() %>% 
               map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "A1:I32")) 
  
### Base de donnees nettoye ----------------------------------------------------
  
  test_rev <- combined_terrestre %>%  map(.,
                                          .f = ~ clean_data(.x)) %>% 
                                          bind_rows(.id = "Semaine Epi")
                
## Fichier 2  ------------------------------------------------------------------
  
combined_terrestre2 <- sheet_names %>% 
                        purrr::set_names() %>% 
                         map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "J1:Q32"))

       `
  test_rev1 <- combined_terrestre2 %>%  map(.,
                                          .f = ~ clean_data2(.x)) %>% 
                                           bind_rows(.id = "Semaine Epi")
 
 
 ## Fichier 3 ------------------------
 
 combined_terrestre3 <- sheet_names %>% 
   purrr::set_names() %>% 
   map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "R1:Y32"))
  
  
  test_rev2 <- combined_terrestre3 %>%  map(.,
                                            .f = ~ clean_data2(.x)) %>% 
                                             bind_rows(.id = "Semaine Epi")
  
# Fichier 4 ---------------------------
  
  combined_terrestre4 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "Z1:AG32"))

    
  test_rev3 <- combined_terrestre4 %>%  map(.,
                                            .f = ~ clean_data2(.x)) %>% 
                                            bind_rows(.id = "Semaine Epi")
  
# Fichier 5 ------------------------------
  
  combined_terrestre5 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AH1:AO32"))
  
  
  test_rev4 <- combined_terrestre5 %>%  map(.,
                                            .f = ~ clean_data2(.x)) %>% 
                                            bind_rows(.id = "Semaine Epi")
# Fichier 6 -------------------------------
  
  combined_terrestre6 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AP1:AW32"))
  
  test_rev5 <- combined_terrestre6 %>%  map(.,
                                            .f = ~ clean_data2(.x)) %>% 
                                             bind_rows(.id = "Semaine Epi")

# Fichier 7 -------------------------------
  
  combined_terrestre7 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AX1:BE32"))
  
  
  test_rev6 <- combined_terrestre7 %>%  map(.,
                                            .f = ~ clean_data2(.x)) %>% 
                                             bind_rows(.id = "Semaine Epi")
## Donnees PSF Terrestres combine
  
  removed1 <- c("0","A.I.G","PSF","PSF MARITIMES")
  
  combined_psf_terrestres <- union_all(test_rev,test_rev1) %>% 
                             union_all(test_rev2) %>% 
                             union_all(test_rev3) %>% 
                             union_all(test_rev4) %>% 
                             union_all(test_rev5) %>% 
                             union_all(test_rev6)
  
  combined_psf_terrestres <- combined_psf_terrestres %>% 
                              filter(!PSF %in% removed1)
  
  #Exporter des data PSF terrestres ---------------
  
  export(combined_psf_terrestres,
          file = here("output","psf_terrestre.xlsx"))
  
## Fichier des donnees aeriennes  ---------------------------------------------- 
  
### Fonction pour selectioner la plage de data d'interet ------
  
  plage <- function(data){
    data <- data[20:55,]
  }  
### Fonction de traitement des data maritimes -----------------
  
  clean_data_maritime <- function(data){
      
      data$jour <- str_sub(names(data)[1],12,13)
      data$annee <- str_extract(names(data)[1],
                                "2020|2021|2022|2023")
      data$mois_chaine <- str_extract(names(data)[1],
                                      "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
      
      data <-  data %>% 
        mutate(mois_num = case_when(
          mois_chaine == "JANVIER"  ~ 1,
          mois_chaine == "FEVRIER"  ~ 2,
          mois_chaine == "Février"  ~ 2,
          mois_chaine == "MARS"     ~ 3,
          mois_chaine == "AVRIL"    ~ 4,
          mois_chaine == "Avril"    ~ 4,
          mois_chaine == "MAI"      ~ 5,
          mois_chaine == "JUIN"     ~ 6,
          mois_chaine == "JUILLET"  ~ 7,
          mois_chaine == "AOUT"     ~ 8,
          mois_chaine == "SEPTEMBRE"~ 9,
          mois_chaine == "OCTOBRE"  ~ 10,
          mois_chaine == "NOVEMBRE" ~ 11,
          mois_chaine == "DECEMBRE" ~ 12),
          date = ymd(paste(annee,mois_num,jour,sep = "-")),
          type_poe    =  "Maritime") %>% 
        # Selectionner et recoder les variables d'interet 
        
        select(date,
               type_poe,
               PSF                 = ...2 ,
               Promp               = ...3,
               Comp                = ...4,
               `Nb. Debarques`     = ...5,
               `Nb. Tests valides` = ...6,
               `Nb. Testes TDR`    = ...7,
               `Nb. Testes Positifs`= ...8,
               `Nb. Disposant d'un carnet de vaccination` = ...9) %>% 
        ## Supprimer les lignes inutiles
        drop_na(PSF) %>% 
        ## Convertir les variables d'interet en type numerique
        mutate(across(.cols = 4:10,.fns = as.numeric)) %>% 
        ## Replacer les variables numeriques vides par zero
        mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
        filter(PSF %in% maritim)
      
      data
    }
  
  clean_data_maritime2 <- function(data){
                           
      data$jour <- str_sub(names(data)[1],12,13)
      data$annee <- str_extract(names(data)[1],
                                "2020|2021|2022|2023")
      data$mois_chaine <- str_extract(names(data)[1],
                                      "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
      
       removed <- c("0","NSI","AID","AIG","PSF MARITIMES","PSF AERIENS")
       
      data <-  data %>% 
        mutate(mois_num = case_when(
          mois_chaine == "JANVIER"  ~ 1,
          mois_chaine == "FEVRIER"  ~ 2,
          mois_chaine == "Février"  ~ 2,
          mois_chaine == "MARS"     ~ 3,
          mois_chaine == "AVRIL"    ~ 4,
          mois_chaine == "Avril"    ~ 4,
          mois_chaine == "MAI"      ~ 5,
          mois_chaine == "JUIN"     ~ 6,
          mois_chaine == "JUILLET"  ~ 7,
          mois_chaine == "AOUT"     ~ 8,
          mois_chaine == "SEPTEMBRE"~ 9,
          mois_chaine == "OCTOBRE"  ~ 10,
          mois_chaine == "NOVEMBRE" ~ 11,
          mois_chaine == "DECEMBRE" ~ 12),
          date = ymd(paste(annee,mois_num,jour,sep = "-")),
          type_poe    =  "Maritime") %>% 
        # Selectionner et recoder les variables d'interet 
        
        select(date,
               type_poe,
               PSF                 = contains("DONNEES"),
               Promp               = ...2,
               Comp                = ...3,
               `Nb. Debarques`     = ...4,
               `Nb. Tests valides` = ...5,
               `Nb. Testes TDR`    = ...6,
               `Nb. Testes Positifs`= ...7,
               `Nb. Disposant d'un carnet de vaccination` = ...8) %>% 
        ## Supprimer les lignes inutiles
        drop_na(PSF) %>% 
        ## Convertir les variables d'interet en type numerique
        mutate(across(.cols = Promp:`Nb. Disposant d'un carnet de vaccination`,.fns = as.numeric)) %>% 
        ## Replacer les variables numeriques vides par zero
        mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
        # Enlever la lignes avec PSF
        filter(!PSF %in% removed) %>% 
        filter(PSF %in% maritim)
      
      
      data        
      
    }
  
## Fichier data maritime 1 ------------------------------------
  
  maritime <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "A1:I60")) %>% 
    map(.f = ~ plage(.x))
  
  
  test_maritime <- maritime %>%  map(.,
                                    .f = ~ clean_data_maritime(.x)) %>% 
                                    bind_rows(.id = "Semaine Epi")
  
  export(test_maritime, file = here("data","testmaritime.xlsx"))
  
## Fichier Maritime 2 -----------------------------------------------
  
  maritime1 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "J1:Q60")) %>% 
    map(.f = ~ plage(.x))
 
  test_maritime1 <- maritime1 %>%  map(.,
                                     .f = ~ clean_data_maritime2(.x)) %>% 
                                      bind_rows(.id = "Semaine Epi")
## Fichier Maritime 3 --------------------------------------------
  
   maritime2 <- sheet_names %>% 
                  purrr::set_names() %>% 
                  map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "R1:Y60")) %>% 
                  map(.f = ~ plage(.x)) 
  
  
  test_maritime2 <- maritime2 %>%  map(.,
                                       .f = ~ clean_data_maritime2(.x)) %>% 
                                              bind_rows(.id = "Semaine Epi")
  
## Fichier Maritime 4 -----------------------------------------------------
  
  maritime3 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "Z1:AG60")) %>% 
    map(.f = ~ plage(.x))
  
  test_maritime3 <- maritime3 %>%  map(.,
                                       .f = ~ clean_data_maritime2(.x)) %>% 
                                        bind_rows(.id = "Semaine Epi")
## Fichier Maritime 5 ---------------------------------------------------------
  
  maritime4 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AH1:AO60")) %>% 
    map(.f = ~ plage(.x))
  
  test_maritime4 <- maritime4 %>%  map(.,
                                       .f = ~ clean_data_maritime2(.x)) %>% 
                                       bind_rows(.id = "Semaine Epi")
## Fichier data PSF maritime 6 ------------------------------------------------------  
  
  maritime5 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AP1:AW60")) %>% 
    map(.f = ~ plage(.x))
  
  test_maritime5 <- maritime5 %>%  map(.,
                                       .f = ~ clean_data_maritime2(.x)) %>% 
                                        bind_rows(.id = "Semaine Epi")
## Fichier data PSF maritime 7 ------------------------------------------------
  
  maritime6 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AX1:BE60")) %>% 
    map(.f = ~ plage(.x))
  
  test_maritime6 <- maritime6 %>%  map(.,
                                       .f = ~ clean_data_maritime2(.x)) %>% 
                                        bind_rows(.id = "Semaine Epi")
## Combiner les fichiers maritimes ----------------------------------------------
  
  combined_psf_maritimes <- union_all(test_maritime,test_maritime1) %>% 
                            union_all(test_maritime2) %>% 
                            union_all(test_maritime3) %>% 
                            union_all(test_maritime4) %>% 
                            union_all(test_maritime5) %>% 
                            union_all(test_maritime6)
### Export file maritime
  
  export(combined_psf_maritimes,
         file = here("output","psf_maritime.xlsx"))
  
# Extraction des data de PSF aeriens -------------------------------------------
  
  aerien1 <- c("NSI","AID","AIG","AIMS")
  
## Fonction extraction PSF aeriens ---------------------------------------------
  
  clean_data_aerien <- function(data){
    
    data$jour <- str_sub(names(data)[1],12,13)
    data$annee <- str_extract(names(data)[1],
                              "2020|2021|2022|2023")
    data$mois_chaine <- str_extract(names(data)[1],
                                    "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
    
    aerien1 <- c("NSI","AID","AIG","AIMS")
    
    
   data <-  data %>% 
      mutate(mois_num = case_when(
        mois_chaine == "JANVIER"  ~ 1,
        mois_chaine == "FEVRIER"  ~ 2,
        mois_chaine == "Février"  ~ 2,
        mois_chaine == "MARS"     ~ 3,
        mois_chaine == "AVRIL"    ~ 4,
        mois_chaine == "Avril"    ~ 4,
        mois_chaine == "MAI"      ~ 5,
        mois_chaine == "JUIN"     ~ 6,
        mois_chaine == "JUILLET"  ~ 7,
        mois_chaine == "AOUT"     ~ 8,
        mois_chaine == "SEPTEMBRE"~ 9,
        mois_chaine == "OCTOBRE"  ~ 10,
        mois_chaine == "NOVEMBRE" ~ 11,
        mois_chaine == "DECEMBRE" ~ 12),
        date = ymd(paste(annee,mois_num,jour,sep = "-")),
        type_poe    =  "Aerien") %>% 
      # Selectionner et recoder les variables d'interet 
      
      select(date,
             type_poe,
             PSF                 = ...2,
             Promp               = ...3,
             Comp                = ...4,
             `Nb. Debarques`     = ...5,
             `Nb. Tests valides` = ...6,
             `Nb. Testes TDR`    = ...7,
             `Nb. Testes Positifs`= ...8,
             `Nb. Disposant d'un carnet de vaccination` = ...9) %>% 
      ## Supprimer les lignes inutiles
      drop_na(PSF) %>% 
      ## Convertir les variables d'interet en type numerique
      mutate(across(.cols = 4:10,.fns = as.numeric)) %>% 
      ## Replacer les variables numeriques vides par zero
      mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
      filter(PSF %in% aerien1)
    
    data
  }
  
  clean_data_aerien2 <- function(data){
    
    data$jour <- str_sub(names(data)[1],12,13)
    data$annee <- str_extract(names(data)[1],
                              "2020|2021|2022|2023")
    data$mois_chaine <- str_extract(names(data)[1],
                                    "JANVIER|FEVRIER|Février|MARS|AVRIL|Avril|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE")
    
    data <-  data %>% 
      mutate(mois_num = case_when(
        mois_chaine == "JANVIER"  ~ 1,
        mois_chaine == "FEVRIER"  ~ 2,
        mois_chaine == "Février"  ~ 2,
        mois_chaine == "MARS"     ~ 3,
        mois_chaine == "AVRIL"    ~ 4,
        mois_chaine == "Avril"    ~ 4,
        mois_chaine == "MAI"      ~ 5,
        mois_chaine == "JUIN"     ~ 6,
        mois_chaine == "JUILLET"  ~ 7,
        mois_chaine == "AOUT"     ~ 8,
        mois_chaine == "SEPTEMBRE"~ 9,
        mois_chaine == "OCTOBRE"  ~ 10,
        mois_chaine == "NOVEMBRE" ~ 11,
        mois_chaine == "DECEMBRE" ~ 12),
        date = ymd(paste(annee,mois_num,jour,sep = "-")),
        type_poe    =  "Aerien") %>% 
      # Selectionner et recoder les variables d'interet 
      
      select(date,
             type_poe,
             PSF                 = contains("DONNEES"),
             Promp               = ...2,
             Comp                = ...3,
             `Nb. Debarques`     = ...4,
             `Nb. Tests valides` = ...5,
             `Nb. Testes TDR`    = ...6,
             `Nb. Testes Positifs`= ...7,
             `Nb. Disposant d'un carnet de vaccination` = ...8) %>% 
      ## Supprimer les lignes inutiles
      drop_na(PSF) %>% 
      ## Convertir les variables d'interet en type numerique
      mutate(across(.cols = Promp:`Nb. Disposant d'un carnet de vaccination`,.fns = as.numeric)) %>% 
      ## Replacer les variables numeriques vides par zero
      mutate_if(is.numeric,~ replace(.,is.na(.),0)) %>% 
      # Enlever la lignes avec PSF
      filter(PSF %in% aerien1)
    
    data        
    
  }
  

## Fichier data PSF aerien 1 ---------------------------------------------------
  
  aerien <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "A1:I80")) 

  psf_aerien <- aerien %>%  map(., .f = ~ clean_data_aerien(.x)) %>% 
                                   bind_rows(.id = "Semaine Epi")
  
## Fichier Aerien 2 ------------------------------------------------------------
  
  psf_aerien1 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "J1:Q80")) 
  
  psf_aerien_1 <- psf_aerien1 %>%  map(., .f = ~ clean_data_aerien2(.x)) %>% 
    bind_rows(.id = "Semaine Epi")

## Fichier Aerien 3 ------------------------------------------------------------  
    
  psf_aerien2 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "R1:Y80"))
  
  
  psf_aerien_2 <- psf_aerien2 %>%  
    map(., .f = ~ clean_data_aerien2(.x)) %>% 
     bind_rows(.id = "Semaine Epi")
  
## Fichier Aerien 4 ------------------------------------------------------------  
  
  psf_aerien3 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "Z1:AG80"))
  
  
  psf_aerien_3 <- psf_aerien3 %>%  
    map(., .f = ~ clean_data_aerien2(.x)) %>% 
    bind_rows(.id = "Semaine Epi")

## Fichier Aerien 5 -----------------------------------------------------------
  
  psf_aerien4 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AH1:AO80"))
  
  psf_aerien_4 <- psf_aerien4 %>%  
    map(., .f = ~ clean_data_aerien2(.x)) %>% 
    bind_rows(.id = "Semaine Epi")
  
## Fichier Aerien 6 ------------------------------------------------------------
  
  psf_aerien5 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AP1:AW80"))
  
  psf_aerien_5 <- psf_aerien5 %>%  
    map(., .f = ~ clean_data_aerien2(.x)) %>% 
    bind_rows(.id = "Semaine Epi")
  
## Fichier Aerien 7 ------------------------------------------------------------
  
  psf_aerien6 <- sheet_names %>% 
    purrr::set_names() %>% 
    map(.f = ~import(here("data","BASE DES DONNEES PSF COMPILEES (1).xlsx"),which = .x,range = "AX1:BE80"))
  
  psf_aerien_6 <- psf_aerien6 %>%  
    map(., .f = ~ clean_data_aerien2(.x)) %>% 
    bind_rows(.id = "Semaine Epi")

## Combiner les fichiers PSF aeriens ------------------------------------------
  
  combine_psf_aerien <- union_all(psf_aerien,psf_aerien_1) %>% 
             union_all(psf_aerien_2) %>% 
             union_all(psf_aerien_3) %>% 
             union_all(psf_aerien_4) %>% 
             union_all(psf_aerien_5) %>% 
             union_all(psf_aerien_6)
#Exporter des data PSF aeriens ---------------
  
  export(combine_psf_aerien,
         file = here("output","psf_aerien.xlsx"))

## Combiner tous les fichier ---------------
  
   combine_all <- union_all(combined_psf_terrestres,combined_psf_maritimes) %>% 
                  union_all(combine_psf_aerien)
## Export global file --------------------------
  
  export(combine_all,
         file = here("output","combined_psf.xlsx"))
  
      