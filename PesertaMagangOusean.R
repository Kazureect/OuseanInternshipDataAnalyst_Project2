# LIBRARY
library("tidyverse")
library(lubridate)
library(tidyverse)
library(moments)
library(car)
library(carData)
library(nortest)
library(ggpubr)
library(lubridate)
library(MASS)
library(caret)
library(readr)
library(dplyr)
library(stringr)
library(patchwork)
library(forcats)
library(xlsx)

### READING DATA
dataMagang <- view(Form_Pendaftaran_Internship_Jawaban_xlsx_Form_Responses_1)
str(dataMagang)

### DATA MANIPULATION
dataMgNew <- dataMagang %>%
  rename("Nama" = "Nama Lengkap") %>%
  rename("Instansi" = "Asal Perguruan Tinggi") %>%
  rename("Provinsi" = "Asal Provinsi") %>%
  rename("Departemen Pilihan" = "Departemen apa yang kamu inginkan?") %>%
  rename("Alasan Pemilihan" = "Apa alasan kamu menginginkan departemen tersebut") %>%
  rename("Skill Dimiliki" = "Skill apa yang kamu miliki saat ini?") %>%
  rename("Skill Dipelajari" = "Skill apa yang sedang kamu pelajari saat ini?") %>%
  rename("Kesediaan Magang" = "apakah kamu bersedia untuk melakuan Internship Program di Ousean Group selama 3 bulan?") %>%
  arrange(Nama)
View(dataMgNew)

### DATA CLEANING

## Missing Value
# Check
is.na(dataMgNew)
sapply(dataMgNew, function(x) sum(is.na(x)))
print(paste("Jumlah data missing:", sum(is.na(dataMgNew))))

## Duplicated Value
# Check
jumlah_duplikat <- dataMgNew %>%
  group_by(Nama) %>%
  summarise(jumlah = n()) %>%
  filter(jumlah > 1) %>%
  arrange(desc(jumlah))

jumlah_duplikat

# Delete
data_bersih <- dataMgNew %>%
  distinct(Nama, .keep_all = TRUE)

View(data_bersih)

## Incosistent Value

# Check
  check_instansi <- data_bersih %>%
    count(Instansi, sort = TRUE)
  
print(check_nama)
print(check_instansi)

## Descriptive Statistics

#Number of Participants

participans <- paste("Total Peserta =", length(data_bersih$Nama))

#Gender
data_plotGender <- data_bersih %>%
  count(`Jenis Kelamin`, name = "jumlahG") %>%
  mutate(
    persen = jumlahG / sum(jumlahG),
    label_final = paste0(scales::percent(persen, accuracy = 1), "\n(", jumlahG, ")")
  )

gender <- ggplot(data_plotGender, aes(x = "", y = jumlahG, fill = `Jenis Kelamin`)) +
  geom_bar(stat = "identity", width = 1, color = "white") +
  coord_polar(theta = "y") +
  labs(
    title = "Gender",
    fill = "",
    x = NULL,
    y = NULL
  ) +
  geom_text(aes(label = label_final), position = position_stack(vjust = 0.5), size = 5) +
  theme_void() +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16),
    legend.position = "right"
  )

#Participants by Provinces
data_plotProvinsi <- data_bersih %>%
  mutate(Provinsi_Grouped = fct_lump_n(
    f = Provinsi, 
    n = 5, 
    other_level = "Others"
  )) %>%
  mutate(Provinsi_Grouped = fct_infreq(Provinsi_Grouped) %>% 
           fct_relevel("Others", after = Inf)) %>%
  count(Provinsi_Grouped, name = "jumlahP")

province <- ggplot(data_plotProvinsi, aes(x = Provinsi_Grouped, y = jumlahP, fill = Provinsi_Grouped)) +
  geom_col() +
  geom_text(aes(label = jumlahP), vjust = -0.5) +
  labs(
    title = "Participans by Province",
    x = "Province",
    y = "Participans"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16),
    legend.position = "none" 
  )

#Institutions
semua_instansi <- data_bersih %>%
  count(Instansi, name = "jumlahI", sort = TRUE)

top <- semua_instansi %>% 
  slice_head(n = 8)

others <- semua_instansi %>%
  slice_tail(n = -8) %>%
  summarise(
    Instansi = "Others",
    jumlahI = sum(jumlahI)
  )

level_order <- c("Others", rev(top$Instansi))

data_final <- bind_rows(top, others) %>%
  mutate(Instansi = factor(Instansi, levels = level_order))

institution <- ggplot(data_final, aes(x = jumlahI, y = Instansi)) +
  geom_col(aes(fill = Instansi)) +
  geom_text(aes(label = jumlahI), hjust = -0.2) +
  labs(
    title = "Participants by Institution",
    x = "Participants",
    y = "Institution"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16),
    legend.position = "none" 
  )

#Departement
genderall_dep <- data_bersih %>%
  count(`Departemen Pilihan`, name = "jumlahD", sort = TRUE)

departement <- ggplot(all_dep, aes(x = jumlahD, y = reorder(`Departemen Pilihan`, jumlahD))) +
  geom_col(aes(fill = `Departemen Pilihan`)) +
  geom_text(aes(label = jumlahD), hjust = -0.2) +
  labs(
    title = "Departments of Interest",
    x = "Participants",
    y = "Departments"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16),
    legend.position = "none" 
  )

#Function
participans
province
institution
departement

desc_stats <- gender + province + institution + departement

desc_stats

## UNIQUE METRICS

# Ratio per Departement
rasio_per_departemen <- data_bersih %>%
  count(`Departemen Pilihan`, `Jenis Kelamin`, name = "jumlahDK") %>%
  pivot_wider(
    names_from = `Jenis Kelamin`, 
    values_from = jumlahDK,
    values_fill = 0
  ) %>%
  mutate(
    Rasio_Teks = paste(`Laki-laki`, Perempuan, sep = " : "),
    Rasio_Numerik = if_else(Perempuan == 0, NA, `Laki-laki` / Perempuan)
  )

# Provinces Distribution
distribusi_provinsi <- data_bersih %>%
  count(Provinsi, name = "jumlahPP", sort = TRUE) %>%
  mutate(Persentase_Distribusi = (jumlahPP / sum(jumlahPP)) * 100)


# Skills Owned Distribution
distribusi_skill <- data_bersih %>%
  separate_rows(`Skill Dimiliki`, sep = ",") %>%
  mutate(`Skill Dimiliki` = str_trim(`Skill Dimiliki`)) %>%
  count(`Skill Dimiliki`, name = "Jumlah", sort = TRUE) %>%
  mutate(Persentase = (Jumlah / sum(Jumlah)) * 100)

#Function
unique_metrics <- print(list(
  Rasio_Departemen = rasio_per_departemen,
  Distribusi_Provinsi = distribusi_provinsi,
  Distribusi_Skill = distribusi_skill
  
))

unique_metrics

write.xlsx(data_bersih, file = "hasil_data_bersih.xlsx")