library(readxl)
library(writexl)
library(tidyverse)
library(tidymodels)
library(ggplot2)
library(rstatix)
library(ggpubr)
library(DescTools)
library(gtsummary)
library(kableExtra)
library(emmeans)
library(mice)
library(pROC)
library(rcompanion)

# What is it: Скачивание кода по ссылке и вставка его сюда.

source("https://raw.githubusercontent.com/aysuvorov/medstats/master/R_scripts/medstats.R")

# +-----------------------------------------------------------------------------

# What is it: Переход в рабочую директорию

path = setwd('/home/vasilisa/work/r/')
SEED = 42

# What is it: Чтение таблицы

df = read_excel('./data/data_26.02_khudor.xlsx', skip = 2) 
old_names = names(df) # названия колонок в 3 строчке (сокращенное английское)

df_tml = read_excel('./data/data_26.02_khudor.xlsx', skip = 1) 
russ_names = names(df_tml) # названия колонок в 2 строчке (русское название без подробного описания данных)

# What is it:
# df = select(df, -c(uin, ...)) удаляет передаваемые столбцы из таблицы
# df = mutate(df, 
#  ...,
#  ...,
#  ...
#) - делает передаваемые изменения в таблице по одному 
# (тут приведения типов и создание новых колонок CAT- сравнение с порогом)
df = df |> select(-c(uin, stenoz_BCA_intraoperacionno, bljashka_intraoperacionno_makropreparat)) |>
    mutate(
        stenoz_operiruemoj_BCA_po_MSKT_CAT = as.numeric(stenoz_operiruemoj_BCA_po_MSKT > 75),
    stenoz_operiruemoj_BCA_UZDS_DIAM_CAT = as.numeric(stenoz_BCA_UZDS_po_diametru > 75),
    stenoz_operiruemoj_BCA_UZDS_SQUARE_CAT = as.numeric(stenoz_BCA_UZDS_po_ploshhadi > 75),
    SKF_MDRD_mlmin173m2 = as.numeric(SKF_MDRD_mlmin173m2)
    )

# What is it: трансформация категориальных признаков (с небольше чем 7 уникальными значениями) в специальный тип factor

df = FactorTransformer(df)

# What is it: заполнение пропусков

mice_names = names(df) # текущие названия колонок в df
tab = df
colnames(tab) = paste0('x', seq(length(mice_names))) # изменения названий колонок на "x1, x2, ..." (видимо для красивого вывода)
tab = complete(mice(tab, m=1, maxit=1, meth='cart', seed=SEED), 1) # заполнение пропусков в таблице
colnames(tab) = mice_names

names(df)

# What is it: записать результат заполнения пропусков для определенных столбцов

for (col in names(df)[c(1:74, 178:180)]) {
    df[[col]] = tab[[col]]
}

# +-----------------------------------------------------------------------------
# What is it: построение таблицы со статистикой по столбцам

stat_table = compare_all(df, "Stabilnost_ASB") 

# What is it: изменения имени в колонке "показатель" на русское и сохранение таблицы

lex_coder_stat_tables(
    stat_table = stat_table,
    column_to_change_values = 'Показатель',
    old_patterns = old_names,
    new_patterns = russ_names) |> write_xlsx(
        paste(path, '/data/1_Descriptives_26_02_2024.xlsx', sep = '')
        )


