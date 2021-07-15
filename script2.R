#Script para scrappear la base de datos de http://www.intermodal-terminals.eu/database/ y crear un excel por una hoja por país con los datos de las terminales
#Realizado por Pablo Nicolas Estevez pablo22estevez@gmail.com

# Instalar paquetes
install.packages("rvest")
install.packages("stringi")
install.packages("openxlsx")
install.packages("expss")

# Cargar Paquetes
library("rvest")
library("stringi")
library("expss")
library("openxlsx")
library(dplyr)
library(rvest)
library(stringr)

#Para que google no te crea robot
url = "http://google.com"
download.file(url, destfile = "scrapedpage.html", quiet=TRUE)
content <- read_html("scrapedpage.html")
options(timeout = 500000)

#Definir año final, mes final, día final, id final:
anof <- 2020
mesf <- 11
diaf <- 28
id <- 428800

#Definir año inicial, mes inicial, día inicial
anoi <- 2020
mesi <- 11
diai <- 22

#se determina las celdas que saca jugador 1 y 2
setes1 <- 12
setes2 <- 14
games1 <- integer()
games2 <- integer()
h <- 1
for(h in 1:5){
games1 <- c(games1, setes1)
r <- games1[length(games1)]
games2 <- c(games2, setes2)
q <- games2[length(games2)] 
n <- 1
for(n in 1:5){
  r <- r+4
  q <- q+4
  games1 <- c(games1, r)
  games2 <- c(games2, q)
  n <- n+1
}
if(setes2 > setes1){
setes1 <- setes1 +29
setes2 <- setes2 +25
h <- h+1
}else{
  setes1 <- setes1 +25
  setes2 <- setes2 +29
  h <- h+1
}
}

#Funcion para leer los sets
sets <- function(set){
a <- stri_locate_first_regex(texto, set)[2] +3
b <- stri_locate_first_regex(texto, set)[2] +500
cortado <- stri_sub(texto, a, b)
b <- stri_locate_first_regex(cortado, "<")[1] -1
stri_sub(cortado, 1, b)
}

#Funcion para leer los games, en orden
games <- function(){
a <- stri_locate_first_regex(cortado, "gamescore")[2] +3
cortado <<- stri_sub(cortado, a, stri_length(cortado))
a <- stri_locate_first_regex(cortado, " ")[2] +1
b <- stri_locate_first_regex(cortado, "<")[1] -1
celda <- stri_sub(cortado, a, b)
celda <- stri_replace_all_regex(celda, ":", "-")
quiebre <- !is.na(stri_locate_first_regex(celda, ","))
b <- stri_locate_first_regex(celda, ",")[1] -1
if(!is.na(stri_sub(celda, 1, b))){celda <- stri_sub(celda, 1, b)}
a <- stri_locate_last_regex(celda, " ")[2] +1
if(is.na(stri_sub(celda, a, stri_length(celda)))){
  if(is.na(stri_locate_last_regex(celda, " "))){resultado <<- celda}else{resultado <<- ""}
  }else{
resultado <<- stri_sub(celda, a, stri_length(celda))}
if(!is.na(resultado)){
  resultado <- str_replace_all(resultado, "\\*", "")
if((resultado == "40-40" | resultado == "30-40" | resultado == "15-40" | resultado == "0-40") & !quiebre){resultado <- "A-40"}
if((resultado == "30-30" | resultado == "15-30" | resultado == "0-30") & !quiebre){resultado <- "40-30"}
if((resultado == "30-15" | resultado == "15-15" | resultado == "0-15") & !quiebre){resultado <- "40-15"}
if((resultado == "30-0" | resultado == "15-0" | resultado == "0-0") & !quiebre){resultado <- "40-0"}
if((resultado == "40-40" | resultado == "40-30" | resultado == "40-15" | resultado == "40-0") & quiebre){resultado <- "40-A"}
if((resultado == "30-30" | resultado == "30-15" | resultado == "30-0") & quiebre){resultado <- "30-40"}
if((resultado == "15-30" || resultado == "15-15" || resultado == "15-0") & quiebre){resultado <- "15-40"}
if((resultado == "0-30" | resultado == "0-15" | resultado == "0-0") & quiebre){resultado <- "0-40"}
}
  resultado <<- resultado
}

#Funcion para contar los puntos
puntos <- function(){
  c <- "Tiebreak"
  if(resultado == "40-0" | resultado == "0-40"){c <- 4}
  if(resultado == "40-15" | resultado == "15-40"){c <- 5}
  if(resultado == "40-30" | resultado == "30-40"){c <- 6}
  if(resultado == "40-A" | resultado == "A-40"){c <-"+7"}
  if(resultado ==""){c <- ""}
  c
}

#Funcion para ordenar games y sets y colocarlo en la tabla
orde <- function(numero){
  if(!is.na(numero)){
  for(n in (1:numero)){
    fi <<- cbind(fi, g[length(g)-1- ((n-1)*2)-anterior*2], g[length(g)- ((n-1)*2) - anterior*2])
    n = n+1}
  if(numero !=13){for(n in 1:(13-numero)){fi <<- cbind(fi, "", "")}}
}else{
  n <- 1
  for(n in 1:26){ 
    fi <<- cbind(fi, "")}
}}

# Funcion que devuelve un dataframe de una linea para el id correspondiente
Fila <- function(ano, mes, dia, id){
url <- stri_c("https://www.tennisbetsite.com/scores/stats/", ano, "/", mes, "/", dia, "/", id, ".html")
texto <<- as.character(xml_child(read_html(url, verbose = T), 2))

c0 <- stri_c(dia, "/", mes, "/", ano)


a <- stri_locate_first_regex(texto, "Singles")[2]
if(is.na(a)){
  c1 <- "DOBLES"
}else{
  c1 <- "SINGLES"
}

if (c1 == "SINGLES"){
a <- stri_locate_first_regex(texto, "playerlink")[2] + 16
b <- stri_locate_first_regex(texto, "playerlink")[2] + 200
c2 <- stri_sub(texto, a, b)
a <- 1
b <- stri_locate_first_regex(c2,">")[1] -2
c2 <- stri_sub(c2, a, b)
}else{
  a <- stri_locate_first_regex(texto, "playerlink")[2] + 16
  b <- stri_locate_first_regex(texto, "playerlink")[2] + 300
  tx <- stri_sub(texto, a, b)
  a <- 1
  b <- stri_locate_first_regex(tx,">")[1] -2
  jugador1 <- stri_sub(tx, a, b)
  a <- stri_locate_first_regex(tx,"title")[1] + 7
  b <- 300
  tx <- stri_sub(tx, a, b)
  a <- 1
  b <- stri_locate_first_regex(tx,">")[1] -2
  jugador2 <- stri_sub(tx, a, b)
  c2 <- stri_c(jugador1, " y ", jugador2)
}

if (c1 == "SINGLES"){
  a <- stri_locate_first_regex(texto, c2)[2] + 10
  b <- stri_locate_first_regex(texto, c2)[2] + 1000
  c3 <- stri_sub(texto, a, b)
  a <- stri_locate_first_regex(c3, "playerlink")[2] + 16
  b <- stri_locate_first_regex(c3, "playerlink")[2] + 200
  c3 <- stri_sub(c3, a, b)
  a <- 1
  b <- stri_locate_first_regex(c3,">")[1] -2
  c3 <- stri_sub(c3, a, b)
}else{
  a <- stri_locate_first_regex(texto, jugador2)[2]
  b <- stri_locate_first_regex(texto, jugador2)[2] + 1000
  tx <- stri_sub(texto, a, b)
  a <- stri_locate_first_regex(tx, "playerlink")[2] + 16
  b <- stri_locate_first_regex(tx, "playerlink")[2] + 200
  tx <- stri_sub(tx, a, b)
  a <- 1
  b <- stri_locate_first_regex(tx,">")[1] -2
  jugador1 <- stri_sub(tx, a, b)
  a <- stri_locate_first_regex(tx,"title")[1] + 7
  b <- 300
  tx <- stri_sub(tx, a, b)
  a <- 1
  b <- stri_locate_first_regex(tx,">")[1] -2
  jugador2 <- stri_sub(tx, a, b)
  c3 <- stri_c(jugador1, " y ", jugador2)
}


if (is.na(stri_locate_first_regex(texto, "Clay"))){
  if(is.na(stri_locate_first_regex(texto, "Grass"))){
    if(is.na(stri_locate_first_regex(texto, "Carpet"))){
      if(is.na(stri_locate_first_regex(texto, "I.hard"))){
      c5 <- "Dura"
      }else{c5 <- "Sintética"}
    }else{c5 <- "Carpet"}
  }else{c5 <- "Césped"}
}else { c5 <- "Clay"}
    
a <- stri_locate_first_regex(texto, "<h2>")[2] +1
b <- stri_locate_first_regex(texto, "</h2>")[1] -1
c6 <- stri_sub(texto, a, b)
b <- stri_locate_first_regex(c6, "Challenger")[1] -2
if(!is.na(stri_sub(c6, 1, b))) { c6 <- stri_sub(c6, 1, b)}
c6 <- tail(strsplit(c6, split=" ")[[1]],1)
if(c6 == "Paulo"){c6 <- "Sao Paulo"}
if(c6 == "Canaria"){c6 <- "Las Palmas de Gran Canaria"}
if(c6 == "Sheikh"){c6 <- "Sharm El Sheikh"}
if(c6 == "Lago"){c6 <- "Quinta Do Lago"}


a <- stri_locate_first_regex(texto, "title")[2] +3
b <- stri_locate_first_regex(texto, "title")[2] +50
pais <- stri_sub(texto, a, b)
b <- stri_locate_first_regex(pais, ">")[1] -2
pais <- stri_sub(pais, 1, b)
c6 <- stri_c(c6, " / ", pais)

if(is.na(stri_locate_first_regex(texto, ",Women,")[1])){
  c7 <- "Masculino"
}else{
  c7 <- "Femenino"
}

a <- stri_locate_first_regex(texto, "<h2>")[2] +1
b <- stri_locate_first_regex(texto, "</h2>")[1] -1
c4 <- stri_sub(texto, a, b)
if(!is.na(stri_locate_first_regex(c4, "Challenger"))){c4 <- "Challenger"}else{if(!is.na(stri_locate_first_regex(c4, "Australian"))){c4 <- "Grand Slam"}
  if(!is.na(stri_locate_first_regex(c4, "Garros"))){c4 <- "Grand Slam"}
  if(!is.na(stri_locate_first_regex(c4, "Wimbledon"))){c4 <- "Grand Slam"}
  if(!is.na(stri_locate_first_regex(c4, "US Open"))){c4 <- "Grand Slam"}
}
l <- stri_sub(c4, 1, 2)
if(l == "W1" | l == "M1" | l == "W2" | l == "M2"){c4 <- "ITF"}
if(c4 != "Challenger" & c4 != "Grand Slam" & c4 != "ITF"){
  if(c7 == "Masculino"){c4 <- "ATP"}else{c4 <- "WTA"}
}


if (sets("gi_p1_set0") != ""){
  if (sets("gi_p1_set3") == ""){
  c8 <-stri_c(sets("gi_p1_set0"), "-", sets("gi_p2_set0"), " ", sets("gi_p1_set1"), "-", sets("gi_p2_set1"), " ", sets("gi_p1_set2"), "-", sets("gi_p2_set2"))
  
}else{
c8 <-stri_c(sets("gi_p1_set0"), "-", sets("gi_p2_set0"), " ", sets("gi_p1_set1"), "-", sets("gi_p2_set1"), " ", sets("gi_p1_set2"), "-", sets("gi_p2_set2"), " ", sets("gi_p1_set3"), "-", sets("gi_p2_set3"), " ", sets("gi_p1_set4"), "-", sets("gi_p2_set4"))
}}else{c8 <- ""}

p1 <- 0
p2 <- 0

if (sets("gi_p1_set0") == 7){
  p1 <- p1+2
}else{
  if(sets("gi_p1_set0") == 6){ 
  p1 <- p1+1}
}
if (sets("gi_p1_set1") == 7){
  p1 <- p1+2
}else{
  if(sets("gi_p1_set1") == 6){ 
    p1 <- p1+1}
}
if (sets("gi_p1_set2") == 7){
  p1 <- p1+2
}else{
  if(sets("gi_p1_set2") == 6){ 
    p1 <- p1+1}
}
if (sets("gi_p1_set3") == 7){
  p1 <- p1+2
}else{
  if(sets("gi_p1_set3") == 6){ 
    p1 <- p1+1}
}
if (sets("gi_p1_set4") == 7){
  p1 <- p1+2
}else{
  if(sets("gi_p1_set4") == 6){ 
    p1 <- p1+1}
}
if (sets("gi_p2_set0") == 7){
  p2 <- p2+2
}else{
  if(sets("gi_p2_set0") == 6){ 
    p2 <- p2+1}
}
if (sets("gi_p2_set1") == 7){
  p2 <- p2+2
}else{
  if(sets("gi_p2_set1") == 6){ 
    p2 <- p2+1}
}
if (sets("gi_p2_set2") == 7){
  p2 <- p2+2
}else{
  if(sets("gi_p2_set2") == 6){ 
    p2 <- p2+1}
}
if (sets("gi_p2_set3") == 7){
  p2 <- p2+2
}else{
  if(sets("gi_p2_set3") == 6){ 
    p2 <- p2+1}
}
if (sets("gi_p2_set4") == 7){
  p2 <- p2+2
}else{
  if(sets("gi_p2_set4") == 6){ 
    p2 <- p2+1}
}

if(p1 > p2){
  c9 <- c2
}else{
  c9 <- c3
}

c10 <- stri_c(sets("gi_p1_set0"), "-", sets("gi_p2_set0"))
if (c10 == "-"){c10 <- ""}
c11 <- stri_c(sets("gi_p1_set1"), "-", sets("gi_p2_set1"))
if (c11 == "-"){c11 <- ""}
c12 <- stri_c(sets("gi_p1_set2"), "-", sets("gi_p2_set2"))
if (c12 == "-"){c12 <- ""}
c13 <- stri_c(sets("gi_p1_set3"), "-", sets("gi_p2_set3"))
if (c13 == "-"){c13 <- ""}
c14 <- stri_c(sets("gi_p1_set4"), "-", sets("gi_p2_set4"))
if (c14 == "-"){c14 <- ""}
if (c14 != ""){ex <- 5}else{
if (c13 != ""){ex <- 4}else{
if (c12 != ""){ex <- 3}else{
if (c11 != ""){ex <- 2}else{
if (c10 != ""){ex <- 1}else
ex <- 0}}}}


cortado <<- texto


fi <<- data.frame(c0,c1,c2,c3,c4,c5,c6,c7,c8,ex,c9)

g <<-data.frame(0)
games()
n <- 1
for(n in 1:65){
  games()
  if(is.na(resultado)){break}
  if(resultado != "" & resultado != "score" ){
g <<- cbind(g, resultado, puntos())}
n <- n+1
}
set1 <- as.integer(sets("gi_p1_set0")) + as.integer(sets("gi_p2_set0"))
set2 <- as.integer(sets("gi_p1_set1")) + as.integer(sets("gi_p2_set1"))
set3 <- as.integer(sets("gi_p1_set2")) + as.integer(sets("gi_p2_set2"))
set4 <- as.integer(sets("gi_p1_set3")) + as.integer(sets("gi_p2_set3"))
set5 <- as.integer(sets("gi_p1_set4")) + as.integer(sets("gi_p2_set4"))

fi <<- cbind(fi, c10)
n <- 1
anterior <<- 0
orde(set1)
fi <<- cbind(fi, c11)
anterior <<- set1
orde(set2)
fi <<- cbind(fi, c12)
anterior <<-anterior+set2
orde(set3)
fi <<- cbind(fi, c13)
anterior <<-anterior+set3
orde(set4)
fi <<- cbind(fi, c14)
anterior <<-anterior+set4
orde(set5)

a <- 0
b <- 0
c <- 0
d <- 0
for(n in 12:38){
  if(fi[n] == 4){a <- a+1}
  if(fi[n] == 5){b <- b+1}
  if(fi[n] == 6){c <- c+1}
  if(fi[n] == "+7"){d <- d+1}
}
fi <<- cbind(fi, a, b, c, d)

a <- 0
b <- 0
c <- 0
d <- 0
for(n in 39:65){
  if(fi[n] == 4){a <- a+1}
  if(fi[n] == 5){b <- b+1}
  if(fi[n] == 6){c <- c+1}
  if(fi[n] == "+7"){d <- d+1}
}
fi <<- cbind(fi, a, b, c, d)

a <- 0
b <- 0
c <- 0
d <- 0
for(n in 66:92){
  if(fi[n] == 4){a <- a+1}
  if(fi[n] == 5){b <- b+1}
  if(fi[n] == 6){c <- c+1}
  if(fi[n] == "+7"){d <- d+1}
}
fi <<- cbind(fi, a, b, c, d)

a <- 0
b <- 0
c <- 0
d <- 0
for(n in 93:119){
  if(fi[n] == 4){a <- a+1}
  if(fi[n] == 5){b <- b+1}
  if(fi[n] == 6){c <- c+1}
  if(fi[n] == "+7"){d <- d+1}
}
fi <<- cbind(fi, a, b, c, d)

a <- 0
b <- 0
c <- 0
d <- 0
for(n in 120:145){
  if(fi[n] == 4){a <- a+1}
  if(fi[n] == 5){b <- b+1}
  if(fi[n] == 6){c <- c+1}
  if(fi[n] == "+7"){d <- d+1}
}
fi <<- cbind(fi, a, b, c, d)

n <- games1[1]
h <- 1
quiebres1 <- 0
for(n in games1){
  if(fi[n] == "40-A" | fi[n] == "30-40" | fi[n] == "15-40" | fi[n] == "0-40"){
    quiebres1 <- quiebres1 + 1
  }
  h <- h+1
  n <- games1[h]
}

n <- games2[1]
h <- 1
quiebres2 <- 0
for(n in games2){
  if(fi[n] == "40-A" | fi[n] == "30-40" | fi[n] == "15-40" | fi[n] == "0-40"){
    quiebres2 <- quiebres2 + 1
  }
  h <- h+1
  n <- games2[h]
}
fi <<- cbind(fi, quiebres1, quiebres2, url)

fi
}



#Crea una tabla con todos los datos
tabla <- Fila(ano, mes, dia, id)
colnames(tabla) <- c("FECHA",	"TIPO PARTIDO",	"TENISTA A",	"TENISTA B",	"TIPO TORNEO",	"TIPO SUPERFICIE",	"TORNEO",	"GENERO",	"MARCADOR GENERAL", 	"CANTIDAD DE SETS",	"GANADOR",	"MARCADOR SET 1",	"GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS",	"MARCADOR SET 2", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 3", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 4", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 5", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "Breaks Jugador 1", "Breaks Jugador 2",	"URL")
id <- id - 1
n <- 1
ano <- anof
mes <- mesf
dia <- diaf
lista <- integer()
for(n in 1:10){
  n <- n+1
  url <- stri_c("https://www.tennisbetsite.com/scores/stats/", ano, "/", mes, "/", dia, "/", id, ".html"
                )
  if(ano == anoi & mes == mesi & dia == diai){
    ano <- anof
    mes <- mesf
    dia <- diaf
    lista <<- c(lista,id)
    id <- id-1}else{
  mtray <- try(as.character(xml_child(read_html(url), 2)))
  if (inherits(mtray, "try-error")){
    if (dia > 1){
      dia <- dia -1
      }else{
        if (mes > 1) {
        dia <- 31
        mes <- mes -1
        }else{
          dia <- 31
          mes <- 12
          ano <- ano -1
        }
      }
    } else{
  if(ano == anoi & mes == mesi & dia == diai){break}
  fila <- Fila(ano, mes, dia, id)
  if((fila[12] != 0 & fila[12] != "")& length(fila) == 169){
  colnames(fila) <- c("FECHA",	"TIPO PARTIDO",	"TENISTA A",	"TENISTA B",	"TIPO TORNEO",	"TIPO SUPERFICIE",	"TORNEO",	"GENERO",	"MARCADOR GENERAL",	"CANTIDAD DE SETS", "GANADOR",	"MARCADOR SET 1",	"GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS",	"MARCADOR SET 2", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 3", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 4", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "MARCADOR SET 5", "GAME 1",	"No PUNTOS",	"GAME 2",	"No PUNTOS",	"GAME 3",	"No PUNTOS",	"GAME 4",	"No PUNTOS",	"GAME 5",	"No PUNTOS",	"GAME 6",	"No PUNTOS",	"GAME 7",	"No PUNTOS",	"GAME 8",	"No PUNTOS",	"GAME 9",	"No PUNTOS",	"GAME 10",	"No PUNTOS",	"GAME 11",	"No PUNTOS",	"GAME 12",	"No PUNTOS",	"GAME 13",	"No PUNTOS", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "4",	"5",	"6",	"+7", "Breaks Jugador 1", "Breaks Jugador 2",	"URL")
  tabla <- rbind(fila, tabla)}else{ lista <<- c(lista,id)}
  id <- id - 1
  }
}}

#Crea el formato en negrita
negrita <- createStyle(
  fontName = NULL,
  fontSize = NULL,
  fontColour = NULL,
  numFmt = "GENERAL",
  border = "TopBottomLeftRight",
  borderColour = getOption("openxlsx.borderColour", "black"),
  borderStyle = getOption("openxlsx.borderStyle", "thin"),
  bgFill = NULL,
  fgFill = NULL,
  halign = NULL,
  valign = NULL,
  textDecoration = "BOLD",
  wrapText = FALSE,
  textRotation = NULL,
  indent = NULL,
  locked = NULL,
  hidden = NULL
)

#crea los dos excel

wb = createWorkbook()
sh = addWorksheet(wb, "Hoja1")
xl_write(tabla, wb, "Hoja1")
addStyle(wb, "Hoja1", negrita, 1, 1:169, gridExpand = FALSE, stack = FALSE)
saveWorkbook(wb, "tenis.xlsx", overwrite = TRUE)
