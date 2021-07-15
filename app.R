#Script para scrappear la base de datos de http://www.intermodal-terminals.eu/database/ y crear un excel por una hoja por país con los datos de las terminales
#Realizado por Pablo Nicolas Estevez pablo22estevez@gmail.com

# Instalar paquetes


# Cargar Paquetes
library("rvest")
library("stringi")
library("expss")
library("openxlsx")
library(dplyr)
library(rvest)
library(stringr)
library(shiny)

ui <- fluidPage(
  downloadButton("pato", "Download")
)
server <- function(input, output) {
  Fila <- function(id){
    url <- stri_c("http://www.intermodal-terminals.eu/database/terminal/view/id/", id)
    texto <- as.character(xml_child(read_html(url, verbose = T), 2))
    
    a <- stri_locate_first_regex(texto, "pageContent")[2] +8
    b <- stri_locate_first_regex(texto, "pageContent")[1] + 400
    c0 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c0, "class")[1] - 5
    c0 <- stri_sub(c0, a, b)
    
    a <- stri_locate_first_regex(texto, "Modes Served")[2] +27
    b <- stri_locate_first_regex(texto, "Modes Served")[2] + 400
    c1 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c1, "strong>")[1] - 3
    c1 <- stri_sub(c1, a, b)
    
    a <- stri_locate_first_regex(texto, "Terminal Operator")[2] +27
    b <- stri_locate_first_regex(texto, "Terminal Operator")[2] + 400
    c2 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c2, "strong>")[1] - 3
    c2 <- stri_sub(c2, a, b)
    
    a <- stri_locate_first_regex(texto, "Address")[2] +19
    b <- stri_locate_first_regex(texto, "Address")[2] + 400
    c3 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c3, "td>")[1] - 3
    c3 <- stri_sub(c3, a, b)
    c3 <- stri_replace_all_regex(c3, "<br>", " ")
    
    a <- stri_locate_first_regex(texto, "Contact Person")[2] +28
    b <- stri_locate_first_regex(texto, "Contact Person")[2] + 400
    c4 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c4, "td>")[1] - 3
    c4 <- stri_sub(c4, a, b)
    c4 <- stri_replace_all_regex(c4, "</strong>", "")
    
    a <- stri_locate_first_regex(texto, "Phone")[2] +20
    b <- stri_locate_first_regex(texto, "Phone")[1] + 400
    c5 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c5, "class")[1] - 24
    c5 <- stri_sub(c5, a, b)
    
    
    a <- stri_locate_first_regex(texto, "FAX")[2] +20
    b <- stri_locate_first_regex(texto, "FAX")[2] + 400
    c6 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c6, "td>")[1] - 3
    c6 <- stri_sub(c6, a, b)
    
    
    a <- stri_locate_first_regex(texto, "E-Mail")[2] +35
    b <- stri_locate_first_regex(texto, "E-Mail")[2] + 400
    c7 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c7, ">")[1] - 2
    c7 <- stri_sub(c7, a, b)
    
    a <- stri_locate_first_regex(texto, "Web")[2] +28
    b <- stri_locate_first_regex(texto, "Web")[2] + 400
    c8 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c8, "target=")[1] - 3
    c8 <- stri_sub(c8, a, b)
    
    a <- stri_locate_first_regex(texto, "Terminal Info")[2] + 19
    b <- stri_locate_first_regex(texto, "Terminal Info")[1] + 400
    c9 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c9, "tbody")[1] - 22
    c9 <- stri_sub(c9, a, b)
    c9 <- stri_replace_all_regex(c9, "<br>", "  ")
    
    
    a <- stri_locate_first_regex(texto, "Total Terminal Area")[2] + 27
    b <- stri_locate_first_regex(texto, "Total Terminal Area")[1] + 400
    c10 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c10, "sup")[1] - 4
    c10 <- as.integer(stri_sub(c10, a, b))
    
    
    a <- stri_locate_first_regex(texto, "Handling of")[2] + 19
    b <- stri_locate_first_regex(texto, "Handling of")[1] + 400
    c11 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c11, "br")[1] - 2
    c11 <- stri_sub(c11, a, b)
    
    a <- stri_locate_first_regex(texto, "total number of tracks")[2] + 3
    b <- stri_locate_first_regex(texto, "total number of tracks")[1] + 400
    c12 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c12, "br")[1]-2
    c12 <- as.integer(stri_sub(c12, a, b))
    
    a <- stri_locate_first_regex(texto, "total usable length")[2] + 3
    b <- stri_locate_first_regex(texto, "total usable length")[1] + 400
    c13 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c13, "td")[1]-5
    c13 <- as.integer(stri_sub(c13, a, b))
    
    a <- stri_locate_first_regex(texto, "Gantry Cranes")[2] + 19
    b <- stri_locate_first_regex(texto, "Gantry Cranes")[1] + 400
    c14 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c14, "td")[1]-8
    c14 <- stri_sub(c14, a, b)
    c14 <- stri_replace_all_regex(c14, "<br>", "  ")
    
    a <- stri_locate_first_regex(texto, "Reachstackers")[2] + 19
    b <- stri_locate_first_regex(texto, "Reachstackers")[1] + 400
    c15 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c15, "br")[1]-2
    c15 <- stri_sub(c15, a, b)
    
    a <- stri_locate_first_regex(texto, "Depot")[1] + 23
    b <- stri_locate_first_regex(texto, "Depot")[1] + 400
    c16 <- stri_sub(texto, a, b)
    a <- 1
    b <- stri_locate_first_regex(c16, "td")[1]-3
    c16 <- stri_sub(c16, a, b)
    
    
    data.frame(NA, c0,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10, c11,c12,c13,c14,c15,c16)
    
    
  }
  
  output$pato <- downloadHandler(
    filename = function() { "terminales.xlsx"},
    content = {
      #Crea una tabla con todos los datos
      tabla <- Fila(2)
      n <- 3
      for(n in 3:10){
        url <- stri_c("http://www.intermodal-terminals.eu/database/terminal/view/id/", n)
        mtray <- try(as.character(xml_child(read_html(url), 2)))
        if (inherits(mtray, "try-error")){
          n <- n +1 
        }else{
          fila <- Fila(n)
          tabla <- rbind(tabla, fila)
          n <- n+1
        }
      }
      colnames(tabla) <- c(" "," ", "Modes Served",	"Terminal Operator",	"Address",	"Contact Person",	"Phone",	"Fax",	"e-mail",	"web",	"Terminal Info",	"Total Terminal Area",	"Handling of	Rails",	"total number of tracks",	"Total usabel length",	"Gantry Cranes",	"Reachstackers",	"Depot")
      
      #Crea la variable país,y país 2 que junta a los chicos, divide en grupos a los datos
      a <- stri_locate_last_regex(tabla$Address, " ")[,1]+1
      b <- stri_length(tabla$Address)
      pais <- stri_sub(tabla$Address, a, b)
      pais[which(pais == "OF")] <- "Macedonia"
      pais[which(pais == "FEDERATION")] <- "Rusia"
      pais <- str_to_title(pais)
      unico <- (unique(pais))
      pais2 <- pais
      n <- 1
      for(n in 1:length(unico)){
        if (count_if(unico[n], pais2)<4){
          pais2[which(pais == unico[n])] <- "chico"
        }
      }
      dividida <- split.data.frame(tabla, pais)
      dividida2 <- split.data.frame(tabla, pais2)
      pais <- sort(unique(pais))
      pais2 <- sort(unique(pais2))
      
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
      n <- 1
      wb = createWorkbook()
      for(n in 1:length(pais)){
        sh = addWorksheet(wb, pais[n])
        xl_write(dividida[n], wb, pais[n])
        addStyle(wb, pais[n], negrita, 1, 3:18, gridExpand = FALSE, stack = FALSE)
        n = n+1
      }
      saveWorkbook(wb, file, overwrite = TRUE)}
  )
}
shinyApp(ui, server)
