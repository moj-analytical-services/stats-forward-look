library(magrittr)
library(dplyr)
library(jsonlite)
library(rvest)
library(stringr)
library(openxlsx)
library(lubridate)
library(tidyr)

searchurl <- "https://www.gov.uk/search/research-and-statistics?content_store_document_type=upcoming_statistics&organisations=ministry-of-justice&order=release-date-oldest"

pages <- read_html(searchurl) %>%
  html_nodes(".govuk-pagination__link-label") %>%
  html_text()

if (length(pages) == 0) {
  pages <- "1"
} else {
  pages <- str_replace(pages,"2 of ","")
}

for (i in 1:as.numeric(pages)) {
  
  searchurl <- paste0(searchurl,"&page=",i)
  
  publication.url <- read_html(searchurl) %>%
    html_nodes(".gem-c-document-list") %>% 
    html_nodes("a") %>% 
    html_attr("href")
  
  publication.name <- read_html(searchurl) %>%
    html_nodes(".gem-c-document-list") %>% 
    html_nodes("a") %>% 
    html_text()
  
  pubattribs <- read_html(searchurl) %>%
    html_nodes(".gem-c-document-list__item")
  
  getinfo <- function(listitems) {
    
    attribs <- html_nodes(listitems,".gem-c-document-list__attribute") %>% html_text()
    
    attribs <-str_split(attribs,": ")
    
    attribs <- sapply(attribs, "[[", 2)
    
    attribs <- str_replace(attribs,"\n","")
    
    return(attribs)
  
}

pubattribs <- t(as.data.frame(lapply(pubattribs,getinfo)))

pubattribs <- as.data.frame(pubattribs)

rownames(pubattribs) <- NULL

names(pubattribs) <- c("publication.type","publication.department","publication.date","publication.status")

prerelease <- as.data.frame(cbind(publication.name,publication.url,pubattribs))

if (i == 1) {
  prerelease_all <- prerelease
} else {
  prerelease_all <- rbind(prerelease_all,prerelease)
}

}

prerelease_all$publication.date <- stringr::str_remove(prerelease_all$publication.date," 9:30am")

prerelease_all$Week <- lubridate::isoweek(dmy(prerelease_all$publication.date))
prerelease_all$Year <- lubridate::year(dmy(prerelease_all$publication.date))


Allweeks <- as.data.frame(c(1:52)) %>%
  crossing(unique(prerelease_all$Year))

names(Allweeks) <- c("Week","Year")


prerelease_weeks <- left_join(Allweeks,filter(prerelease_all,publication.status != "cancelled")) %>%
  arrange(Year,Week)

prerelease_weeks$WC <- format(as.Date(paste(prerelease_weeks$Year, prerelease_weeks$Week, 1, sep="-"), "%Y-%U-%u"),"%d %b %Y")

rowvector <- 1:nrow(prerelease_weeks)

prerelease_all <- prerelease_weeks[min(rowvector[!is.na(prerelease_weeks$publication.date)]):max(rowvector[!is.na(prerelease_weeks$publication.date)]),]

names(prerelease_all) <- c("Week",
                           "Year",
                           "Publication Title", 
                           "Announcement URL",
                           "Statistics Type",
                           "Department",
                           "Publication Date",
                           "Status",
                           "Week Commencing")

govuk_link <- c("https://www.gov.uk/search/research-and-statistics?content_store_document_type=upcoming_statistics&organisations%5B%5D=ministry-of-justice&order=release-date-oldest")
names(govuk_link) <- c("Click here to view on the gov.uk Research and Statistics calendar")
class(govuk_link) <- "hyperlink"

wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb,"Forward Look")
openxlsx::writeData(wb,"Forward Look","MoJ Statistics Forward Look",
                    startRow = 1)
openxlsx::writeData(wb,"Forward Look",paste("This list contains a week-by-week view of  MoJ Official and National Statistics that have been pre-announced on the gov.uk release calendar as at",format(Sys.Date(),"%d %B %Y")), startRow = 2)
openxlsx::writeData(wb,"Forward Look",govuk_link,startRow = 3)
openxlsx::writeData(wb,"Forward Look",select(prerelease_all,
                                             c("Week Commencing","Publication Title","Publication Date","Status","Week")),startRow = 4)

titleStyle <- createStyle(fontSize = 14, textDecoration = "bold")
subtitleStyle <- createStyle(fontSize = 12)
linkStyle <- createStyle(fontSize = 12, valign = "top")
evenStyle <- createStyle(bgFill = "#b4c6e7")
oddStyle <- createStyle(bgFill = "#d9e1f2")
hideStyleEven <- createStyle(bgFill = "#b4c6e7", fontColour = "#b4c6e7")
hideStyleOdd <- createStyle(bgFill = "#d9e1f2", fontColour = "#d9e1f2")
border <- createStyle(border="top", borderColour = "#FFFFFF")

conditionalFormatting(wb, "Forward Look", cols = 1:5, rows = 1:nrow(prerelease_all)+4, rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=0)",
                      style = evenStyle)
conditionalFormatting(wb, "Forward Look", cols = 1:5, rows = 1:nrow(prerelease_all)+4, rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=1)",
                      style = oddStyle)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+4, rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=0,$E5=$E4)",
                      style = hideStyleEven)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+4, rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=1,$E5=$E4)",
                      style = hideStyleOdd)
conditionalFormatting(wb, "Forward Look", cols = 1:5, rows = 1:nrow(prerelease_all)+4, rule = "=AND($E5<>$E4)",
                      style = border)

setColWidths(wb,1,cols = c(1:5),widths=c(18,"auto",24,12,12),hidden=c(rep(FALSE,4),TRUE))
setRowHeights(wb,1,3,30)

header_st <- createStyle(fgFill = "#1F497D", textDecoration = "Bold", fontColour = "#FFFFFF")

cell_st <- createStyle(halign = "left")

openxlsx::addStyle(wb,1,header_st,4,c(1:5))
addStyle(wb, 1, style = titleStyle, rows = 1, cols = 1)
addStyle(wb, 1, style = subtitleStyle, rows = 2, cols = 1)
addStyle(wb, 1, style = linkStyle, rows = 3, cols = 1, stack = TRUE)

addStyle(wb, 1, style = cell_st, cols = 1:5, rows = 5:nrow(prerelease_all)+4, gridExpand = TRUE, stack = TRUE)

showGridLines(wb, 1, showGridLines = FALSE)

openxlsx::saveWorkbook(wb, paste0("Forward Look/Forward Look.xlsx"), overwrite = TRUE)

