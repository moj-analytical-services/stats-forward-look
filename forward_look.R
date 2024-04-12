# <------------------------------------------------------- LIBRARIES ------------------------------------------------------->

install.packages("renv")
library(renv)
renv::restore()

library(magrittr)
library(dplyr)
library(jsonlite)
library(rvest)
library(stringr)
library(openxlsx)
library(lubridate)
library(tidyr)


# <------------------------------------------------------- FUNCTIONS ------------------------------------------------------->
process_html <- function(url, node) {
  processed <- read_html(url) %>%
    html_nodes(node) 
  
  return(processed)
}

get_info <- function(listitems) {
  attribs <- html_nodes(listitems,".gem-c-document-list__attribute") %>% html_text()
  attribs <- str_split(attribs,": ")
  attribs <- sapply(attribs, "[[", 2)
  attribs <- str_replace(attribs,"\n","")
  
  return(attribs)
}


# <------------------------------------------------------- MAIN CODE ------------------------------------------------------->
search_url <- "https://www.gov.uk/search/research-and-statistics?content_store_document_type=upcoming_statistics&organisations=ministry-of-justice&order=release-date-oldest"
pages      <- process_html(search_url, ".govuk-pagination__link-label") %>%
      html_text()

if (length(pages) == 0) {
    pages <- "1"
} else {
    pages <- str_replace(pages,"2 of ","")
}

for (i in 1:as.numeric(pages)) {
  search_url <- paste0(search_url,"&page=",i)
  
  publication_url  <- process_html(search_url, ".gem-c-document-list") %>% 
    html_nodes("a") %>% 
    html_attr("href")
  publication_name <- process_html(search_url, ".gem-c-document-list") %>% 
    html_nodes("a") %>%
    html_text()
  pub_attribs      <- process_html(search_url, ".gem-c-document-list__item")
  
  pub_attribs <- t(as.data.frame(lapply(pub_attribs,get_info)))
  pub_attribs <- as.data.frame(pub_attribs)
  
  rownames(pub_attribs) <- NULL
  names(pub_attribs)    <- c("publication.type",
                             "publication.department",
                             "publication.date.temp",
                             "publication.status")
  
  prerelease <- as.data.frame(cbind(publication_name,publication_url,pub_attribs))
  
  if (i == 1) {
      prerelease_all <- prerelease
  } else {
      prerelease_all <- rbind(prerelease_all,prerelease)
  }
  
}

prerelease_all$publication.date.temp <- stringr::str_remove(prerelease_all$publication.date.temp," 9:30am")
prerelease_all$publication.date.temp <- stringr::str_remove(prerelease_all$publication.date.temp," 9:03am")

prerelease_all2 <- prerelease_all %>%
  mutate(publication.date = 
           case_when(
             grepl("^[[:digit:]]+", publication.date.temp) == TRUE ~ publication.date.temp,
             grepl("February", publication.date.temp) == TRUE & grepl("2024|2028|2032|2036|2040", publication.date.temp) == TRUE ~ paste0("29 ", publication.date.temp),
             grepl("February", publication.date.temp) == TRUE & grepl("2024|2028|2032|2036|2040", publication.date.temp) == FALSE ~ paste0("28 ", publication.date.temp),
             grepl("April|June|September|November", publication.date.temp) == TRUE ~ paste0("30 ", publication.date.temp),
             TRUE ~ paste0("31 ", publication.date.temp)
           ),
         Week = lubridate::isoweek(dmy(publication.date)),
         Year = lubridate::year(dmy(publication.date))) %>%
  select(-publication.date.temp)

dates <- as.data.frame(format(seq(as.Date("2 January 2023", format = "%d %b %Y"), as.Date("27 December 2026", format = "%d %b %Y"), by="days"), format="%d %b %Y"))

dates_new <- dates %>%
  mutate(Week = rep(1:52, each = 7, times = 4)) %>%
  mutate(week_num2 = rep(1:7, each = 1, times = nrow(dates)/7)) %>%
  rename("Week_commences" = 1) %>%
  mutate(Year = as.numeric(substr(Week_commences, nchar(Week_commences)-4, nchar(Week_commences))))%>%
  filter(week_num2 == 1)


prerelease_weeks <- dates_new %>%
  full_join(prerelease_all2 %>% filter(publication.status != "cancelled"), by = c("Week" = "Week",
                                                                                  "Year" = "Year")) %>%
  arrange(Year,Week) %>%
  select(Week, Year, publication_name, publication_url, publication.type, publication.department, publication.date, publication.status, Week_commences)


rowvector      <- 1:nrow(prerelease_weeks)
prerelease_all <- prerelease_weeks[min(rowvector[!is.na(prerelease_weeks$publication.date)]):
                                     max(rowvector[!is.na(prerelease_weeks$publication.date)]),]

names(prerelease_all) <- c("Week",
                           "Year",
                           "Publication Title", 
                           "Announcement URL",
                           "Statistics Type",
                           "Department",
                           "Publication Date",
                           "Status",
                           "Week Commencing")

# for testing purposes
# prerelease_all = rbind(prerelease_all, c(6, 2024, "Experimental release"))
# prerelease_all = rbind(prerelease_all, c(6, 2024, "ad-hoc"))
# prerelease_all = rbind(prerelease_all, c(6, 2024, "ad hoc"))

# new column specifies the type of the publication
keywords <- list("ad-hoc", "ad hoc", "experimental")
prerelease_all['Type'] <- NA

for (row in 1:nrow(prerelease_all)) {
    if (is.na(prerelease_all$`Publication Title`[row])) {
        prerelease_all$Type[row] <- NA
    } else {  
        for (i in 1:length(keywords)) {
            if (grepl(keywords[i], tolower(prerelease_all$`Publication Title`[row]), fixed = TRUE)) {
                prerelease_all$Type[row] <- (gsub(" ", "-", keywords[i]))
                break
            } else {
                prerelease_all$Type[row] <- "standard"
            }
        }
    }
}

# <---------------------------------------------------- CREATE WORKBOOK ---------------------------------------------------->
govuk_link        <- c("https://www.gov.uk/search/research-and-statistics?content_store_document_type=upcoming_statistics&organisations%5B%5D=ministry-of-justice&order=release-date-oldest")
names(govuk_link) <- c("Click here to view on the gov.uk Research and Statistics calendar")
class(govuk_link) <- "hyperlink"

wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "Forward Look")
openxlsx::writeData(wb, "Forward Look","MoJ Statistics Forward Look",
                    startRow = 1)
openxlsx::writeData(wb,"Forward Look", paste("This list contains a week-by-week view of  MoJ Official and National Statistics that have been pre-announced on the gov.uk release calendar as at",
                                            format(Sys.Date(),"%d %B %Y")), startRow = 2)
openxlsx::writeData(wb,"Forward Look", govuk_link,startRow = 3)
selections <- c("Week Commencing",
                "Publication Title",
                "Publication Date",
                "Status",
                "Week",
                "Type")
openxlsx::writeData(wb,"Forward Look", select(prerelease_all, all_of(selections)), startRow = 4)

titleStyle    <- createStyle(fontSize = 14, textDecoration = "bold")
subtitleStyle <- createStyle(fontSize = 12)
linkStyle     <- createStyle(fontSize = 12, valign = "top")
evenStyle     <- createStyle(bgFill = "#b4c6e7")
oddStyle      <- createStyle(bgFill = "#d9e1f2")
hideStyleEven <- createStyle(bgFill = "#b4c6e7", fontColour = "#b4c6e7")
hideStyleOdd  <- createStyle(bgFill = "#d9e1f2", fontColour = "#d9e1f2")
border        <- createStyle(border="top", borderColour = "#FFFFFF")

conditionalFormatting(wb, "Forward Look", cols = 1:6, rows = 1:nrow(prerelease_all)+4,
                      rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=0)", style = evenStyle)
conditionalFormatting(wb, "Forward Look", cols = 1:6, rows = 1:nrow(prerelease_all)+4,
                      rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=1)", style = oddStyle)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+4,
                      rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=0,$E5=$E4)", style = hideStyleEven)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+4,
                      rule = "=AND(LEN($E5)>0,MOD(RIGHT($E5,2),2)=1,$E5=$E4)", style = hideStyleOdd)
conditionalFormatting(wb, "Forward Look", cols = 1:6, rows = 1:nrow(prerelease_all)+4,
                      rule = "=AND($E5<>$E4)", style = border)

setColWidths(wb, 1, cols = c(1:6), widths=c(18,"auto",30,10,10), hidden=c(rep(FALSE,4),TRUE))
setRowHeights(wb, 1, 3, 30)

header_st <- createStyle(fgFill = "#1F497D", textDecoration = "Bold", fontColour = "#FFFFFF")
cell_st   <- createStyle(halign = "left")

openxlsx::addStyle(wb,1,header_st,4,c(1:6))

addStyle(wb, 1, style = titleStyle, rows = 1, cols = 1)
addStyle(wb, 1, style = subtitleStyle, rows = 2, cols = 1)
addStyle(wb, 1, style = linkStyle, rows = 3, cols = 1, stack = TRUE)
addStyle(wb, 1, style = cell_st, cols = 1:6, rows = 5:nrow(prerelease_all)+4, gridExpand = TRUE, stack = TRUE)
showGridLines(wb, 1, showGridLines = FALSE)

saveWorkbook(wb, "Forward Look/Forward Look.xlsx", overwrite = TRUE, returnValue = FALSE)

