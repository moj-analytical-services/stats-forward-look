# <------------------------------------------------------- LIBRARIES ------------------------------------------------------->


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
library(lubridate)
options(lubridate.week.start = 1)


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
    html_text(trim=TRUE)
  pub_attribs      <- process_html(search_url, ".gem-c-document-list__item")
  
  pub_attribs <- t(as.data.frame(lapply(pub_attribs,get_info)))
  pub_attribs <- as.data.frame(pub_attribs)
  
  rownames(pub_attribs) <- NULL
  names(pub_attribs)    <- c("publication.type",
                             "publication.department",
                             "publication.date.temp",
                             "publication.status")
  
  publication_desc <- character()
  for(j in publication_url){
    publication_desc_t <- process_html(paste0("https://www.gov.uk", j), ".govuk-grid-row") %>% html_nodes(".gem-c-lead-paragraph") %>%
      html_text(trim=TRUE)
    publication_desc <- c(publication_desc, publication_desc_t)
  }
  
  prerelease <- as.data.frame(cbind(publication_name,publication_url, publication_desc, pub_attribs))
  
  if (i == 1) {
    prerelease_all <- prerelease
  } else {
    prerelease_all <- rbind(prerelease_all,prerelease)
  }
  
}

prerelease_all_t <- prerelease_all %>% mutate(long_title = sub("[[:space:]]*[,:[:digit:]].*$", "", publication_name))

for (i in seq_along(prerelease_all_t$long_title)) {
  if (grepl("(His|Her|HM).*Prison.*Probation Service", prerelease_all_t$long_title[i], ignore.case = TRUE)) {
    rest <- sub("(?i)(His|Her|HM).*Prison.*Probation Service", "", prerelease_all_t$long_title[i], perl = TRUE)
    prerelease_all_t$long_title[i] <- paste("HMPPS", trimws(rest))
  }
}


#Combine forward look with publication leads
lookup_wb <- read.xlsx("Stats Publication Leads.xlsx") %>% select(long_title, `Publication.Month(s)`, `Lead.Contact`, G6, G7, `Justice.Data`, Mailbox)
prerelease_all_t <- prerelease_all_t %>% 
  mutate(long_title=str_squish(str_to_lower(long_title))) %>%
  left_join(lookup_wb %>% mutate(long_title=str_squish(str_to_lower(long_title))), by="long_title") 

prerelease_all_t$publication.date.temp <- stringr::str_remove(prerelease_all_t$publication.date.temp," 9:30am")
prerelease_all_t$publication.date.temp <- stringr::str_remove(prerelease_all_t$publication.date.temp," 9:03am")
prerelease_all_t$publication.date.temp <- stringr::str_remove(prerelease_all_t$publication.date.temp," 10:00am")

prerelease_all2 <- prerelease_all_t %>%
  mutate(publication.date = 
           case_when(
             grepl("^[[:digit:]]+", publication.date.temp) == TRUE ~ publication.date.temp,
             grepl("February", publication.date.temp) == TRUE & grepl("2024|2028|2032|2036|2040", publication.date.temp) == TRUE ~ paste0("29 ", publication.date.temp),
             grepl("February", publication.date.temp) == TRUE & grepl("2024|2028|2032|2036|2040", publication.date.temp) == FALSE ~ paste0("28 ", publication.date.temp),
             grepl("April|June|September|November", publication.date.temp) == TRUE ~ paste0("30 ", publication.date.temp),
             TRUE ~ paste0("31 ", publication.date.temp)
         ),
         publication.date=dmy(publication.date),
         Week_commences = floor_date(publication.date, "week", week_start=1)) %>%
         #Week = week(dmy(publication.date)),
         #Year = year(dmy(publication.date)))
  select(-publication.date.temp)

start_date <- floor_date(Sys.Date() - years(1), "week", week_start=1)
end_date <- Sys.Date() + years(3)

dates <- tibble(
  Week_commences = seq(start_date, end_date, by="week"),
  Week = week(Week_commences),
  Year = year(Week_commences),
  Month = month(Week_commences, label=TRUE, abbr=FALSE)
)

prerelease_weeks <- dates %>%
  full_join(prerelease_all2, by="Week_commences") %>%
  arrange(Week_commences)

rowvector <- seq_len(nrow(prerelease_weeks))
current_week <- floor_date(Sys.Date(), unit="week", week_start=1)

start_index <- which(prerelease_weeks$Week_commences==current_week)
if(length(start_index)==0) {
  start_index <- min(rowvector[!is.na(prerelease_weeks$publication.date)])
} else{
  start_index <- min(start_index)
}
end_index <- max(rowvector[!is.na(prerelease_weeks$publication.date)])

prerelease_all <- prerelease_weeks[start_index:end_index, ]

#Add in month headers for readability 
prerelease_all <- prerelease_all %>%
  mutate(
    Month = factor(
      Month, 
      levels=c("January", "February", "March", "April", "May", "June", "July", 
               "August", "September", "October", "November", "December"),
      ordered=TRUE
    )
  ) %>%
  group_by(Year, Month) %>%
  group_split() %>%
  lapply(function(g) {
    month_chr <- as.character(unique(g$Month))
    year_val <- unique(g$Year)
    header_txt <- paste0(month_chr, " ", year_val)
    g <- g %>% mutate(Week_commences = format(Week_commences, "%a %d %b %Y"))
    m_header <- tibble(
      Week_commences = header_txt,
      Year = year_val,
      Month = month_chr
    )
    bind_rows(m_header, g)
  }) %>%
  bind_rows() %>%
  select(-c(Year, Month, long_title)) %>%
  mutate(publication.date=format(as.Date(publication.date, format="%d %b %Y"), "%a %d %b %Y"))


names(prerelease_all) <- c("Week Commencing",
                           "Week",
                           "Publication Title", 
                           "Announcement URL",
                           "Description",
                           "Statistics Type",
                           "Department",
                           "Status",
                           "Usual publication month(s)",
                           "Lead contact",
                           "Grade 6",
                           "Grade 7",
                           "Justice Data",
                           "Mailbox",
                           "Publication Date")

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
modifyBaseFont(wb, fontSize = 9, fontColour = "#000000", fontName = "Arial")
openxlsx::writeData(wb, "Forward Look","MoJ Statistics Forward Look",
                    startRow = 1)
openxlsx::writeData(wb,"Forward Look", paste("This list contains a week-by-week view of regular MoJ Official Statistics releases that have been pre-announced on the gov.uk release calendar. The list is updated every week on a Friday."), startRow = 2)
openxlsx::writeData(wb,"Forward Look", govuk_link,startRow = 3)
openxlsx::writeData(wb, "Forward Look", paste("Last updated: ", format(Sys.Date(),"%A %d %B %Y")), startRow=4)

selections <- c("Week Commencing",
                "Publication Title",
                "Publication Date",
                "Status",
                "Week",
                "Type",
                "Usual publication month(s)", 
                "Lead contact",
                "Grade 7",
                "Grade 6",
                "Mailbox",
                "Justice Data")
openxlsx::writeData(wb,"Forward Look", select(prerelease_all, all_of(selections)), startRow = 6)

arialStyle    <- createStyle(fontName="Arial")
titleStyle    <- createStyle(fontSize = 16, textDecoration = "bold")  
subtitleStyle <- createStyle(fontSize = 12)
linkStyle     <- createStyle(fontSize = 12, valign = "top")
#evenStyle     <- createStyle(bgFill = "#b4c6e7")
#oddStyle      <- createStyle(bgFill = "#d9e1f2")
hideStyleEven <- createStyle(bgFill = "#FFFFFF", fontColour = "#FFFFFF")
hideStyleOdd  <- createStyle(bgFill = "#FFFFFF", fontColour = "#FFFFFF")
border        <- createStyle(border="top", borderColour = "#000000")
border_left   <- createStyle(border="left", borderColour= "#000000")
wrap_style    <- createStyle(wrapText = TRUE, valign="center")
m_titleStyle  <- createStyle(fontSize = 10, textDecoration = "bold", bgFill="#D9D9D9", valign="center")
bold_st       <- createStyle(textDecoration = "bold", fontSize=9)
bold_st2       <- createStyle(textDecoration = "bold", fontSize=12)
conf          <- createStyle(bgFill = "#C6EFCE", fontColour = "#006100")
prov          <- createStyle(bgFill = "#FFEB9C", fontColour = "#9C5700")
canc          <- createStyle(bgFill = "#FFC7CE", fontColour = "#9C0006")

#conditionalFormatting(wb, "Forward Look", cols = 1:7, rows = 1:nrow(prerelease_all)+6,
#                      rule = "=AND(LEN($E7)>0,MOD(RIGHT($E7,2),2)=0)", style = evenStyle)
#conditionalFormatting(wb, "Forward Look", cols = 1:7, rows = 1:nrow(prerelease_all)+6,
#                      rule = "=AND(LEN($E7)>0,MOD(RIGHT($E7,2),2)=1)", style = oddStyle)
conditionalFormatting(wb, "Forward Look", cols = 1:12, rows = 7:(nrow(prerelease_all)+7),
                      rule = '=AND($A7<>$A6)', style = border)
conditionalFormatting(wb, "Forward Look", cols=1:13, rows=7:(nrow(prerelease_all)+6),
                      rule = '=LEFT($A7, 3)="MON"', style=border_left)
conditionalFormatting(wb, "Forward Look", cols=1:12, rows=7:(nrow(prerelease_all)+6),
                      rule = '=LEFT($A7,3)<>"MON"', style=m_titleStyle, stack=TRUE)
conditionalFormatting(wb, "Forward Look", cols=2, rows=7:(nrow(prerelease_all)+6), 
                      rule = '=AND($B7<>"")', style=bold_st)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+6,
                      rule = "=AND(LEN($E7)>0,MOD(RIGHT($E7,2),2)=0,$E7=$E6)", style = hideStyleEven)
conditionalFormatting(wb, "Forward Look", cols = 1, rows = 1:nrow(prerelease_all)+6,
                      rule = "=AND(LEN($E7)>0,MOD(RIGHT($E7,2),2)=1,$E7=$E6)", style = hideStyleOdd)
conditionalFormatting(wb, "Forward Look", cols=4, rows=7:(nrow(prerelease_all)+6),
                      rule = '=AND($D7="confirmed")', style=conf)
conditionalFormatting(wb, "Forward Look", cols=4, rows=7:(nrow(prerelease_all)+6),
                      rule = '=AND($D7="provisional")', style=prov)
conditionalFormatting(wb, "Forward Look", cols=4, rows=7:(nrow(prerelease_all)+6),
                      rule = '=AND($D7="cancelled")', style=canc)

setColWidths(wb, 1, cols = c(1:12), widths=c(25,80,25,10,10,10,30,30,30,20,45,15), hidden=c(rep(FALSE,4),TRUE, TRUE, rep(FALSE,6)))
setRowHeights(wb, 1, 3, 30)

hdr_rows <- which(!grepl("Mon", prerelease_all$`Week Commencing`)) + 6 
week_rows <- which(grepl("Mon", prerelease_all$`Week Commencing`) & !grepl("Offender management statistics", prerelease_all$`Publication Title`)) + 6
setRowHeights(wb, 1, c(6, hdr_rows), 20)
setRowHeights(wb, 1, week_rows, 20)

header_st <- createStyle(fgFill = "#1D609D", textDecoration = "Bold", fontSize=10, fontColour = "#FFFFFF", valign="center")
cell_st   <- createStyle(halign = "left", valign="center")

addStyle(wb, 1, header_st,6,c(1:12))
addStyle(wb, 1, style = titleStyle, rows = 1, cols = 1)
addStyle(wb, 1, style = subtitleStyle, rows = 2, cols = 1)
addStyle(wb, 1, style = linkStyle, rows = 3, cols = 1, stack = TRUE)
addStyle(wb, 1, style = bold_st2, rows = 4, cols=1, stack=TRUE)
addStyle(wb, 1, style = cell_st, cols = 1:12, rows = 1:nrow(prerelease_all)+6, gridExpand = TRUE, stack = TRUE)
addStyle(wb, 1, style = wrap_style, rows = 1:(nrow(prerelease_all)+6), cols = 8:10, gridExpand = TRUE, stack=TRUE)
addStyle(wb, 1, style= border_left, rows = 7:(nrow(prerelease_all)+6), cols=13, gridExpand = TRUE, stack=TRUE)

showGridLines(wb, 1, showGridLines = FALSE)

addFilter(wb, 1, rows=6, cols=1:12)

saveWorkbook(wb, "Forward Look/Forward Look.xlsx", overwrite = TRUE, returnValue = FALSE)
