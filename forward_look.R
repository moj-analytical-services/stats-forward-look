library(magrittr)
library(dplyr)
library(jsonlite)
library(rvest)
library(stringr)
library(openxlsx)

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

names(prerelease_all) <- c("Publication Title", 
                           "Announcement URL",
                           "Statistics Type",
                           "Department",
                           "Publication Date",
                           "Status")

openxlsx::write.xlsx(select(filter(prerelease_all,Status != "cancelled"),
                            c("Publication Title","Publication Date","Status")),
                     "Excel.xlsx")

# create workbook
wb <- createWorkbook()

# add Excel sheet
addWorksheet(wb, "Forward_Look")

# create style, in this case bold header
header_st <- createStyle(textDecoration = "Bold")

# Write data with header style defined above
writeData(wb, "Forward_Look", 
          select(filter(prerelease_all,Status != "cancelled"),
                                     c("Publication Title","Publication Date","Status")),
          headerStyle = header_st)

setColWidths(wb,1,cols = c(1:3),widths="auto")

# save to .xlsx file
saveWorkbook(wb, paste0("Forward Look_",Sys.Date(),".xlsx"), overwrite = TRUE)


