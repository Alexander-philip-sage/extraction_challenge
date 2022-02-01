install.packages("writexl")
install.packages("readxl")
install.packages("pdftools")
install.packages("tidyverse")
install.packages("officer")
install.packages("glue")
install.packages("janeaustenr")
install.packages("dplyr")
install.packages("data.table")
install.packages("tm")
install.packages("tidytext")
install.packages("reprex")

library(writexl)
library(readxl)
library(pdftools)
library(tidyverse)
library(officer)
library(glue)
library(janeaustenr)
library(dplyr)
library(data.table)
library(stringr)
library(reprex)
library(tm)
library(tidytext)

Files_Path_Input <- "C:/Users/wuchaudhry/Desktop/G500/Sample_Transcripts/Transcipt_Inputs/"
Files_Path_Output <- "C:/Users/wuchaudhry/Desktop/G500/Sample_Transcripts/Transcript_Output/"

#Word Input
temp_DOC = list.files(Files_Path_Input, full.names = TRUE, pattern = ".doc")
temp_DOC
Final_Output1 <- data.table()
for(i in temp_DOC){
  PDF_Processing <- i
  Name_Full <- str_split(PDF_Processing, "_")
  first_name <- Name_Full[[1]][4]
  first_name
  last_name <- Name_Full[[1]][5]
  print(last_name)
  sample_doc <- read_docx(i)
  sample_doc <- data.table(sample_doc)
  # print(sample_doc)
  DOC_Tokenize <- tibble(line = 1, text = sample_doc$sample_doc)
  DOC_Tokenize
  DOC_Tokenize <- data.table(DOC_Tokenize %>% unnest_tokens(word, text))
  DOC_Tokenize
  #GPA Extraction 
  t1_GPA_Last <- DOC_Tokenize %>% tail(40) #%>% filter(between(word, 2, 4)) 
  t1_GPA_Last <- t1_GPA_Last[!is.na(as.numeric(as.character(t1_GPA_Last$word))),]
  t1_GPA_Last
  t1_GPA_Grade <- t1_GPA_Last[as.numeric(t1_GPA_Last$word) %inrange% c(2,4)]
  # t1_GPA_Grade <- t1_GPA_Last %>% select(word) %>% filter(word<4) %>% filter(word>2) 
  t1_GPA_Grade <- max(t1_GPA_Grade$word)
  t1_GPA_Grade
  #Hours Extraction
  t1_GPA_Hours <- t1_GPA_Last[as.numeric(word) %inrange% c(29,600)]
  t1_GPA_Hours2 <- max(unique(as.numeric(t1_GPA_Hours$word)))
  t1_GPA_Hours2
  
  data.table(first_name,last_name,t1_GPA_Grade,t1_GPA_Hours2)
  #Full Outputs: 
  Final_Output_word <-  data.frame(First_Name = first_name,Last_Name = last_name, 
                               Cumulative_GPA = t1_GPA_Grade,Cumulative_Credit_Hours = t1_GPA_Hours2,
                               'College/University' = NA)
  print(Final_Output_word)
  Final_Output1 <- rbind(Final_Output1,Final_Output_word)
  
}
print(Final_Output1)

# 
# PDF_Processing <- temp_DOC[1]
# PDF_Processing
# Name_Full <- str_split(PDF_Processing, "_")
# first_name <- Name_Full[[1]][4]
# first_name
# last_name <- Name_Full[[1]][5]
# last_name
# sample_doc <- read_docx(PDF_Processing)
# sample_doc <- data.table(sample_doc)
# sample_doc
# 
# DOC_Tokenize <- tibble(line = 1, text = sample_doc$sample_doc)
# DOC_Tokenize
# DOC_Tokenize <- data.table(DOC_Tokenize %>% unnest_tokens(word, text))
# DOC_Tokenize
# #GPA Extraction 
# t1_GPA_Last <- DOC_Tokenize %>% tail(40) #%>% filter(between(word, 2, 4)) 
# t1_GPA_Last <- t1_GPA_Last[!is.na(as.numeric(as.character(t1_GPA_Last$word))),]
# t1_GPA_Last
# t1_GPA_Grade <- t1_GPA_Last[as.numeric(t1_GPA_Last$word) %inrange% c(2,4)]
# # t1_GPA_Grade <- t1_GPA_Last %>% select(word) %>% filter(word<4) %>% filter(word>2) 
# t1_GPA_Grade <- max(t1_GPA_Grade$word)
# t1_GPA_Grade
# #Hours Extraction
# t1_GPA_Hours <- t1_GPA_Last[as.numeric(word) %inrange% c(29,600)]
# t1_GPA_Hours2 <- max(unique(as.numeric(t1_GPA_Hours$word)))
# t1_GPA_Hours2
# 
# data.table(first_name,last_name,t1_GPA_Grade,t1_GPA_Hours2)
# #Full Outputs: 
# Final_Output1 <-  data.frame(First_Name = first_name,Last_Name = last_name, 
#                              Cumulative_GPA = t1_GPA_Grade,Cumulative_Credit_Hours = t1_GPA_Hours2,
#                              'College/University' = NA)
# Final_Output1
# write_xlsx(Final_Output1,path = paste0(Files_Path_Output,'Candidate_Trancript_Analysis_Output_WORD',".xlsx"),col_names = TRUE, 
#            format_headers = TRUE,use_zip64 = FALSE)
# 



# head(sample_doc,2)
# content <- docx_summary(sample_doc)
# t2 <- data.table(Col1 = content$text)
# head(t2)
# t2_GPA <- dplyr::filter(t2, grepl('GPA:', Col1, fixed = TRUE))
# t2_GPA <- dplyr::filter(t2_GPA, !grepl('GPA Hours', Col1, fixed = TRUE))
# t2_GPA
# t2_GPA_1 <- as.numeric(gsub("[^0-9.-]", "", t2_GPA))
# t2_GPA_1
# # content %>% filter(content_type == "table cell")
# # paragraphs <- content %>% filter(content_type == "paragraph")
# 
# lst(0:4)


#PDF Input
# txt <- data.table(txt)
# txt
# txt <- pdf_text("C:/Users/wuchaudhry/Desktop/G500/Sample_Transcripts/Rutgers Transcript....pdf")
temp_PDF = list.files(Files_Path_Input, full.names = TRUE, pattern = ".pdf")
temp_PDF


remove(Final_Output2)
# Final_Output2 <- data.frame(First_Name = first_name,Last_Name = last_name, 
#                            Cumulative_GPA = t1_GPA_Grade,Cumulative_Credit_Hours = t1_GPA_Hours2,
#                            'College/University' = NA)
print(Final_Output2)
Final_Output2 <- data.table()
for (i in seq_along(temp_PDF)) {
  # i=1         # 2. sequence
  # cat(pdf_text(temp_PDF[i])) # 3. body
    PDF_Processing <- temp_PDF[i]
    PDF_Processing
    Name_Full <- str_split(PDF_Processing, "_")
    first_name <- Name_Full[[1]][4]
    first_name
    last_name <- Name_Full[[1]][5]
    last_name
    print(last_name)
    txt <- pdf_text(PDF_Processing)
    #Tokenization
    t1_GPA <-data.table(tokenize(txt))
    t4 <- tibble(line = 1, text = txt)
    t4
    tokens <- data.table(t4 %>% unnest_tokens(word, text))
    tokens
    #GPA Extraction 
    # t1_GPA_Grade <- as.numeric(gsub("[^0-9.-]", "", tokens))
    t1_GPA_Last <- tokens %>% tail(40) #%>% filter(between(word, 2, 4)) 
    t1_GPA_Last <- t1_GPA_Last[!is.na(as.numeric(as.character(t1_GPA_Last$word))),]
    t1_GPA_Last
    t1_GPA_Grade <- t1_GPA_Last[as.numeric(t1_GPA_Last$word) %inrange% c(2,4)]
    t1_GPA_Grade <- t1_GPA_Last %>% select(word) %>% filter(word<4) %>% filter(word>2) 
    t1_GPA_Grade <- max(t1_GPA_Grade$word)
    t1_GPA_Grade
    #Hours Extraction
    # t1_GPA_Hours <- tokens %>% filter(word <4) 
    t1_GPA_Hours <- t1_GPA_Last[as.numeric(word) %inrange% c(30,600)]
    t1_GPA_Hours2 <- max(unique(as.numeric(t1_GPA_Hours$word)))
    t1_GPA_Hours2
    

    Final_Output <- data.frame(First_Name = first_name,Last_Name = last_name, 
                               Cumulative_GPA = t1_GPA_Grade,Cumulative_Credit_Hours = t1_GPA_Hours2,
                               'College/University' = NA)
    write_xlsx(Final_Output,path = paste0(Files_Path_Output,first_name,"-",last_name,"_PDF",".xlsx"),col_names = TRUE, 
               format_headers = TRUE,use_zip64 = FALSE)
    #print(Final_Output)
    Final_Output2 <- rbind(Final_Output2,Final_Output)
    # Final_Output2 <- Final_Output %>% add_row(Final_Output[1])
    # Final_Output %>% full_join(Final_Output[i])

    #print(Final_Output2)    
}

# Final_Output2 <- dedupe(Final_Output2)
print(Final_Output2)
write_xlsx(Final_Output2,path = paste0(Files_Path_Output,'Candidate_Trancript_Analysis_Output_PDF',".xlsx"),col_names = TRUE, 
           format_headers = TRUE,use_zip64 = FALSE)

Final_Output3 <- full_join(Final_Output2, Final_Output1)
Final_Output3
write_xlsx(Final_Output3,path = paste0(Files_Path_Output,'Candidate_Trancript_Analysis_Output_Full_List',".xlsx"),col_names = TRUE, 
           format_headers = TRUE,use_zip64 = FALSE)






