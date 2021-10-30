# Dean Crouch
# OHS6176 Assessment Two
# 2021-10-26
# Analysis of mine atmo sampling data
#
# Agents have been renamed to shorter string labels.


# load libraries
library(xlsx)        # load XLSX into R
library(dplyr)       # useful functions like count
library(onewaytests) # contains Shapiro.Wilks test

sig_digit=2

# import data from Excel workbook
data <- read.xlsx("Data.xlsx", sheetIndex = 1)
roster_list <- read.xlsx("Data.xlsx", sheetIndex = 2)
agent_list <- read.xlsx("Data.xlsx", sheetIndex = 3)

# create list of unique SEG
SEG_list <- c(unique(data$AreaWorked))
# create DF with a separate row with a combination of each SEG and agent.
# This will ignore data is the agent isn't in the agent_list XLSX table.
result_df <- expand.grid(SEG=SEG_list, agent = agent_list$agent)
# reorder result_df by SEG then Agent.
result_df <- result_df[order(result_df$SEG),]

# add new columns for stats, with default values that set datatype.
result_df$count <- 0
result_df$max <- 0
result_df$mean <- 0
result_df$SD <- 0
result_df$shapiroLin <- ""
result_df$GM <- ""
result_df$GSD <- ""
result_df$shapiroLog <- ""
result_df$shape <- ""
result_df$P95 <- ""
result_df$WES <- ""
result_df$WES_adj <- ""
result_df$LOD <- ""
result_df$mean_WES <- ""
result_df$P95_WES <- ""


for (i in 1:length(result_df$SEG)){
  
  # TRUE/FALSE of rows that match the SEG
  index <- data$AreaWorked == result_df$SEG[i]
  
  # Set the var to be used as column name form the agent_list
  var <- as.character(result_df$agent[i])
  
  # reset the data_list to MAKE SURE there aren't any roll overs
  data_list = NULL
  
  # Get all the results with matching SEG, and from the correct column that matches the agent.
  data_subset <- get(var, data[index,])
  
  for(j in 1:length(data_subset)) {             # Using for-loop to add extracted data dataframe/vector to list
    data_list <- c(data_list,data_subset[j])
    }
  
  # only count the value if it is not NA. 6 "NA is a row with a length of 6, dont make that mistake again.
  result_df$count[i] <- sum(!is.na(data_list))
  
  if (result_df$count[i] > 0){
    result_df$max[i]  <- format(max(data_list), digits=sig_digit)
    result_df$mean[i] <- format(mean(data_list), digits=sig_digit)
    result_df$SD[i]   <- format(sd(data_list), digits=sig_digit)
    result_df$GM[i]   <- format(exp(mean(log(data_list))), digits=sig_digit)
    result_df$GSD[i]  <- format(exp(sd(log(data_list))), digits=sig_digit)   


    # WES, adj & LOD from the agent_list table into each row
    result_df$WES[i] <- subset(agent_list$WES,agent_list$agent==result_df$agent[i])
    result_df$LOD[i] <- subset(agent_list$LOD,agent_list$agent==result_df$agent[i])    
    result_df$WES_adj[i] <- format(as.numeric(result_df$WES[i])*min(subset(roster_list$adj,roster_list$SEG==result_df$SEG[i])), digits=sig_digit)
     
    if (result_df$count[i]>2){
      # check if data is normally distributed
      shapiro_result <- shapiro.test(data_list)  
      result_df$shapiroLin[i] <- format(shapiro_result$p.value,digits=sig_digit)
      if (shapiro_result$p.value > 0.05){ # set shape to linear if > 0.05
        result_df$shape[i] <- "lin"
      } 
      
      if(result_df$shapiroLin[i] <= 0.05){
        #check if data is lognormal
        shapiro_result <- shapiro.test(log(data_list))
        result_df$shapiroLog[i] <- format(shapiro_result$p.value, digits=sig_digit)
        if (shapiro_result$p.value > 0.05){
          result_df$shape[i] <- "log"
        }else{
          result_df$shape[i] <- "Non"  
        }
      }
      
      
      # Based on result_df$shape > calculate the 95th percentile
      if(result_df$shape[i]=="lin"){
        result_df$P95[i] <- format(as.numeric(result_df$mean[i])+as.numeric(result_df$SD[i])*1.645, digits=sig_digit)
      }else{
        result_df$P95[i]=format(as.numeric(result_df$GM[i])*as.numeric(result_df$GSD[i])^1.634,digits=sig_digit)
      }
      
      
      # create two columns with the mean and P95 as a proportion of the WES (after adjustment)

      print(i)
      result_df$mean_WES[i] <- format(as.numeric(result_df$mean[i])/as.numeric(result_df$WES[i]),digits=sig_digit)
      result_df$P95_WES[i]  <- format(as.numeric(result_df$P95[i])/as.numeric(result_df$WES[i]),digits=sig_digit)
      

    }
  }
}

# export the results_df as xlsx file
write.xlsx(result_df, file=paste("summary-", format(Sys.Date(),format="%Y%m%d") ,"_",format(Sys.time(),format="%H%M%S"),".xlsx",sep=""), sheetName = "Summary", 
           col.names = TRUE, row.names = TRUE, append = FALSE)





##################### Outside GSD=3 (weird) SEGs

weird_SEG <- subset(result_df, result_df$GSD >= 3, select=c(SEG, agent, GSD))


# Get all the results with matching SEG, and from the correct column that matches the agent.
MP_Sil_df  <- get("SIL", data[data$AreaWorked == "Mobile plant",])
hist(MP_Sil_df)

OPT_Sil_df <- get("SIL", data[data$AreaWorked == "Open-Cut Loading & Transport",])
hist(OPT_Sil_df)

OPT_Mg_df <- get("Mg", data[data$AreaWorked == "Open-Cut Loading & Transport",])
hist(OPT_Mg_df)

## QQ plt of variable
qqnorm(OPT_Sil_df, pch = 5, frame = FALSE)
qqline(OPT_Sil_df, col = "steelblue", lwd = 3)




MP_Sil_SD <- format(exp(sd(log(MP_Sil_df))), digits=sig_digit)
new_MP_Sil_SD <- format(exp(sd(log(which(MP_Sil_df>0.01)))), digits=sig_digit)
