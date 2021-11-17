# Libraries
library(tidyquant)
library(tidyverse)
library(timetk)
library(openxlsx)


# Draws stock data from three individual industries (Tech, Heath, & Retail) 
# Calculates portfolio value, portfolio cost, and positioning 
# time_series_analysis graphs of the stocks by industry
# snapshots and archives the data into a folder, to analyze 3 industries over the course of four months 


# Draws stock data from three individual industries (Tech) 
# Table 1 
Tech_data_tbl <- c("AAPL","GOOG","NFLX","NVDA") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
tech_stocks = data.frame(Tech_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
tech_stocks$Units <- round(runif(1, 1, 5))
tech_stocks$Price_Bought <- (runif(1, 400, 700))
tech_stocks$Portfolio_Cost <- tech_stocks$Price_Bought * tech_stocks$Units
tech_stocks$Portfolio_Price <- tech_stocks$Units * tech_stocks$adjusted
tech_stocks$Portfolio_Revenue <- tech_stocks$Portfolio_Price - tech_stocks$Portfolio_Cost
tech_stocks

Tech_position = sum(tech_stocks$Portfolio_Revenue)
Tech_position

# tech_plot_table - for graph 
tech_plot_table <- select(tech_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
tech_plot <- tech_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
tech_plot

# Archive the data into a folder to create snapshots in time of your portfolio, and stock performance
#create workbook
wb <- createWorkbook()

# add Worksheet
addWorksheet(wb, sheetName = "Stock_Anysis_Tech")

# add plot
print(tech_plot)
wb %>% insertPlot(sheet = "Stock_Anysis_Tech", startCol = "G", startRow = 3)

#add Data 
writeDataTable(wb, sheet = "Stock_Anysis_Tech", x = tech_plot_table)

# save workbook
saveWorkbook(wb, "005_excel_workbook/Stock_Anysis_Tech.xlsx", overwrite = TRUE)

# Health Industry
# Table 2 
Health_data_tbl <- c("TDOC","PFE","MRNA","CVS") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
health_stocks = data.frame(Health_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
health_stocks$Units <- round(runif(1, 2, 5))
health_stocks$Price_Bought <- (runif(1, 100, 130))
health_stocks$Portfolio_Cost <-  health_stocks$Price_Bought *  health_stocks$Units
health_stocks$Portfolio_Price <-  health_stocks$Units *  health_stocks$adjusted
health_stocks$Portfolio_Revenue <-  health_stocks$Portfolio_Price -  health_stocks$Portfolio_Cost
health_stocks

Health_position = sum(health_stocks$Portfolio_Revenue)
Health_position

# health_plot_table - for graph 
health_plot_table <- select(health_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
health_plot <- health_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
health_plot

# Archive the data into a folder to create snapshots in time of your portfolio, and stock performance
#create workbook
#wb <- createWorkbook()

# add Worksheet
addWorksheet(wb, sheetName = "Stock_Anysis_Health")

# add plot
print(health_plot)
wb %>% insertPlot(sheet = "Stock_Anysis_Health", startCol = "G", startRow = 3)

#add Data 
writeDataTable(wb, sheet = "Stock_Anysis_Health", x = health_plot_table)

# save workbook
saveWorkbook(wb, "005_excel_workbook/Stock_Anysis_Health.xlsx", overwrite = FALSE)



# Retail Indsutry 
# Table 3 
Retail_data_tbl <- c("LULU","TGT","WMT","CRI") %>%
  tq_get(from = "2021-07-30", to = "2021-10-30")
retail_stocks = data.frame(Retail_data_tbl)

# Calculates portfolio value, portfolio cost, and positioning 
retail_stocks$Units <- round(runif(1, 2, 10))
retail_stocks$Price_Bought <- (runif(1, 100, 195))
retail_stocks$Portfolio_Cost <- retail_stocks$Price_Bought *  retail_stocks$Units
retail_stocks$Portfolio_Price <- retail_stocks$Units *  retail_stocks$adjusted
retail_stocks$Portfolio_Revenue <- retail_stocks$Portfolio_Price -  retail_stocks$Portfolio_Cost
retail_stocks

Retail_position = sum(retail_stocks$Portfolio_Revenue)
Retail_position 

# retail_plot_table - for graph 
retail_plot_table <- select(retail_stocks, symbol, date, adjusted, Portfolio_Revenue)

#graph
retail_plot <- retail_plot_table %>%
  group_by(symbol) %>%
  plot_time_series(date, adjusted, 
                   .facet_ncol = 2, 
                   .facet_scales = "free",
                   .interactive = FALSE)
retail_plot

# Archive the data into a folder to create snapshots in time of your portfolio, and stock performance
#create workbook
#wb <- createWorkbook()

# add Worksheet
addWorksheet(wb, sheetName = "Stock_Anysis_Retail")

# add plot
print(retail_plot)
wb %>% insertPlot(sheet = "Stock_Anysis_Retail", startCol = "G", startRow = 3)

#add Data 
writeDataTable(wb, sheet = "Stock_Anysis_Retail", x = retail_plot_table)

# save workbook
saveWorkbook(wb, "005_excel_workbook/Stock_Anysis_Retail.xlsx", overwrite = FALSE)




