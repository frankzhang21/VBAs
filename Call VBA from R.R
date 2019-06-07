library(RDCOMClient)

# Open a specific workbook in Excel:
xlApp <- COMCreate("Excel.Application")
xlWbk <- xlApp$Workbooks()$Open("C:\\Users\\fzhang\\OneDrive - Travelzoo\\ceshi.xlsm")

# this line of code might be necessary if you want to see your spreadsheet:
# its ok to run macro without visible excel application
xlApp[['Visible']] <- TRUE 

# Run the macro called "MyMacro":
xlApp$Run("test")

# Close the workbook and quit the app:
xlWbk$Close(FALSE)# not save
xlWbk$close(TRUE) # save
xlApp$Quit()

# Release resources:
rm(xlWbk, xlApp)
gc()

