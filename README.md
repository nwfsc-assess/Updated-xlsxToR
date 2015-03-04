# Updated-xlsxToR
I have made substantial additions and bug fixes to the xlsxToR R function. A list of additions and bug fixes are inside the function. An example of usage is: 

test<-xlsxToR("TrawlSurveyDataPackage_Canary_ExploreReducedSurvey_2014.xlsx", keep=c(2,4,6,8), skip=c(0,5,5,5)) 

Where the first argument is the Excel file's path, the second is those sheets you want to read into R, and the third is the number of header lines to skip on top of the file, not counting blank lines. The result is a R list if more than one sheet is selected or just the table if one sheet is selected. By default, my addition to simplify names of both sheet and column names is set to TRUE. If some other simplification is wanted, the function can be edited (I put a comment on the lines to edit).
