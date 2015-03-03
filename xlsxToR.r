# The MIT License (MIT)
#
# Copyright (c) 2012 Schaun Jacob Wheeler
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the right
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# Items added or fixed by John Wallace (JRW; John.Wallace@noaa.gov)

# - Put the require()'ing of packages inside the function 
# - Fixed the sorting when the number of Excel sheets is greater than 9. Before the fix, character sorting put e.g. 11-19 between 1 and 2.  
# - Added the ability to skip header lines on a sheet, either by giving a single number for all sheets or a numeric vector the length of the sheets selected.
# - ************** When counting the number of lines to skip, blank lines do not count. **************
# - Added simplify_names feature to simplify both sheet and column names (if a different simplification is wanted, search on simplify_names to find the lines to change).
# - Added the feature that keep_sheets can be a numeric vector of sheets to keep. Note that leaving keep_sheets = null results in all sheets being selected.
# - Fixed bug that broke the sheet selection i.e. keep=c(1,3).
# - Added verbose feature (for use in debugging).
# - Fixed a bug of not deleting temp files, see below.
# - Reset the row names to start at one.


xlsxToR <- function(file, keep_sheets = NULL, skip = NULL, header = TRUE, simplify_names = TRUE, verbose = FALSE) {
  
  require(XML)
  require(plyr)
  require(pbapply)

# JRW 16 Jan 2015 There was a huge bug of not deleting temp files which caused e.g. changes in Excel files not to be read into R and deleted files appearing to be read into R.
  #                 The fix below is a hack caused by MS del command not having a recursive feature, a cleaner solution can be done with a Unix/Linx rm command clone.
  #                 Also, R does not have a dir.remove() similar to the file.remove()
  #                 The unlink() during on.exit() should work for Windows and non-Windows systems (not tested), but I added the big hammer in Windows for the rare cases 
  #                     where a crash stops on.exit() from working.
  temp_dir <- file.path(tempdir(), "xlsxToRtemp")
  suppressWarnings(dir.create(temp_dir))
  suppressWarnings({ if(.Platform$OS.type == "windows") shell(paste("rmdir", gsub("/", "\\\\", temp_dir), " /s /q")) }) 
  temp_dir <- file.path(tempdir(), "xlsxToRtemp"); on.exit(unlink(temp_dir, T, T))
  dir.create(temp_dir)
  
  
  
  file.copy(file, temp_dir)
  new_file <- list.files(temp_dir, full.name = TRUE, pattern = basename(file))
  unzip(new_file, exdir = temp_dir)
  
  # Get OS
  # These lines are included because R documentation states that Excel handles 
  # date origins differently on Mac than on Windows. However, manual inspection
  # of Excel files created on Windows and Mac indicated that in fact the origin
  # is handled the same across both platforms. I've kept the original code here
  # commented out in case it can be of use in the future.
  # mac <- xmlToList(xmlParse(list.files(
  #   paste0(temp_dir, "/docProps"), full.name = TRUE, pattern = "app.xml")))
  # mac <- grepl("Macintosh", mac$Application)
  # if(mac) {
  #   os_origin <- "1899-12-30" # documentation says should be "1904-01-01"
  # } else {
  #   os_origin <- "1899-12-30"
  # }
  
  # Get names of sheets
  sheet_names <- xmlToList(xmlParse(list.files(
    paste0(temp_dir, "/xl"), full.name = TRUE, pattern = "workbook.xml")))
  sheet_names <- rbind.fill(lapply(sheet_names$sheets, function(x) {
    as.data.frame(as.list(x), stringsAsFactors = FALSE) }))
  rownames(sheet_names) <- NULL
  sheet_names <- as.data.frame(sheet_names,stringsAsFactors = FALSE)
  sheet_names$id <- gsub("\\D", "", sheet_names$id)
# JRW 09 Dec 2014; added ability to simplify sheet names
  if(simplify_names) sheet_names$name <- gsub("\\W", "", gsub("-", "_", gsub(" ", "_", sheet_names$name))) # Edit this line to change how sheet names are simplified
  if(verbose) { print(sheet_names); cat("\n\n"); flush.console()}
  
  # Get column classes
  styles <- xmlParse(list.files(
    paste0(temp_dir, "/xl"), full.name = TRUE, pattern = "styles.xml"))
  styles <- xpathApply(styles, "//x:xf[@applyNumberFormat and @numFmtId]", 
    namespaces = "x", xmlAttrs)
  styles <- lapply(styles, function(x) {
    x[grepl("applyNumberFormat|numFmtId", names(x))]})
  styles <- do.call("rbind", (lapply(styles, 
    function(x) as.data.frame(as.list(x[c("applyNumberFormat", "numFmtId")]),
      stringsAsFactors = FALSE))))
  
# JRW 09 Dec 2014; added ability to use sheet number
  if(!is.null(keep_sheets)) {                       
    if(is.numeric(keep_sheets))  
         sheet_names <- sheet_names[keep_sheets,]
    else  
          sheet_names <- sheet_names[sheet_names$name %in% keep_sheets,]
  }
  cat("\n\nSheets selected:\n"); print(sheet_names); cat("\n\n"); flush.console()  

  worksheet_paths <- list.files(
    paste0(temp_dir, "/xl/worksheets"), 
    full.name = TRUE, 
    pattern = paste0(
      "sheet(", 
      paste(sheet_names$id, collapse = "|"), 
      ")\\.xml$"))


# JRW 09 Dec 1014; more than 9 Excel datasheets sorted incorrectly because character sorting differs from numerical sorting,
   #                   e.g. '11' sorts between '1' and '2'
   # A fix to that is below:
   
   worksheet_paths <- worksheet_paths[order(as.numeric(unlist(strsplit(unlist(lapply(strsplit(unlist(lapply(strsplit(worksheet_paths, "worksheet"),  
                                             function(x) x[2])), "sheet"), function(x) x[2])), ".xml"))))]

  if(verbose) { print(worksheet_paths); cat("\n\n"); flush.console() }
  
  worksheets <- lapply(worksheet_paths, function(x) xmlRoot(xmlParse(x))[["sheetData"]])
  if(verbose) { print(seq_along(worksheets)); cat("\n\n"); print(sheet_names$id); cat("\n\n"); flush.console() }
  

  # sheet_names$id <- as.numeric(sheet_names$id)

  worksheets <- pblapply(seq_along(worksheets), function(i) {
    
    x <- xpathApply(worksheets[[i]], "//x:c", namespaces = "x", function(node) {
      c("v" = xmlValue(node[["v"]]), xmlAttrs(node))
    })
    
    if(length(x) > 0) {
      
      x_rows <- unlist(lapply(seq_along(x), function(i) rep(i, length(x[[i]]))))
      x <- unlist(x)
      
      x <- reshape(
        data.frame(
          "row" = x_rows,
          "ind" = names(x),
          "value" = x,
          stringsAsFactors = FALSE), 
        idvar = "row", timevar = "ind", direction = "wide")
# JRW 15 Jan 2015, fixed below so sheet selection works, i.e. keep=c(1,3), was: x$sheet <- sheet_names[sheet_names$id == i, "name"] 
      x$sheet <- sheet_names[sheet_names$id == sheet_names$id[i], "name"] 
      colnames(x) <- gsub("^value\\.", "", colnames(x))
# JRW 09 Dec 2014; added ability to simplify column names
      if(simplify_names) colnames(x) <- gsub("\\W", "", gsub("-", "_", gsub(" ", "_", colnames(x)))) # Edit this line to change how column names are simplified
  
    }
    x
  })
  worksheets <- do.call("rbind.fill", 
    worksheets[sapply(worksheets, class) == "data.frame"])
  
  entries <- xmlParse(list.files(paste0(temp_dir, "/xl"), full.name = TRUE, 
    pattern = "sharedStrings.xml$"))
  entries <- xpathSApply(entries, "//x:si", namespaces = "x", xmlValue)
  names(entries) <- seq_along(entries) - 1
  
  entries_match <- entries[
    match(worksheets$v[worksheets$t == "s" & !is.na(worksheets$t)], 
      names(entries))]
  worksheets$v[worksheets$t == "s" & !is.na(worksheets$t)] <- entries_match
  worksheets$cols <- match(gsub("\\d", "", worksheets$r), LETTERS)
  worksheets$rows <- as.numeric(gsub("\\D", "", worksheets$r))
    
  if(!any(grepl("^s$", colnames(worksheets)))) {
    worksheets$s <- NA
  }
  
# JRW updated 15 Jan 2015; added the ability to skip header lines on a sheet, either by giving a single number for all sheets or a numeric vector the length of the sheets selected.
   if(!is.null(skip)) { 
      N <- length(unique(worksheets$sheet))
      if(length(skip) == 1 | length(skip) == N) {
           if(length(skip) == 1) skip <- rep(skip, N)
      } else { 
           stop("\n\nThe length of 'skip' must be 1 or the number of sheets selected\n\n"); flush.console() 
      }
      if(verbose) { print(skip); cat("\n\n"); flush.console() }
   } 
 

# JRW 15 Jan 2015: Changed the 2 lines below to have an index for the skip feature, was: workbook <- lapply(unique(worksheets$sheet), function(x) { 
#      and;   y <- worksheets[worksheets$sheet == x,]
  workbook <- lapply(seq_along(unique(worksheets$sheet)), function(i) {
    
    y <- worksheets[worksheets$sheet == unique(worksheets$sheet)[i],]
    if(verbose) { print(y[1:10,]); cat("\n\n"); flush.console() }
    y_style <- as.data.frame(tapply(y$s, list(y$rows, y$cols), identity), 
      stringsAsFactors = FALSE)
    y <- as.data.frame(tapply(y$v, list(y$rows, y$cols), identity), 
      stringsAsFactors = FALSE)
    

# JRW Dec 2014:   Added code below to skip header lines
    if(!is.null(skip)) { 
      if(verbose) { print(y[1:(skip[i] + 3),]); cat("\n\n\n"); print(skip); print(i); print(skip[i]); cat("\n\n"); flush.console() }
      if(skip[i] != 0) {
        y <- y[-(1:skip[i]),]
        y_style <- y_style[-(1:skip[i]),]
    }}  

    if(verbose) { print(y[1:4,]); cat("\n\n\n"); flush.console() }

    if(header) {
      colnames(y) <- y[1,]
      y <- y[-1,]
      y_style <- y_style[-1,]
    }

    if(verbose) { print(y[1:4,]); cat("\n\n"); flush.console() }
      

    y_style <- sapply(y_style, function(x) {
      out <- names(which.max(table(x)))
      out[is.null(out)] <- NA
      out
    })
    
    if(length(styles) > 0) {
      y_style <- styles$numFmtId[match(y_style, styles$applyNumberFormat)]
    }
    
    y_style[y_style %in% 14:17] <- "date"
    y_style[y_style %in% c(18:21, 45:47)] <- "time"
    y_style[y_style %in% 22] <- "datetime"
    y_style[is.na(y_style) & !sapply(y, function(x)any(grepl("\\D", x)))] <- "numeric"
    y_style[is.na(y_style)] <- "character"
    y_style[!(y_style %in% c("date", "time", "datetime", "numeric"))] <- "character"
    
    y[] <- lapply(seq_along(y), function(i) {
      switch(y_style[i],
        character = y[,i],
        numeric = as.numeric(y[,i]),
        date = as.Date(as.numeric(y[,i]), origin = os_origin),
        time = strftime(as.POSIXct(as.numeric(y[,i]), origin = os_origin), format = "%H:%M:%S"),
        datetime = as.POSIXct(as.numeric(y[,i]), origin = os_origin))
    })
# JRW 16 Jan 2015; Reset the row names to start at one
    rownames(y) <- 1:nrow(y)
    if(verbose) { print(y[1:4,]); cat("\n\n"); flush.console() } 
    y 
  })

  if(simplify_names) {
      sheet_names$name <- gsub("\\W", "", gsub("-", "_", gsub(" ", "_", sheet_names$name)))
      for( i in 1:length(workbook)) colnames(workbook[[i]]) <- gsub("\\W", "", gsub("-", "_", gsub(" ", "_", colnames(workbook[[i]]))))
  }

  if(length(workbook) == 1) {
    workbook <- workbook[[1]]
  } else { 
    names(workbook) <- sheet_names$name
  }
  
  workbook
}
