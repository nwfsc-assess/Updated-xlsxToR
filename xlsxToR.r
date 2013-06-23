xlsxToR <- function(file, keep_sheets = NULL, header = FALSE) {
  
  require(XML)
  require(plyr)
  
  suppressWarnings(file.remove(tempdir()))
  file.copy(file, tempdir())
  new_file <- list.files(tempdir(), full.name = TRUE, pattern = basename(file))
  new_file_rename <- gsub("xlsx$", "zip", new_file)
  file.rename(new_file, new_file_rename)
  
  unzip(new_file_rename, exdir = tempdir())

  # Get OS
  mac <- xmlToList(xmlParse(list.files(
    paste0(tempdir(), "/docProps"), full.name = TRUE, pattern = "app.xml")))
  mac <- grepl("Macintosh", mac$Application)
  if(mac) {
    os_origin <- "1899-12-30" # documentation says should be "1904-01-01"
  } else {
    os_origin <- "1899-12-30"
  }
  
  # Get names of sheets
  sheet_names <- xmlToList(xmlParse(list.files(
    paste0(tempdir(), "/xl"), full.name = TRUE, pattern = "workbook.xml")))
  sheet_names <- do.call("rbind", sheet_names$sheets)
  rownames(sheet_names) <- NULL
  sheet_names <- as.data.frame(sheet_names,stringsAsFactors = FALSE)
  
  # Get column classes
  styles <- xmlToList(xmlParse(list.files(
    paste0(tempdir(), "/xl"), full.name = TRUE, pattern = "styles.xml")))
  styles <- styles$cellXfs[
    sapply(styles$cellXfs, function(x) any(names(x) == "applyNumberFormat"))]
  styles <- do.call("rbind", lapply(styles, 
    function(x) as.data.frame(as.list(x[c("applyNumberFormat", "numFmtId")]),
      stringsAsFactors = FALSE)))
  
  worksheet_paths <- list.files(paste0(tempdir(), "/xl/worksheets"), 
    full.name = TRUE, pattern = "xml$")
  
  worksheets <- lapply(worksheet_paths, function(x) xmlToList(xmlParse(x))$sheetData)
  worksheets <- lapply(seq_along(worksheets), function(i) {
    x <- lapply(worksheets[[i]], function(y) {
      y <- y[names(y) == "c"]
      y <- lapply(y, function(z) {
        z <- unlist(z)
        names(z) <- gsub("\\.?attrs\\.?", "", names(z))
        as.data.frame(as.list(z), stringsAsFactors = FALSE)
      })
      do.call("rbind.fill", y)
    })
    x <- do.call("rbind.fill", x)
    x$sheet <- sheet_names[sheet_names$sheetId == i, "name"] 
    x
  })
  worksheets <- do.call("rbind.fill", 
    worksheets[sapply(worksheets, class) == "data.frame"])
  
  entries <- xmlToList(xmlParse(list.files(paste0(tempdir(), "/xl"), 
    full.name = TRUE, pattern = "sharedStrings.xml$")))
  entries <- unlist(entries)
  entries <- entries[names(entries) == "si.t"]
  names(entries) <- seq_along(entries) - 1
  
  entries_match <- entries[match(worksheets$v, names(entries))]
  worksheets$v[worksheets$t == "s" & !is.na(worksheets$t)] <- 
    entries_match[worksheets$t == "s"& !is.na(worksheets$t)]
  worksheets$cols <- match(gsub("\\d", "", worksheets$r), LETTERS)
  worksheets$rows <- as.numeric(gsub("\\D", "", worksheets$r))
  
  if(!is.null(keep_sheets)) {
    worksheets <- worksheets[sheet %in% keep_sheets,]
  }
  
  if(!any(grepl("^s$", colnames(worksheets)))) {
    worksheets$s <- NA
  }
  
  workbook <- lapply(unique(worksheets$sheet), function(x) {
    y <- worksheets[worksheets$sheet == x,]
    y_style <- as.data.frame(tapply(y$s, list(y$rows, y$cols), identity), 
      stringsAsFactors = FALSE)
    y <- as.data.frame(tapply(y$v, list(y$rows, y$cols), identity), 
      stringsAsFactors = FALSE)
    
    if(header) {
      colnames(y) <- y[1,]
      y <- y[-1,]
      y_style <- y_style[-1,]
    }
    
    y_style <- sapply(y_style, 
      function(x) ifelse(length(unique(x)) == 1, unique(x), NA))
    if(length(styles) > 0) {
      y_style <- styles$numFmtId[match(y_style, styles$applyNumberFormat)]
    }
    y_style[y_style %in% 14:17] <- "date"
    y_style[y_style %in% c(18:21, 45:47)] <- "time"
    y_style[y_style %in% 22] <- "datetime"
    y_style[is.na(y_style) & !sapply(y, function(x)any(grepl("\\D", x)))] <- "numeric"
    y_style[is.na(y_style)] <- "character"
    
    y[] <- lapply(seq_along(y), function(i) {
      switch(y_style[i],
        character = y[,i],
        numeric = as.numeric(y[,i]),
        date = as.Date(as.numeric(y[,i]), origin = os_origin),
        time = strftime(as.POSIXct(as.numeric(y[,i]), origin = os_origin), format = "%H:%M:%S"),
        datetime = as.POSIXct(as.numeric(y[,i]), origin = os_origin))
    }) 
    y 
  })
  
  if(length(workbook) == 1) {
    workbook <- workbook[[1]]
  }
  
  workbook
}