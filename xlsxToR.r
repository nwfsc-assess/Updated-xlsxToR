xlsxToR <- function(file, keep_sheets = NULL, header = FALSE) {
  
  require(XML)
  require(plyr)
  require(pbapply)
  
  file.copy(file, tempdir())
  new_file <- list.files(tempdir(), full.name = TRUE, pattern = basename(file))
  unzip(new_file, exdir = tempdir())
  
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
  sheet_names$id <- gsub("\\D", "", sheet_names$id)
  
  # Get column classes
  styles <- xmlToList(xmlParse(list.files(
    paste0(tempdir(), "/xl"), full.name = TRUE, pattern = "styles.xml")))
  styles <- styles$cellXfs[
    sapply(styles$cellXfs, function(x) any(names(x) == "applyNumberFormat"))]
  styles <- do.call("rbind", lapply(styles, 
    function(x) as.data.frame(as.list(x[c("applyNumberFormat", "numFmtId")]),
      stringsAsFactors = FALSE)))
  
  if(!is.null(keep_sheets)) {
    sheet_names <- sheet_names[sheet_names$name %in% keep_sheets,]
    
  }
    
  worksheet_paths <- list.files(
    paste0(tempdir(), "/xl/worksheets"), 
    full.name = TRUE, 
    pattern = paste0(
      "sheet(", 
      paste(sheet_names$id, collapse = "|"), 
      ")\\.xml$"))
  
  worksheets <- lapply(worksheet_paths, function(x) xmlRoot(xmlParse(x))[["sheetData"]])
  
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
      
      x$sheet <- sheet_names[sheet_names$id == i, "name"] 
      colnames(x) <- gsub("^value\\.", "", colnames(x))
    }
    x
  })
  worksheets <- do.call("rbind.fill", 
    worksheets[sapply(worksheets, class) == "data.frame"])
  
  entries <- xmlParse(list.files(paste0(tempdir(), "/xl"), full.name = TRUE, 
    pattern = "sharedStrings.xml$"))
  entries <- xpathSApply(entries, "//x:t", namespaces = "x", xmlValue)
  names(entries) <- seq_along(entries) - 1
    
  entries_match <- entries[match(worksheets$v, names(entries))]
  worksheets$v[worksheets$t == "s" & !is.na(worksheets$t)] <- 
    entries_match[worksheets$t == "s"& !is.na(worksheets$t)]
  worksheets$cols <- match(gsub("\\d", "", worksheets$r), LETTERS)
  worksheets$rows <- as.numeric(gsub("\\D", "", worksheets$r))
  
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