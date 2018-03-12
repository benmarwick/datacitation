#' Create title page
#'
#' Creates text for the title and abstract page for MS Word documents.
#' \emph{This function is not exported.}
#'
#' @param x List. Meta data of the document as a result from \code{\link[yaml]{yaml.load}}.
#' @seealso \code{\link{apa6_word}}

word_title_page <- function(x) {
  # Create title page and abstract
  # Hack together tables for centered elements -.-

  apa_terms <- options("papaja.terms")[[1]]

  authors <- sapply(x$author, function(y) {
    paste0(y["name"], collapse = "")
  })

  authors[1] <- paste0(authors[1], " (", x$author[[1]]$email, ", corresponding author)")

  authors <- paste(authors,  collapse = ", ")

  affiliations <- mapply(function(x, y) c(paste0("**", x$name, "**", ", ", y["institution"])),
                         x$author, x$affiliation )
  affiliations <- paste(affiliations, collapse = "\n")

  padding <- paste0(c("\n", rep("&nbsp;", 148)), collapse = "") # Add spacer to last row
  # author_note <- paste(author_note, padding, sep = "\n")

  c(
    "\n\n"
    , paste(knitr::kable(c("&nbsp;",
                         "&nbsp;",
                         toupper(x$title),
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                           "DO NOT CITE IN ANY CONTEXT WITHOUT PERMISSION OF THE AUTHOR(S)",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                         "&nbsp;",
                           authors,
                         "&nbsp;",
                         "&nbsp;",
                           padding,
                           affiliations,
                           padding),
                         format = "pandoc",
                         align = "c"),
            collapse = "\n"),
    "##### \n"
  )
}



#' APA article (6th edition)
#'
#' Template for creating an article according to APA guidelines (6th edition) in PDF format.
#'
#' @param fig_caption Logical. Indicates if figures are rendered with captions.
#' @param number_sections Logical. Indicates if section headers are be numbered. If
#' \code{TRUE}, figure/table numbers will be of the form X.i, where X is the current first
#' -level section number, and i is an incremental number (the i-th figure/table); if
#' \code{FALSE}, figures/tables will be numbered sequentially in the document from 1, 2,
#'  ..., and you cannot cross-reference section headers in this case.
#' @param toc Logical. Indicates if a table of contents is included.
#' @param pandoc_args Additional command line options to pass to pandoc
#' @param keep_tex Logical. Keep the intermediate tex file used in the conversion to PDF.
#' @param md_extensions Character. Markdown extensions to be added or removed from the
#'  default definition or R Markdown. See the \code{\link[rmarkdown]{rmarkdown_format}} for additional details.
#' @param ... Further arguments to pass to \code{\link[bookdown]{pdf_document2}} or \code{\link[bookdown]{word_document2}}.
#' @details
#'    When creating PDF documents the YAML option \code{class} is passed to the class options of the LaTeX apa6 document class.
#'    In this case, additional options are available. Refer to the apa6 document class
#'    \href{ftp://ftp.fu-berlin.de/tex/CTAN/macros/latex/contrib/apa6/apa6.pdf}{documentation} to find out about class options
#'    such as paper size or draft watermarks.
#'
#'    When creating PDF documents the output device for figures defaults to \code{c("pdf", "postscript", "png", "tiff")},
#'    so that each figure is saved in all four formats at a resolution of 300 dpi.
#' @seealso \code{\link[bookdown]{pdf_document2}}, \code{\link[bookdown]{word_document2}}
#' @examples NULL
#' @export

apa6_pdf <- function(
  fig_caption = TRUE
  , number_sections = FALSE
  , toc = FALSE
  , pandoc_args = NULL
  , keep_tex = TRUE
  , ...
) {
  validate(fig_caption, check_class = "logical", check_length = 1)
  validate(number_sections, check_class = "logical", check_length = 1)
  validate(toc, check_class = "logical", check_length = 1)
  validate(keep_tex, check_class = "logical", check_length = 1)

  # Get APA6 template
  template <-  system.file(
    "rmarkdown", "templates", "apa6", "resources"
    , "apa6.tex"
    , package = "datacitation"
  )
  if(template == "") stop("No LaTeX template file found.")

  # Call pdf_document() with the appropriate options
  format <- bookdown::pdf_document2(
    template = template
    , fig_caption = fig_caption
    , number_sections = number_sections
    , toc = toc
    , keep_tex = keep_tex
    , pandoc_args = pandoc_args
    , ...
  )

  # Set chunk defaults
  format$knitr$opts_chunk$echo <- FALSE
  format$knitr$opts_chunk$message <- FALSE
  # format$knitr$opts_chunk$results <- "asis"
  format$knitr$opts_chunk$fig.cap <- " " # Ensures that figure environments are added
  format$knitr$opts_knit$rmarkdown.pandoc.to <- "latex"
  # format$knitr$knit_hooks$inline <- inline_numbers

  format$knitr$opts_chunk$dev <- c("pdf", "png") # , "postscript", "tiff"
  format$knitr$opts_chunk$dpi <- 300
  format$clean_supporting <- FALSE # Always keep images files

  ## Overwrite preprocessor to set CSL defaults
  saved_files_dir <- NULL

  # Preprocessor functions are adaptations from the RMarkdown package
  # (https://github.com/rstudio/rmarkdown/blob/master/R/pdf_document.R)
  # to ensure right geometry defaults in the absence of user specified values
  pre_processor <- function(metadata, input_file, runtime, knit_meta, files_dir, output_dir) {
    # save files dir (for generating intermediates)
    saved_files_dir <<- files_dir

    args <- pdf_pre_processor(metadata, input_file, runtime, knit_meta, files_dir, output_dir)

    # Set citeproc = FALSE by default to invoke ampersand filter
    if(is.null(metadata$citeproc) || metadata$citeproc) {
      metadata$citeproc <- FALSE
      assign("yaml_front_matter", metadata, pos = parent.frame())
    }

    args
  }

  format$pre_processor <- pre_processor

  if(Sys.info()["sysname"] == "Windows") {
    format$on_exit <- function() if(file.exists("_papaja_ampersand_filter.bat")) file.remove("_papaja_ampersand_filter.bat")
  }

  format
}


#' @describeIn apa6_pdf Format to create .docx-files. \code{class} parameter is ignored. \emph{This function
#'    should be considered experimental.}
#' @export

apa6_word <- function(
  fig_caption = TRUE
  , pandoc_args = NULL
  , md_extensions = NULL
  , ...
) {
  validate(fig_caption, check_class = "logical", check_length = 1)

  # Get APA6 reference file
  reference_docx <- system.file(
    "rmarkdown", "templates", "apa6", "resources"
    , "apa6_man.docx"
    , package = "datacitation"
  )
  if(reference_docx == "") stop("No .docx-reference file found.")

  # Call word_document() with the appropriate options
  format <- bookdown::word_document2(
    reference_docx = reference_docx
    , fig_caption = fig_caption
    , pandoc_args = pandoc_args
    , ...
  )

  # Set chunk defaults
  format$knitr$opts_chunk$echo <- FALSE
  format$knitr$opts_chunk$message <- FALSE
  # format$knitr$opts_chunk$results <- "asis"
  format$knitr$opts_knit$rmarkdown.pandoc.to <- "docx"
  # format$knitr$knit_hooks$inline <- inline_numbers
  # format$knitr$knit_hooks$plot <- function(x, options) {
  #   options$fig.cap <- paste("*", getOption("papaja.terms")$figure, ".* ", options$fig.cap)
  #   knitr::hook_plot_md(x, options)
  # }

  #   format$knitr$opts_chunk$dev <- c("png", "pdf", "svg", "tiff")
  #   format$knitr$opts_chunk$dpi <- 300
  format$clean_supporting <- FALSE # Always keep images files


  ## Overwrite preprocessor to set CSL defaults
  saved_files_dir <- NULL
  from_rmarkdown <- utils::getFromNamespace("from_rmarkdown", "rmarkdown")
  .from <- from_rmarkdown(fig_caption, md_extensions)

  # Preprocessor functions are adaptations from the RMarkdown package
  # (https://github.com/rstudio/rmarkdown/blob/master/R/pdf_document.R)
  # to ensure right geometry defaults in the absence of user specified values
  pre_processor <- function(metadata, input_file, runtime, knit_meta, files_dir, output_dir, from = .from) {
    # save files dir (for generating intermediates)
    saved_files_dir <<- files_dir

    args <- word_pre_processor(metadata, input_file, runtime, knit_meta, files_dir, output_dir, from)

    # Set citeproc = FALSE by default to invoke ampersand filter
    if(is.null(metadata$citeproc) || metadata$citeproc) {
      metadata$citeproc <- FALSE
      assign("yaml_front_matter", metadata, pos = parent.frame())
    }

    args
  }

  format$pre_processor <- pre_processor

  if(Sys.info()["sysname"] == "Windows") {
    format$on_exit <- function() if(file.exists("_papaja_ampersand_filter.bat")) file.remove("_papaja_ampersand_filter.bat")
  }

  format
}


# Set hook to print default numbers
inline_numbers <- function (x) {
  if (is.numeric(x)) {
    printed_number <- ifelse(
      x == round(x)
      , as.character(x)
      , printnum(x)
    )
    n <- length(printed_number)
    if(n == 1) {
      printed_number
    } else if(n == 2) {
      paste(printed_number, collapse = " and ")
    } else if(n > 2) {
      paste(paste(printed_number[1:(n - 1)], collapse = ", "), printed_number[n], sep = ", and ")
    }
  } else if(is.character(x)) x
}


# Preprocessor functions are adaptations from the RMarkdown package
# (https://github.com/rstudio/rmarkdown/blob/master/R/pdf_document.R)
# to ensure right geometry defaults in the absence of user specified values

set_csl <- function(x) {
  # Use APA6 CSL citations template if no other file is supplied
  has_csl <- function(text) {
    length(grep("^csl:.*$", text)) > 0
  }

  if (!has_csl(readLines(x, warn = FALSE))) {
    csl_template <- system.file(
      "rmarkdown", "templates", "apa6", "resources"
      , "apa6.csl"
      , package = "datacitation"
    )
    if(csl_template == "") stop("No CSL template file found.")
    return(c("--csl", rmarkdown::pandoc_path_arg(csl_template)))
  } else NULL
}

pdf_pre_processor <- function(metadata, input_file, runtime, knit_meta, files_dir, output_dir) {
  # Parse and modify YAML header
  input_text <- readLines(input_file, encoding = "UTF-8")
  yaml_params <- get_yaml_params(input_text)

  # yaml_params$author <- author_ampersand(yaml_params$author)

  ## Add modified YAML header
  yaml_delimiters <- grep("^(---|\\.\\.\\.)\\s*$", input_text)
  augmented_input_text <- c("---", yaml::as.yaml(yaml_params), "---", input_text[(yaml_delimiters[2] + 1):length(input_text)])
  writeLines(augmented_input_text, input_file, useBytes = TRUE)

  args <- NULL
  if(is.null(metadata$citeproc) || metadata$citeproc) {

    # Set CSL
    args <- set_csl(input_file)

    # Set ampersand filter
    args <- set_ampersand_filter(args, metadata$csl)
  }

  args
}

word_pre_processor <- function(metadata, input_file, runtime, knit_meta, files_dir, output_dir, from) {
  # Parse and modify YAML header
  input_text <- readLines(input_file, encoding = "UTF-8")
  yaml_params <- get_yaml_params(input_text)

  # yaml_params$author <- author_ampersand(yaml_params$author)

  ## Create title page
  yaml_delimiters <- grep("^(---|\\.\\.\\.)\\s*$", input_text)
  augmented_input_text <- c(word_title_page(yaml_params), input_text[(yaml_delimiters[2] + 1):length(input_text)])

  ## Remove abstract to avoid redundancy introduced by pandoc
  yaml_params$abstract <- NULL

  yaml_params$title <- NULL
  yaml_params$date <- NULL

  ## Add modified YAML header
  augmented_input_text <- c("---", yaml::as.yaml(yaml_params), "---", augmented_input_text)
  writeLines(augmented_input_text, input_file, useBytes = TRUE)

  args <- NULL
  if(is.null(metadata$citeproc) || metadata$citeproc) {

    # Set CSL
    args <- set_csl(input_file)

    # Set ampersand filter
    args <- set_ampersand_filter(args, metadata$csl)
  }

  # Process markdown
  process_markdown <- utils::getFromNamespace("process_markdown", "bookdown")
  process_markdown(input_file, from, args, TRUE)

  args
}


get_yaml_params <- function(x) {
  yaml_delimiters <- grep("^(---|\\.\\.\\.)\\s*$", x)

  if(length(yaml_delimiters) >= 2 &&
     (yaml_delimiters[2] - yaml_delimiters[1] > 1) &&
     grepl("^---\\s*$", x[yaml_delimiters[1]])) {
    yaml_params <- yaml::yaml.load(paste(x[(yaml_delimiters[1] + 1):(yaml_delimiters[2] - 1)], collapse = "\n"))
    yaml_params
  } else NULL
}


author_ampersand <- function(x) {
  n_authors <- length(x)
  if(n_authors >= 2) {
    if(n_authors > 2) {
      x[[n_authors]]$name <- paste("&", x[[n_authors]]$name)
      for(i in 2:n_authors) {
        x[[i]]$name <- paste(",", x[[i]]$name)
      }
    } else {
      x[[n_authors]]$name <- paste("\\ &", x[[n_authors]]$name) # Otherwise space before ampersand disappears
    }
  }
  x
}

set_ampersand_filter <- function(args, csl_file) {
  pandoc_citeproc <- utils::getFromNamespace("pandoc_citeproc", "rmarkdown")

  if(!is.null(args)) { # CSL has not been specified manually
    # Correct in-text ampersands
    filter_path <- system.file(
      "rmarkdown", "templates", "apa6", "resources"
      , "ampersand_filter.sh"
      , package = "papaja"
    )

    if(Sys.info()["sysname"] == "Windows") {
      filter_path <- gsub("\\.sh", ".bat", filter_path)
      ampersand_filter <- readLines(filter_path)
      ampersand_filter[2] <- gsub("PATHTORSCRIPT", paste0(R.home("bin"), "/Rscript.exe"), ampersand_filter[2])
      filter_path <- "_papaja_ampersand_filter.bat"
      writeLines(ampersand_filter, filter_path)
    }

    args <- c(args, "--filter", pandoc_citeproc(), "--filter", filter_path)
  } else {
    args <- c(args, "--csl", csl_file, "--filter", pandoc_citeproc())
  }

  args
}

#' Validate function input
#'
#' This function can be used to validate the input to functions. \emph{This function is not exported.}
#'
#' @param x Function input.
#' @param name Character. Name of variable to validate; if \code{NULL} variable name of object supplied to \code{x} is used.
#' @param check_class Character. Name of class to expect.
#' @param check_mode Character. Name of mode to expect.
#' @param check_integer Logical. If \code{TRUE} an object of type \code{integer} or a whole number \code{numeric} is expected.
#' @param check_NA Logical. If \code{TRUE} an non-\code{NA} object is expected.
#' @param check_infinite Logical. If \code{TRUE} a finite object is expected.
#' @param check_length Integer. Length of the object to expect.
#' @param check_dim Numeric. Vector of object dimensions to expect.
#' @param check_range Numeric. Vector of length 2 defining the expected range of the object.
#' @param check_cols Character. Vector of columns that are intended to be in a \code{data.frame}.
#'
#' @examples
#' \dontrun{
#' in_paren <- TRUE # Taken from printnum()
#' validate(in_paren, check_class = "logical", check_length = 1)
#' validate(in_paren, check_class = "numeric", check_length = 1)
#' }

validate <- function(
  x
  , name = NULL
  , check_class = NULL
  , check_mode = NULL
  , check_integer = FALSE
  , check_NA = TRUE
  , check_infinite = TRUE
  , check_length = NULL
  , check_dim = NULL
  , check_range = NULL
  , check_cols = NULL
) {
  if(is.null(name)) name <- deparse(substitute(x))

  if(is.null(x)) stop(paste("The parameter '", name, "' is NULL.", sep = ""))

  if(!is.null(check_dim) && !all(dim(x) == check_dim)) stop(paste("The parameter '", name, "' must have dimensions " , paste(check_dim, collapse=""), ".", sep = ""))
  if(!is.null(check_length) && length(x) != check_length) stop(paste("The parameter '", name, "' must be of length ", check_length, ".", sep = ""))

  if(!check_class=="function"&&any(is.na(x))) {
    if(check_NA) stop(paste("The parameter '", name, "' is NA.", sep = ""))
    else return(TRUE)
  }

  if(check_infinite && "numeric" %in% methods::is(x) && is.infinite(x)) stop(paste("The parameter '", name, "' must be finite.", sep = ""))
  if(check_integer && "numeric" %in% methods::is(x) && x %% 1 != 0) stop(paste("The parameter '", name, "' must be an integer.", sep = ""))

  for(x.class in check_class) {
    if(!methods::is(x, x.class)) stop(paste("The parameter '", name, "' must be of class '", x.class, "'.", sep = ""))
  }

  for (x.mode in check_mode) {
    if(!check_mode %in% mode(x)) stop(paste("The parameter '", name, "' must be of mode '", x.mode, "'.", sep = ""))
  }

  if(!is.null(check_cols)) {
    test <- check_cols %in% colnames(x)

    if(!all(test)) {
      stop(paste0("Variable '", check_cols[!test], "' is not present in your data.frame.\n"))
    }
  }

  if(!is.null(check_range) && any(x < check_range[1] | x > check_range[2])) stop(paste("The parameter '", name, "' must be between ", check_range[1], " and ", check_range[2], ".", sep = ""))
  TRUE
}


#' Prepare numeric values for printing
#'
#' Converts numeric values to character strings for reporting.
#'
#' @param x Numeric. Can be either a single value, vector, or matrix.
#' @param gt1 Logical. Indicates if the absolute value of the statistic can, in principal, greater than 1.
#' @param zero Logical. Indicates if the statistic can, in principal, be 0.
#' @param margin Integer. If \code{x} is a matrix, the function is applied either across rows (\code{margin = 1})
#'    or columns (\code{margin = 2}).
#' @param na_string Character. String to print if element of \code{x} is \code{NA}.
#' @param ... Further arguments that may be passed to \code{\link{formatC}}
#' @details If \code{x} is a vector, \code{digits}, \code{gt1}, and \code{zero} can be vectors
#'    according to which each element of the vector is formated. Parameters are recycled if length of \code{x}
#'    exceeds length of the parameter vectors. If \code{x} is a matrix, the vectors specify the formating
#'    of either rows or columns according to the value of \code{margin}.
#' @examples
#' printnum(1/3)
#' printnum(1/3, gt1 = FALSE)
#' printnum(1/3, digits = 5)
#'
#' printnum(0)
#' printnum(0, zero = FALSE)
#'
#' printp(0.0001)
#' @export

printnum <- function(
  x
  , gt1 = TRUE
  , zero = TRUE
  , margin = 1
  , na_string = getOption("papaja.na_string")
  , ...
) {
  if(is.null(x)) stop("The parameter 'x' is NULL. Please provide a value for 'x'")

  ellipsis <- list(...)

  validate(gt1, check_class = "logical")
  validate(zero, check_class = "logical")
  validate(margin, check_class = "numeric", check_integer = TRUE, check_length = 1, check_range = c(1, 2))
  validate(na_string, check_class = "character", check_length = 1)

  ellipsis$x <- x
  ellipsis$gt1 <- gt1
  ellipsis$zero <- zero
  ellipsis$na_string <- na_string

  ellipsis <- defaults(
    ellipsis
    , set.if.null = list(
      digits = 2
      , big.mark = ","
    )
  )

  if(!is.null(ellipsis$digits)) {
    validate(ellipsis$digits, "digits", check_class = "numeric", check_integer = TRUE, check_range = c(0, Inf))
  }

  if(length(x) > 1) {
    # print_args <- list(digits = digits, gt1 = gt1, zero = zero)
    vprintnumber <- function(i, x){
      ellipsis.i <- lapply(X = ellipsis, FUN = sel, i)
      do.call("printnumber", ellipsis.i)
    }
  }

  if(is.matrix(x) | is.data.frame(x)) {
    x_out <- apply(
      X = x
      , MARGIN = (3 - margin) # Parameters are applied according to margin
      , FUN = function(x){
        ellipsis$x <- x
        do.call("printnum", ellipsis)
      }
      # Inception!
    )

    if(margin == 2) {
      x_out <- t(x_out) # Reverse transposition caused by apply
      dimnames(x_out) <- dimnames(x)
    }

    if(!is.matrix(x_out) && is.matrix(x)) x_out <- as.matrix(x_out, ncol = ncol(x))
    if(is.data.frame(x)) x_out <- as.data.frame(x_out)

  } else if(is.numeric(x) & length(x) > 1) {
    # print_args <- lapply(print_args, rep, length = length(x)) # Recycle arguments
    x_out <- sapply(seq_along(x), vprintnumber, x)
    names(x_out) <- names(x)
  } else {
    x_out <- do.call("printnumber", ellipsis)
  }
  x_out
}


printnumber <- function(x, gt1 = TRUE, zero = TRUE, na_string = "", ...) {

  ellipsis <- list(...)
  validate(x, check_class = "numeric", check_NA = FALSE, check_length = 1, check_infinite = FALSE)
  if(is.na(x)) return(na_string)
  if(is.infinite(x)) return("$\\infty$")
  if(!is.null(ellipsis$digits)) {
    validate(ellipsis$digits, "digits", check_class = "numeric", check_integer = TRUE, check_length = 1, check_range = c(0, Inf))
  }

  validate(gt1, check_class = "logical", check_length = 1)
  validate(zero, check_class = "logical", check_length = 1)
  validate(na_string, check_class = "character", check_length = 1)
  if(!gt1 & abs(x) > 1) warning("You specified gt1 = FALSE, but passed absolute value(s) that exceed 1.")

  ellipsis <- defaults(
    ellipsis
    , set.if.null = list(
      digits = 2
      , format = "f"
      , flag = "0"
      , big.mark = ","
    )
  )

  x_out <- round(x, ellipsis$digits) + 0 # No sign if x_out == 0

  if(sign(x_out) == -1) {
    xsign <- "-"
    lt <- "> "
    gt <- "< "
  } else {
    xsign <- ""
    lt <- "< "
    gt <- "> "
  }



  if(x_out == 0 & !zero) x_out <- paste0(lt, xsign, ".", paste0(rep(0, ellipsis$digits-1), collapse = ""), "1") # Too small to report

  if(!gt1) {
    if(x_out == 1) {
      x_out <- paste0(gt, xsign, ".", paste0(rep(9, ellipsis$digits), collapse = "")) # Never report 1
    } else if(x_out == -1) {
      x_out <- paste0(lt, xsign, ".", paste0(rep(9, ellipsis$digits), collapse = "")) # Never report 1
    }
    ellipsis$x <- x_out
    x_out <- do.call("formatC", ellipsis)
    x_out <- gsub("0\\.", "\\.", x_out)
  } else {
    ellipsis$x <- x_out
    x_out <- do.call("formatC", ellipsis)
  }
  x_out
}


#' @describeIn printnum Convenience wrapper for \code{printnum} to print p-values with three decimal places.
#' @export

printp <- function(x, na_string = "") {
  validate(x, check_class = "numeric", check_range = c(0, 1))
  validate(na_string, check_class = "character", check_length = 1)

  p <- printnum(x, digits = 3, gt1 = FALSE, zero = FALSE, na_string = na_string)
  p
}
