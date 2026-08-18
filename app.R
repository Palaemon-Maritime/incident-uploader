# ==============================================================================
#  Palaemon Maritime — Incident Uploader
#
#  Upload the ListOfIncidents .xls export and this app will:
#    1. convert it to JSON matching the schema of Global_Incidents.json
#    2. compare the CONVERTED JSON against what is already in the repo
#    3. append only the incidents that are not already there
#    4. commit the result to GitHub
#    5. record the upload in upload_log.json so the running total survives
#
#  Auth: the GitHub token is read from the GITHUB_PAT environment variable,
#        set in the shinyapps.io dashboard. It is never written in this file
#        and never touches the user's machine.
# ==============================================================================

library(shiny)
library(readxl)
library(jsonlite)
library(stringr)
library(httr)
library(base64enc)
library(openssl)

# ---- Repository configuration ------------------------------------------------

GH_OWNER     <- "PalaemonMaritime"
GH_REPO      <- "json.api"
GH_BRANCH    <- "main"
GH_DATA_PATH <- "Global_Incidents.json"
GH_LOG_PATH  <- "upload_log.json"

PREFERRED_SHEET <- "FullList"

# Field order in the output JSON, matching the existing dataset.
# The Excel export also has a "Boarded?" column that the JSON does not use;
# it is deliberately absent here, so it gets dropped.
FIELD_ORDER <- c(
  "Date", "Ship Name", "Ship Type", "IMO No.", "Area",
  "Latitude", "Longitude", "Incident details",
  "Consequences for crew etc", "Action taken by master/crew",
  "Reported?", "Reported to...", "Reporting State",
  "Coastal State Action Taken", "MSC/Circ",
  "Latitude_dd", "Longitude_dd"
)

# ---- Small helpers -----------------------------------------------------------

or_else <- function(value, fallback) {
  if (is.null(value) || length(value) == 0) return(fallback)
  if (length(value) == 1 && is.na(value[[1]])) return(fallback)
  value
}

# A value is missing if it is NA, NULL, or an empty/whitespace string.
# The literal string "NA" is NOT missing — it appears throughout the real
# dataset as a meaningful value in "Coastal State Action Taken".
is_blank <- function(x) {
  if (is.null(x)) return(TRUE)
  if (length(x) == 0) return(TRUE)
  if (all(is.na(x))) return(TRUE)
  if (is.character(x) && !nzchar(trimws(x))) return(TRUE)
  FALSE
}

as_text <- function(x) {
  if (is_blank(x)) return(NULL)
  out <- trimws(as.character(x))
  if (!nzchar(out)) return(NULL)
  out
}

# "1° 11.80' N" -> 1.1967   |   "103° 25.10' W" -> -103.4183
dms_to_decimal <- function(x) {
  if (is_blank(x)) return(NULL)
  m <- str_match(
    as.character(x),
    "([0-9]+(?:\\.[0-9]+)?)\\s*\u00b0\\s*([0-9]+(?:\\.[0-9]+)?)\\s*'\\s*([NSEWnsew])"
  )
  if (is.na(m[1, 1])) return(NULL)
  value <- as.numeric(m[1, 2]) + as.numeric(m[1, 3]) / 60
  if (toupper(m[1, 4]) %in% c("S", "W")) value <- -value
  round(value, 4)
}

as_date_string <- function(x) {
  if (is_blank(x)) return(NULL)
  if (inherits(x, "Date"))    return(format(x, "%Y-%m-%d"))
  if (inherits(x, "POSIXct")) return(format(as.Date(x), "%Y-%m-%d"))
  if (is.numeric(x))          return(format(as.Date(x, origin = "1899-12-30"), "%Y-%m-%d"))

  s <- trimws(as.character(x))
  # Only the unambiguous ISO form is accepted. Left to guess, as.Date() reads
  # "04/07/2026" as year 4 and writes "4-07-20" without complaint, and a
  # silently mangled date is worse than one passed through untouched. Anything
  # else is kept exactly as written, where it is visible and correctable.
  parsed <- tryCatch(suppressWarnings(as.Date(s, format = "%Y-%m-%d")),
                     error = function(e) NA)
  if (!is.na(parsed)) return(format(parsed, "%Y-%m-%d"))
  s
}

# IMO numbers are strings with no decimal tail: 9324291.0 -> "9324291"
as_imo <- function(x) {
  if (is_blank(x)) return(NULL)
  if (is.numeric(x)) return(format(round(as.numeric(x)), scientific = FALSE, trim = TRUE))
  s <- sub("\\.0+$", "", trimws(as.character(x)))
  if (!nzchar(s)) return(NULL)
  s
}

as_int <- function(x) {
  if (is_blank(x)) return(NULL)
  n <- suppressWarnings(as.numeric(x))
  if (is.na(n)) return(as_text(x))
  as.integer(round(n))
}

as_flag <- function(x) {
  if (is_blank(x)) return(NULL)
  if (is.logical(x)) return(unname(x[[1]]))
  s <- tolower(trimws(as.character(x)))
  if (s %in% c("true", "yes", "y", "1"))  return(TRUE)
  if (s %in% c("false", "no", "n", "0")) return(FALSE)
  NULL
}

# ---- Excel -> incident objects -----------------------------------------------

row_to_incident <- function(row) {
  pick <- function(column) if (column %in% names(row)) row[[column]][[1]] else NA

  incident <- list()
  put <- function(key, value) {
    if (!is.null(value)) incident[[key]] <<- value
  }

  put("Date",                        as_date_string(pick("Date")))
  put("Ship Name",                   as_text(pick("Ship Name")))
  put("Ship Type",                   as_text(pick("Ship Type")))
  put("IMO No.",                     as_imo(pick("IMO No.")))
  put("Area",                        as_text(pick("Area")))
  put("Latitude",                    as_text(pick("Latitude")))
  put("Longitude",                   as_text(pick("Longitude")))
  put("Incident details",            as_text(pick("Incident details")))
  put("Consequences for crew etc",   as_text(pick("Consequences for crew etc")))
  put("Action taken by master/crew", as_text(pick("Action taken by master/crew")))
  put("Reported?",                   as_flag(pick("Reported?")))
  put("Reported to...",              as_text(pick("Reported to...")))
  put("Reporting State",             as_text(pick("Reporting State")))
  put("Coastal State Action Taken",  as_text(pick("Coastal State Action Taken")))
  put("MSC/Circ",                    as_int(pick("MSC/Circ")))
  put("Latitude_dd",                 dms_to_decimal(pick("Latitude")))
  put("Longitude_dd",                dms_to_decimal(pick("Longitude")))

  incident[intersect(FIELD_ORDER, names(incident))]
}

excel_to_incidents <- function(path) {
  # readxl's own errors are about zip archives and file signatures, which mean
  # nothing to someone who just picked the wrong file.
  unreadable <- function(e) {
    stop("That file could not be opened as a spreadsheet. Make sure it is the ",
         ".xls download from the incident portal and that it isn't still open ",
         "in Excel.", call. = FALSE)
  }

  sheets <- tryCatch(excel_sheets(path), error = unreadable)
  sheet  <- if (PREFERRED_SHEET %in% sheets) PREFERRED_SHEET else sheets[1]

  df <- tryCatch(read_excel(path, sheet = sheet), error = unreadable)

  if (nrow(df) == 0) {
    stop("That sheet has no rows in it. Check the export and try again.", call. = FALSE)
  }
  if (!("Date" %in% names(df)) || !("Ship Name" %in% names(df))) {
    stop("This doesn't look like a ListOfIncidents export \u2014 no 'Date' or ",
         "'Ship Name' column. Upload the file as downloaded, without editing ",
         "the headers.", call. = FALSE)
  }

  incidents <- lapply(seq_len(nrow(df)), function(i) row_to_incident(df[i, ]))
  # Drop entirely empty rows; trailing blanks are common in these exports.
  incidents[vapply(incidents, function(x) length(x) > 0, logical(1))]
}

# ---- Duplicate detection -----------------------------------------------------

# Strings reaching this function come from two places that mark encoding
# differently: readxl returns UTF-8-marked strings, jsonlite returns unmarked
# ones. R treats those as unequal even when the underlying bytes are identical,
# which silently breaks matching on any value containing a degree sign or a
# curly apostrophe. Declaring both as UTF-8 and then reducing to printable
# ASCII makes the comparison depend on content alone.
normalise <- function(x) {
  if (is.null(x) || length(x) == 0) return("")
  s <- as.character(x)[1]
  if (is.na(s) || !nzchar(s)) return("")
  Encoding(s) <- "UTF-8"
  s <- gsub("[[:space:]]+", " ", s)
  s <- gsub("[^\x20-\x7E]", "", s, perl = TRUE, useBytes = TRUE)
  tolower(trimws(s))
}

# Primary key is Date + Ship Name + IMO.
#
# Many real records have no IMO and a ship name of "Name Withheld". For those,
# Date + Ship Name alone would collide across genuinely different incidents, so
# the decimal position and the opening of the narrative are used instead. The
# decimal coordinates are used rather than the "13° 36.80' N" strings because
# numbers carry no encoding, so they cannot fail to match for textual reasons.
incident_key <- function(incident) {
  imo  <- normalise(incident[["IMO No."]])
  stem <- paste(normalise(incident[["Date"]]),
                normalise(incident[["Ship Name"]]), sep = "|")

  if (nzchar(imo)) {
    paste(stem, imo, sep = "|")
  } else {
    coord <- function(field) {
      value <- incident[[field]]
      if (is.null(value) || length(value) == 0 || is.na(value[[1]])) return("")
      formatC(round(as.numeric(value[[1]]), 4), format = "f", digits = 4)
    }
    paste(stem,
          coord("Latitude_dd"), coord("Longitude_dd"),
          substr(normalise(incident[["Incident details"]]), 1, 300),
          sep = "|")
  }
}

# Keys are matched with %in% on a character vector rather than looked up in an
# environment. Environment names go through native-encoding translation, which
# fails on the degree signs and curly apostrophes that fill this dataset.
split_new_and_duplicate <- function(incoming, existing) {
  seen <- if (length(existing) == 0) {
    character(0)
  } else {
    vapply(existing, incident_key, character(1))
  }

  keep <- logical(length(incoming))
  for (i in seq_along(incoming)) {
    key <- incident_key(incoming[[i]])
    if (key %in% seen) {
      keep[i] <- FALSE          # already in the dataset
    } else {
      keep[i] <- TRUE
      seen <- c(seen, key)      # also catches repeats within a single upload
    }
  }

  list(new = incoming[keep], duplicates = sum(!keep))
}

# ---- GitHub ------------------------------------------------------------------

# ---- Authentication ----------------------------------------------------------
#
# Two modes, chosen by which variables are set on the host.
#
#   GitHub App (preferred, no expiry)
#     GITHUB_APP_ID   the App's numeric ID
#     GITHUB_APP_KEY  the App's private key, the whole .pem file contents
#   The private key does not expire. The app exchanges it for an installation
#   token that lasts an hour and is renewed automatically, so nothing needs
#   renewing by hand, ever.
#
#   Personal access token (fallback)
#     GITHUB_PAT
#   Simpler to set up, but an organisation caps fine-grained tokens at 366 days
#   by default, so this mode needs a diarised renewal. The app reads the expiry
#   date GitHub reports and warns on screen well before it lapses.

auth_state <- new.env(parent = emptyenv())

base64url <- function(x) {
  if (is.character(x)) x <- charToRaw(x)
  chartr("+/", "-_", sub("=+$", "", base64encode(x)))
}

app_jwt <- function(app_id, pem) {
  # Environment variables often arrive with the newlines flattened.
  pem <- gsub("\\\\n", "\n", pem)
  key <- openssl::read_key(pem)

  now     <- as.integer(Sys.time())
  header  <- base64url('{"alg":"RS256","typ":"JWT"}')
  payload <- base64url(sprintf('{"iat":%d,"exp":%d,"iss":"%s"}',
                               now - 60, now + 540, app_id))
  signing_input <- paste(header, payload, sep = ".")
  signature <- openssl::signature_create(charToRaw(signing_input),
                                         hash = openssl::sha256, key = key)
  paste(signing_input, base64url(signature), sep = ".")
}

installation_token <- function(app_id, pem) {
  cached <- auth_state$installation
  if (!is.null(cached) && cached$expires > Sys.time() + 300) return(cached$value)

  jwt <- tryCatch(app_jwt(app_id, pem), error = function(e) {
    stop("The GitHub App private key could not be read. Check that ",
         "GITHUB_APP_KEY holds the entire .pem file, including the BEGIN and ",
         "END lines.", call. = FALSE)
  })

  jwt_headers <- add_headers(
    Authorization = paste("Bearer", jwt),
    Accept = "application/vnd.github+json",
    `X-GitHub-Api-Version` = "2022-11-28",
    `User-Agent` = "palaemon-incident-uploader"
  )

  found <- GET(sprintf("https://api.github.com/repos/%s/%s/installation",
                       GH_OWNER, GH_REPO), jwt_headers)
  if (status_code(found) == 404) {
    stop(sprintf("The GitHub App is not installed on %s/%s. Install it on that ",
                 "repository and grant Contents: Read and write.",
                 GH_OWNER, GH_REPO), call. = FALSE)
  }
  if (status_code(found) >= 300) {
    stop(sprintf("GitHub rejected the App credentials (HTTP %d). Check ",
                 "GITHUB_APP_ID matches the key.", status_code(found)),
         call. = FALSE)
  }

  minted <- POST(sprintf("https://api.github.com/app/installations/%s/access_tokens",
                         content(found, as = "parsed")$id), jwt_headers)
  if (status_code(minted) >= 300) {
    stop(sprintf("Could not obtain an installation token (HTTP %d).",
                 status_code(minted)), call. = FALSE)
  }

  body <- content(minted, as = "parsed")
  auth_state$installation <- list(
    value   = body$token,
    expires = as.POSIXct(body$expires_at, format = "%Y-%m-%dT%H:%M:%SZ", tz = "UTC")
  )
  body$token
}

github_token <- function() {
  app_id <- Sys.getenv("GITHUB_APP_ID")
  app_key <- Sys.getenv("GITHUB_APP_KEY")
  if (nzchar(app_id) && nzchar(app_key)) {
    auth_state$mode <- "app"
    return(installation_token(app_id, app_key))
  }

  pat <- Sys.getenv("GITHUB_PAT")
  if (nzchar(pat)) {
    auth_state$mode <- "pat"
    return(pat)
  }

  stop("No GitHub credentials are configured for this app. Set GITHUB_APP_ID ",
       "and GITHUB_APP_KEY, or GITHUB_PAT, in the hosting dashboard.",
       call. = FALSE)
}

github_headers <- function() {
  add_headers(
    Authorization = paste("Bearer", github_token()),
    Accept = "application/vnd.github+json",
    `X-GitHub-Api-Version` = "2022-11-28",
    `User-Agent` = "palaemon-incident-uploader"
  )
}

# GitHub reports a personal access token's expiry date in a response header.
# Recording it is what stops the token from dying without warning.
note_token_expiry <- function(response) {
  header <- headers(response)[["github-authentication-token-expiration"]]
  if (is.null(header) || !nzchar(header)) return(invisible(NULL))
  parsed <- suppressWarnings(as.Date(substr(trimws(header), 1, 10)))
  if (!is.na(parsed)) auth_state$pat_expires <- parsed
  invisible(NULL)
}

# A short line for the interface describing how the app is authenticating and,
# for a token that expires, how long is left.
credential_status <- function() {
  if (identical(auth_state$mode, "app")) {
    return(list(level = "ok",
                text = "Authenticated as a GitHub App. Credentials renew automatically."))
  }
  expiry <- auth_state$pat_expires
  if (is.null(expiry)) {
    return(list(level = "ok",
                text = "Authenticated with an access token that does not expire."))
  }
  days <- as.integer(expiry - Sys.Date())
  if (days < 0) {
    list(level = "bad",
         text = sprintf("The access token expired on %s. Uploads will fail until it is replaced.",
                        format(expiry, "%d %B %Y")))
  } else if (days <= 45) {
    list(level = "warn",
         text = sprintf("The access token expires on %s, in %d days. Replace it before then or uploads will stop.",
                        format(expiry, "%d %B %Y"), days))
  } else {
    list(level = "ok",
         text = sprintf("Access token valid until %s (%d days).",
                        format(expiry, "%d %B %Y"), days))
  }
}

github_url <- function(path) {
  sprintf("https://api.github.com/repos/%s/%s/contents/%s", GH_OWNER, GH_REPO, path)
}

github_read <- function(path) {
  res <- GET(github_url(path), query = list(ref = GH_BRANCH), github_headers())
  note_token_expiry(res)

  if (status_code(res) == 404) {
    return(list(exists = FALSE, sha = NULL, text = NULL))
  }
  if (status_code(res) %in% c(401, 403)) {
    stop("GitHub rejected the token. It may have expired or lost access to ",
         "the repository.", call. = FALSE)
  }
  if (status_code(res) >= 300) {
    stop(sprintf("Could not read %s from GitHub (HTTP %d).", path,
                 status_code(res)), call. = FALSE)
  }

  body <- content(res, as = "parsed", type = "application/json")
  text <- rawToChar(base64decode(gsub("[\r\n]", "", body$content)))
  Encoding(text) <- "UTF-8"

  list(exists = TRUE, sha = body$sha, text = text)
}

github_write <- function(path, text, sha, message) {
  payload <- list(
    message = message,
    content = base64encode(charToRaw(enc2utf8(text))),
    branch  = GH_BRANCH
  )
  if (!is.null(sha)) payload$sha <- sha

  res <- PUT(github_url(path), body = payload, encode = "json", github_headers())

  if (status_code(res) == 409) {
    stop("The file changed on GitHub while this upload was in progress. ",
         "Try again.", call. = FALSE)
  }
  if (status_code(res) >= 300) {
    detail <- tryCatch(content(res, as = "parsed")$message,
                       error = function(e) NULL)
    stop(sprintf("GitHub refused the commit (HTTP %d)%s", status_code(res),
                 if (!is.null(detail)) paste0(": ", detail) else "."),
         call. = FALSE)
  }
  invisible(TRUE)
}

# simplifyVector = FALSE is required. Without it jsonlite flattens the dataset
# into a data frame and every record gains null fields it never had.
parse_incidents <- function(text) {
  if (is.null(text) || !nzchar(trimws(text))) return(list())
  parsed <- fromJSON(text, simplifyVector = FALSE)
  if (!is.list(parsed)) list() else parsed
}

# pretty = 2 and digits = 4 match the formatting already in the repo, so the
# commit diff shows only the appended records rather than the whole file.
serialise_incidents <- function(incidents) {
  as.character(toJSON(incidents, pretty = 2, auto_unbox = TRUE, digits = 4))
}

# ---- Upload history ----------------------------------------------------------

read_history <- function() {
  file <- tryCatch(github_read(GH_LOG_PATH), error = function(e) NULL)
  if (is.null(file) || !isTRUE(file$exists)) {
    return(list(entries = list(), sha = NULL))
  }
  entries <- tryCatch(fromJSON(file$text, simplifyVector = FALSE),
                      error = function(e) list())
  if (!is.list(entries)) entries <- list()
  list(entries = entries, sha = file$sha)
}

append_history <- function(entry) {
  current <- read_history()
  entries <- c(current$entries, list(entry))
  github_write(
    GH_LOG_PATH,
    as.character(toJSON(entries, pretty = 2, auto_unbox = TRUE)),
    current$sha,
    sprintf("Update upload_log.json - %s", format(Sys.Date(), "%Y-%m-%d"))
  )
  entries
}

# ---- Interface ---------------------------------------------------------------

app_css <- "
:root {
  --ink:      #10161d;
  --ink-soft: #56636f;
  --hairline: #d8dee4;
  --paper:    #f7f8f9;
  --surface:  #ffffff;
  --signal:   #1f6f5c;
  --warn:     #8a5a1c;
  --stop:     #9b2c2c;
}
body {
  background: var(--paper);
  color: var(--ink);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
  font-size: 15px;
  line-height: 1.55;
}
.wrap { max-width: 640px; margin: 0 auto; padding: 56px 24px 80px; }
.eyebrow {
  font-size: 11px; letter-spacing: .16em; text-transform: uppercase;
  color: var(--ink-soft); margin: 0 0 6px;
}
h1 { font-size: 27px; font-weight: 600; letter-spacing: -.01em; margin: 0 0 8px; }
.lede { color: var(--ink-soft); margin: 0 0 32px; }
.panel {
  background: var(--surface); border: 1px solid var(--hairline);
  border-radius: 3px; padding: 24px;
}
.target {
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12.5px; color: var(--ink-soft);
  border-top: 1px solid var(--hairline); margin-top: 24px; padding-top: 16px;
}
.target span { color: var(--ink); }
.panel .form-group { margin-bottom: 18px; }
.btn-go {
  background: var(--ink); color: #fff; border: 0; border-radius: 3px;
  padding: 11px 22px; font-size: 15px; font-weight: 500; width: 100%;
}
.btn-go:hover:enabled, .btn-go:focus:enabled { background: #263441; color: #fff; }
.btn-go:disabled { opacity: .4; }
.note { border-left: 2px solid; padding: 12px 0 12px 14px; margin-top: 22px; }
.note.ok   { border-color: var(--signal); color: var(--signal); }
.note.none { border-color: var(--warn);   color: var(--warn); }
.note.bad  { border-color: var(--stop);   color: var(--stop); }
.credential {
  font-size: 12.5px; color: var(--ink-soft); margin-top: 20px;
}
.ledger { width: 100%; border-collapse: collapse; font-size: 14px; }
.ledger th {
  text-align: left; font-size: 11px; letter-spacing: .1em; text-transform: uppercase;
  color: var(--ink-soft); font-weight: 500;
  border-bottom: 1px solid var(--hairline); padding: 0 12px 8px 0;
}
.ledger td {
  padding: 9px 12px 9px 0; border-bottom: 1px solid var(--paper);
  font-variant-numeric: tabular-nums;
}
.ledger td.num, .ledger th.num { text-align: right; padding-right: 0; }
.ledger tr:last-child td { border-bottom: 0; }
.ledger .stamp {
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 13px;
}
.tally {
  display: flex; justify-content: space-between; align-items: baseline;
  border-top: 1px solid var(--hairline); margin-top: 4px; padding-top: 14px;
}
.tally .label {
  font-size: 11px; letter-spacing: .12em; text-transform: uppercase;
  color: var(--ink-soft);
}
.tally .value { font-size: 22px; font-weight: 600; font-variant-numeric: tabular-nums; }
.modal-header, .modal-footer { border-color: var(--hairline); }
.modal-title { font-size: 18px; font-weight: 600; }
"

ui <- fluidPage(
  tags$head(
    tags$title("Incident uploader \u2014 Palaemon Maritime"),
    tags$style(HTML(app_css))
  ),
  div(
    class = "wrap",
    p(class = "eyebrow", "Palaemon Maritime"),
    h1("Incident uploader"),
    p(class = "lede",
      "Upload the ListOfIncidents export. New incidents are added to the ",
      "dataset; anything already recorded is left alone."),
    div(
      class = "panel",
      uiOutput("picker"),
      actionButton("publish", "Add to dataset", class = "btn-go"),
      uiOutput("status"),
      uiOutput("credentials"),
      div(class = "target",
          "Writing to ", span(paste0(GH_OWNER, "/", GH_REPO)),
          " \u00b7 ", span(GH_BRANCH), " \u00b7 ", span(GH_DATA_PATH))
    )
  )
)

# ---- Server ------------------------------------------------------------------

server <- function(input, output, session) {

  state <- reactiveValues(status = NULL, picker_version = 0)

  # Re-rendering the input with a new nonce is how the chosen file is cleared
  # after a successful upload, without pulling in shinyjs.
  output$picker <- renderUI({
    state$picker_version
    fileInput(
      "workbook", "Incident export",
      accept = c(".xls", ".xlsx",
                 "application/vnd.ms-excel",
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      buttonLabel = "Choose file",
      placeholder = "No file selected",
      width = "100%"
    )
  })

  count_of <- function(n, word) {
    paste0(format(n, big.mark = ","), " ", word, if (n == 1) "" else "s")
  }

  render_history <- function(entries) {
    if (length(entries) == 0) {
      return(p(class = "lede", style = "margin:0;", "No uploads recorded yet."))
    }

    rows <- lapply(rev(entries), function(e) {
      tags$tr(
        tags$td(class = "stamp", or_else(e$timestamp, "\u2014")),
        tags$td(class = "num", format(or_else(e$added,   0), big.mark = ",")),
        tags$td(class = "num", format(or_else(e$skipped, 0), big.mark = ",")),
        tags$td(class = "num", format(or_else(e$total,   0), big.mark = ","))
      )
    })

    total_added <- sum(vapply(entries,
                              function(e) as.numeric(or_else(e$added, 0)),
                              numeric(1)))
    in_dataset  <- or_else(entries[[length(entries)]]$total, 0)

    tagList(
      tags$table(
        class = "ledger",
        tags$thead(tags$tr(
          tags$th("Upload"),
          tags$th(class = "num", "Added"),
          tags$th(class = "num", "Skipped"),
          tags$th(class = "num", "In dataset")
        )),
        tags$tbody(rows)
      ),
      div(class = "tally",
          span(class = "label", "Incidents in dataset"),
          span(class = "value", format(in_dataset, big.mark = ","))),
      div(class = "tally", style = "border-top:0; padding-top:6px;",
          span(class = "label", "Added across all uploads"),
          span(class = "value", format(total_added, big.mark = ",")))
    )
  }

  observeEvent(input$publish, {
    state$status <- NULL

    upload <- input$workbook
    if (is.null(upload)) {
      state$status <- list(kind = "bad", text = "Choose an incident export first.")
      return()
    }

    outcome <- withProgress(message = "Working", value = 0, {
      tryCatch({

        incProgress(0.15, detail = "Reading the workbook")
        incoming <- excel_to_incidents(upload$datapath)
        if (length(incoming) == 0) {
          stop("No incidents found in that file.", call. = FALSE)
        }

        incProgress(0.25, detail = "Fetching the current dataset")
        remote   <- github_read(GH_DATA_PATH)
        existing <- parse_incidents(remote$text)

        incProgress(0.20, detail = "Checking for duplicates")
        split <- split_new_and_duplicate(incoming, existing)

        if (length(split$new) == 0) {
          list(kind = "none", added = 0, skipped = split$duplicates,
               total = length(existing))
        } else {
          combined <- c(existing, split$new)

          incProgress(0.25, detail = "Committing to GitHub")
          github_write(
            GH_DATA_PATH,
            serialise_incidents(combined),
            remote$sha,
            sprintf("Update Global_Incidents.json - %s",
                    format(Sys.Date(), "%Y-%m-%d"))
          )

          incProgress(0.15, detail = "Recording the upload")
          entry <- list(
            timestamp = format(Sys.time(), "%Y-%m-%d %H:%M", tz = "UTC"),
            file      = upload$name,
            added     = length(split$new),
            skipped   = split$duplicates,
            total     = length(combined)
          )
          history <- tryCatch(append_history(entry), error = function(e) NULL)

          list(kind = "ok", added = length(split$new),
               skipped = split$duplicates, total = length(combined),
               history = history, entry = entry)
        }
      },
      error = function(e) list(kind = "bad", text = conditionMessage(e)))
    })

    if (identical(outcome$kind, "bad")) {
      state$status <- list(kind = "bad", text = outcome$text)
      return()
    }

    if (identical(outcome$kind, "none")) {
      state$status <- list(
        kind = "none",
        text = paste0("Nothing to add \u2014 all ",
                      count_of(outcome$skipped, "incident"),
                      " in that file are already in the dataset.")
      )
      return()
    }

    state$status <- list(
      kind = "ok",
      text = paste0("\u2713 Success \u2014 ", count_of(outcome$added, "incident"),
                    " added, ", count_of(outcome$skipped, "duplicate"),
                    " skipped.")
    )

    # If the history file could not be written, the dataset commit still
    # succeeded — show this upload on its own rather than an empty table.
    entries <- outcome$history
    if (is.null(entries)) entries <- list(outcome$entry)

    showModal(modalDialog(
      title = "Upload complete",
      render_history(entries),
      footer = actionButton("dismiss", "Close", class = "btn-go",
                            style = "width:auto;"),
      easyClose = FALSE
    ))
  })

  observeEvent(input$dismiss, {
    removeModal()
    state$status <- NULL
    state$picker_version <- state$picker_version + 1  # clears the chosen file
  })

  # Checked when the app opens, so an expiring or broken credential is visible
  # before someone tries to upload rather than after it fails.
  output$credentials <- renderUI({
    check <- tryCatch({
      github_read(GH_DATA_PATH)
      credential_status()
    }, error = function(e) list(level = "bad", text = conditionMessage(e)))

    if (identical(check$level, "ok")) {
      return(div(class = "credential", check$text))
    }
    div(class = paste("note", if (identical(check$level, "warn")) "none" else "bad"),
        tags$strong(check$text))
  })

  output$status <- renderUI({
    s <- state$status
    if (is.null(s)) return(NULL)
    div(class = paste("note", s$kind), tags$strong(s$text))
  })
}

shinyApp(ui, server)
