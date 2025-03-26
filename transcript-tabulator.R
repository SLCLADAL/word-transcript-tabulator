install.packages("officer")
data_file <- "data/AmAus01_transcript.docx"
library(officer)
doc <- read_docx(data_file)
docx_summary(doc)
doc_df <- docx_summary(doc)

doc_df[1:10,]

split_metadata <- function(doc_summary, split_at){
  return(list(doc_summary[1:split_at,], doc_summary[split_at:nrow(doc_summary),]))
}

# How do we handle blocks of blank lines?
guess_transcript_start <- function(doc_summary){
  for (i in 1:nrow(doc_summary)){
    row = doc_summary[i,]
    # Header ends with first blank line
    if (trimws(row$text) == '') {
      return(i)
    }
    # Header ends with first list paragraph style (numbered list?)
    else if (!is.na(row$style_name) & row$style_name == 'List Paragraph'){
      return(i)
    }
  }
  # If we don't match any other rules, set the header to be zero lines long
  # at the top of the file.
  return(1)
}

extract_speaker_codes <- function(doc_summary){
  speaker_match = '^[[:alpha:]]+?(?=:)'
  matches = regexpr(speaker_match, doc_summary$text, perl = TRUE)
  doc_summary$speaker_code <- regmatches(doc_summary$text, matches) 
  return(doc_summary)
}

extract_speaker_codes(loaded_docs[1])


# Processing turns:
# Extract speaker (if any)
# Extract timecodes (if any)
# If no speaker?
# Separate the speaker and the text - leave timecodes in place for interpretation



lapply(doc_df$text, strsplit, ':')

doc_files = list.files("data", full.names=TRUE, pattern="*.docx")

loaded_docs <- lapply(lapply(doc_files[1:10], read_docx), docx_summary)
transcripts_starts <- lapply(loaded_docs, guess_transcript_start)

doc_files[9]

out <- split_metadata(doc_df, guess_transcript_start(doc_df))
meta = out[[1]]
transcript = out[[2]]

# Metadata/header detection:
# Seems like it's possible *most* of the time
# The other cases will be annoying - work a bit on how do we show people that?
# 

# Assumption: dozens of word docs uploaded - can manually verify things
# like header locations.

# Next session:
# Keep working on the metadata split across all the files.

# If this is an iterative process, how do we record decisions made about things
# like metadata/transcript split points?


## Output columns
## Text
## Turn number
## Line number or flag for overlapping? (diff to turn number?)
## Speaker ID (local or global?)
## Source file
## Time code

## Speaker metadata (consistency across transcripts?)

doc_df$level==1

## Work out how to split metadata and how do we get user to verify (cut at a certain point, cut by formatting styles). Could be a function that just tests whether there are style/formatting changes that might signify start of transcript.

## Take a look at time code on their own line with no turn 

## Lines with no speakers could also have different formatting

## Next step - write metadata splitter (split into header rows, transcript)
## Then - work on transcript rows
## Open question - do we want to do this for Michael's transcript properly or make it more general? 

library(openxlsx)

write_xlsx(list(pivoted, to_pivot), 'test.xlsx')
