install.packages("officer")
data_file <- "data/AmAus01_transcript.docx"
library(officer)
doc <- read_docx(data_file)
docx_summary(doc)
doc_df <- docx_summary(doc)
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