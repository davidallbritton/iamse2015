## R source file for creating a word cloud from an Excel file
## 
## David Allbritton, 2015
##
## Install R and RStudio, place this file in the same folder as the
## Excel file containing the texts, and then click this file to run it.
## (If it does not run automatically, click "Source" after RStudio opens)
##
## The Excel file should contain one text per row in the first column, where
## a "text" could consist of one student's answer to a survey question.
## 
## Optionally, you can change the default Excel file name, minimum word frequency, 
## column number, and list of "stop words" (words to be omitted from the word cloud) here:
## 
inputFile <- "textsForWordcloud.xlsx"
columnNum <- 1
min.freq  <- 5
myStopWords <- c("case", "western", "however", "also", "actually", "something", "rather")
##
## Or after running this script once you can simply run the command
##
##   createWordCloud(inputFile, columnNum, myStopWords, min.freq)
##
## on the R command line and substitute new parameters as needed
## Examples:
##   createWordCloud("aDifferentExcelFile.xls", columnNum, myStopWords, min.freq)
##   createWordCloud(inputFile, 2, myStopWords, min.freq) #texts in column 2
##   createWordCloud(inputFile, columnNum, myStopWords, 6) #min frequency = 6
##   createWordCloud("someFile.xls", 3, myStopWords, 7) #column 3, min freq = 7


#######################################################################
## Nothing below this line needs to be changed ########################

# load or install the required R packages for reading Excel files and data manipulation:
if (!require(xlsx)) {install.packages("xlsx")}
if (!require(plyr)) {install.packages("plyr")}
# load or install R packages required specifically for doing a word cloud:
#   see: http://www.listendata.com/2014/11/create-wordcloud-with-r.html
if (!require(wordcloud)) {install.packages("wordcloud")}
if (!require(tm)) {install.packages("tm")}
if (!require(ggplot2)) {install.packages("ggplot2")}

# define a function to generate a word cloud:
doWordCloud <- function(docs ,myStopWords = c("the","a"), x.min.freq = 5){
  # arguments:
  # docs = vector of texts to use in corpus, one text per row
  # myStopWords = vector of words to exclude from the word cloud
  # min.freq = minimum word frequency for inclusion in wordcloud; default = 5
  #
  #   This funciton based on an example from:
  #    http://www.listendata.com/2014/11/create-wordcloud-with-r.html
  #   for more examples also see:  
  #    http://onertipaday.blogspot.com/2011/07/word-cloud-in-r.html
  #
  # clean up the documents vector:
  toSpace <- content_transformer(function(x,pattern) gsub(pattern," ", x))
  docs <- tm_map(docs, toSpace, "/")
  docs <- tm_map(docs, toSpace, "@")
  docs <- tm_map(docs, toSpace, "\\|")
  docs <- tm_map(docs, content_transformer(tolower))
  docs <- tm_map(docs, removeNumbers)
  docs <- tm_map(docs, removePunctuation)
  docs <- tm_map(docs, removeWords, stopwords("english"))
  # remove own custom stop words
  docs <- tm_map(docs, removeWords, myStopWords) 
  docs <- tm_map(docs, stripWhitespace)
  toString <- content_transformer(function(x, from, to) gsub(from, to, x))
  docs <- tm_map(docs, toString, "Pharm Web", "Pharmweb")
  #
  dtm <- DocumentTermMatrix(docs)
  #
  freq <- sort(colSums(as.matrix(dtm)), decreasing=TRUE)
  wf <- data.frame(word=names(freq), freq=freq)
  #
  p <- ggplot(subset(wf, freq>x.min.freq), aes(word, freq))
  p <-p+ geom_bar(stat="identity")
  p <-p+ theme(axis.text.x=element_text(angle=45, hjust=1))
  #
  wordcloud(names(freq), freq, min.freq=x.min.freq, colors=brewer.pal(6,"Dark2"),random.order=FALSE)
}

# Define a function to read in the texts and call the doWordCloud function
createWordCloud <- function (inFile = "textsForWordcloud.xls", colNum = columnNum, myStopWords = c("the","a"), minimum.freq = 5) {
  # read in the data file; one question per column, one "text" per row
  cleandata <- read.xlsx(file=inFile, 1)  #reading from sheet 1
  #
  # create vectors of texts to pass to the wordcloud generating function
  myTexts <- Corpus(VectorSource(cleandata[,colNum]))  #default is column 1
  # make the word cloud:
  doWordCloud(myTexts, myStopWords, minimum.freq)
}

# Call the above function to read in the texts and generate the word cloud:
createWordCloud(inputFile, columnNum, myStopWords, min.freq)


          
          