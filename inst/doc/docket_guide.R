## -----------------------------------------------------------------------------

library(docket)

template_path <- system.file("template_document", "Template.docx", package="docket")

print(template_path)


## -----------------------------------------------------------------------------
library(knitr)
template_screenshot <- system.file("template_document", "Document_Template.png", package="docket")
include_graphics(template_screenshot)

## -----------------------------------------------------------------------------
myDictionary <- getDictionary(template_path)

print(myDictionary)

## -----------------------------------------------------------------------------
myDictionary <- getDictionary(template_path)

#Set dictionary values
myDictionary[1,2] <- Sys.getenv("USERNAME") #Author name
myDictionary[2,2] <- as.character(Sys.Date()) # Date report created
myDictionary[3,2] <- 123
myDictionary[4,2] <- 456
myDictionary[5,2] <- 789
myDictionary[6,2] <- sum(as.numeric(myDictionary[3:5,2]))

print(myDictionary)

## -----------------------------------------------------------------------------
#check the dictionary to ensure it is valid
print(checkDictionary(myDictionary))


## -----------------------------------------------------------------------------
output_path <- paste0(normalizePath(tempdir(), winslash = "/"), "/output document.docx")

#If docket accepts the input dictionary as valid, create a filled template

if (checkDictionary(myDictionary) == TRUE) {
  docket(template_path, myDictionary, output_path)
}

print(file.exists(output_path))

## -----------------------------------------------------------------------------
print(output_path)
output_screenshot <- system.file("template_document", "Processed_Document.png", package="docket")
include_graphics(output_screenshot)

## -----------------------------------------------------------------------------
# Path to the sample template file included in the package
template_path <- system.file("batch_document", "batchTemplate.docx", package="docket")

temp_dir <- normalizePath(tempdir(), winslash = "/")
output_paths <- as.list(paste0(temp_dir, paste0("/batch document", 1:5, ".docx")))

# Create a dictionary by using the getDictionary function on the sample template file
myBatchDictionary <- getBatchDictionary(template_path, output_paths)

## -----------------------------------------------------------------------------
myBatchDictionary[2,2:ncol(myBatchDictionary)] <- Sys.getenv("USERNAME") #Author name
myBatchDictionary[3,2:ncol(myBatchDictionary)] <- as.character(Sys.Date())
myBatchDictionary[4,2:ncol(myBatchDictionary)] <- 123
myBatchDictionary[5,2:ncol(myBatchDictionary)] <- 456
myBatchDictionary[6,2:ncol(myBatchDictionary)] <- 789
myBatchDictionary[7,2:ncol(myBatchDictionary)] <- sum(as.numeric(myBatchDictionary[4:6,2]))

print(checkBatchDictionary(myBatchDictionary))

## -----------------------------------------------------------------------------
#Create multiple populated documents
if (checkBatchDictionary(myBatchDictionary) == TRUE) {
 batchDocket(template_path, myBatchDictionary)
}

#Verify documents exist
for (i in 1:length(output_paths)) {
   if (file.exists(output_paths[[i]])) {
     print(paste("docket", i, "Successfully Created"))
  }
}


## -----------------------------------------------------------------------------
if (file.exists(output_path)) {
  file.remove(output_path)
}

for (i in 1:length(output_paths)) {
   if (file.exists(output_paths[[i]])) {
     file.remove(output_paths[[i]])
  }
}

