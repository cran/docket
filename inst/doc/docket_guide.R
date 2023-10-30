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
output_path <- paste0(dirname(template_path), "/output document.docx")

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
if (file.exists(output_path)) {
  file.remove(output_path)
}

