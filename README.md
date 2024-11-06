# nmt-msword-utils
NMT utils to extract texts to be translated from word documents 

This utility can too be used to extract texts from word files to be vectorized and used with LLM.

Once translated, word documents can be reintroduced at the same place using this utility 
in order to translate the document.



# Requirements :
This software needs Microsoft Open XML SDK

Dotnet 7.0  
#with ubuntu , setup: 
dotnet-sdk-7.0  

# build:
cd csharp/ConsoleApp1  
dotnet build  

This will create a binary in the bin/Debug/net7.0 folder  

# usage:
##extract word texts :
nmt-word-util -extract_text true -input_filename worddocument.doc  

cli possible args :
extract_texts : extract texts to a text file from a word document

translate : will translate document 

input_filename [inpt.docx]  
output_filename [output_docx]  


## translation
one text files have been translated , thanks to this utility they can be reintroduced in the original document at the same place 

