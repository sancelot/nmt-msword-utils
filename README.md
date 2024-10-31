# nmt-msword-utils
NMT utils to extract texts to be translated from word documents 


# Requirements :

Dotnet 7.0
#with ubuntu , setup:dotnet-sdk-7.0

# build:
cd csharp/ConsoleApp1
dotnet build

This will create a binary in the bin/Debug/net7.0 folder

# usage:
##extract word texts :
ConsoleApp1 -extract_text true -input_filename worddocument.doc

## translation
one text files have been translated , thanks to this utility they can be reintroduced in the original document at the same place 

