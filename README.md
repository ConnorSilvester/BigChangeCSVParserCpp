# BigChangeCSVParserCpp

This program is designed to parse a csv file into two xlsx files, these files are required by bigchange for completed jobs.

### Supported Templates So Far
### Contact Data
**Headers**
- Contact name
- Reference
- Group
- Address
- Postcode
- City
- Country
- Primary person - Mobile phone

### Job Data
**Headers**
- Job reference
- Job type
- Contact name
- Contact postcode
- Contact address
- Job contact person phone
- Planned start time
- Job contact person
- Contact group

# Setup
If you want to download from source then you will need to compile the code using the flag `-lxlsxwriter`

This lib is part of the `libxlsxwriter-dev` package and will need to be installed and imported to compile.

For Linux `g++ -o Main Main.cpp -lxlsxwriter`

For MacOS `clang++ -o Main Main.cpp -lxlsxwriter`

# Usage
To start the program just run the binary file.

To import data just place the csv file into the same dir as the project.

The code will generate a `log.txt` file, this file is important, if anything goes wrong with the program it will be logged here.

It will also generate the two excel files with the data filled in and ready to be uploaded to bigchange.

Please check both files before uploading them in case of errors in the csv data.

