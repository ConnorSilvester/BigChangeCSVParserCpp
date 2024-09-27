# BigChangeCSVParserCpp

This program is designed to parse a csv file into two xlsx files, these files are required by bigchange for completed jobs.

The input data is designed to be from the AnyJunk portal, in there format, see below for me details.

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

For MacOS `clang++ -std=c++17 -o Main Main.cpp -lxlsxwriter`

For MacOS Static Link `clang++ -std=c++17 -o Main Main.cpp path/to/libxlsxwriter.a -lz -lxml2`

Library is usually installed in `/usr/local/Cellar/libxlsxwriter/1.1.8/lib/libxlsxwriter.a`

# Usage
To start the program just run the binary file.

To import data just place the csv file into the same dir as the project.

The code will generate a `log.txt` file, this file is important, if anything goes wrong with the program it will be logged here.

It will also generate the two excel files with the data filled in and ready to be uploaded to bigchange.

Please check both files before uploading them in case of errors in the csv data.

# Input Data
The input data should be in the format of the Anyjunk portal.

`Collection Reference,Customer Name,Collect At,Postcode,Address,Contact Name`

Although the program does parse and look for the correct data its not guaranteed to work with all data.

An example row might looks like this.

`27825195215,Company Name Lmt,16/09/2024 10:00 - 14:00,S99 FFE,"United Kingdom, England, Manchester, Bury",James - 01234567896`

Notice how the address is surrounded by `""` this is intentional to allow the address itself to contain commas,
however this means that this is now a requirement in the input data.

If any of the input data is missing you will get an error in the `log.txt` file that's generated.
