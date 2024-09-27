# BigChangeCSVParserCpp

This program is designed to parse a csv file into two XLSX files, these files are required by bigchange for completed jobs.

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
To compile from source, you will need the `libxlsxwriter` library, you will also need to include the flag `-lxlsxwriter`.

This library is part of the `libxlsxwriter-dev` package.

This can be installed with `sudo apt install libxlsxwriter-dev` or `brew install libxlsxwriter`

For Linux `g++ -o Main Main.cpp -lxlsxwriter`

For MacOS `clang++ -std=c++17 -o Main Main.cpp -lxlsxwriter`

For MacOS Static Link `clang++ -std=c++17 -o Main Main.cpp path/to/libxlsxwriter.a -lz -lxml2`

Once installed, you can find the library files (like `libxlsxwriter.a`) in the Homebrew directory, typically `/usr/local/Cellar/libxlsxwriter/`.

To find info on the package if its not at that location you can run `brew info libxlsxwriter` to find the location.

# Usage
To start the program just run the binary file with `./Main`.

To read data place the csv file into the same dir as the executable.

The code will generate a `log.txt` file, this file is for error logging.

It will also generate the two excel files with the data filled in and ready to be uploaded to bigchange.

Ensure you review both Excel files before uploading, as input errors may affect the output.

# Input Data
The input data should follow the format of the Anyjunk portal.

`Collection Reference,Customer Name,Collect At,Postcode,Address,Contact Name`

Although the program does parse and look for the correct data its not guaranteed to work with all data.

An example row might looks like this.

`27825195215,Company Name Lmt,16/09/2024 10:00 - 14:00,S99 FFE,"United Kingdom, England, Manchester, Bury",James - 01234567896`

Note: The address is surrounded by `""` this is to allow the address itself to contain commas, this is a requirement in the input data.

If any of the input data is missing you will get an error message in the `log.txt` file that's generated.
