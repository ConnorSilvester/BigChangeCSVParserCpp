#include <xlsxwriter.h> //libxlsxwriter-dev
#include <string>
#include <sstream>
#include <iostream>
#include <filesystem>
#include <fstream>
#include <vector>
#include <algorithm>
#include <cctype>
#include <map>
#include <random>
#include <chrono>
#include <iomanip>


std::string get_current_time_HH_MM_SS(){
    using namespace std::chrono;

    auto now = system_clock::now();
    
    std::time_t now_c = system_clock::to_time_t(now);

    std::tm now_tm = *std::localtime(&now_c);

    std::ostringstream oss;
    oss << std::put_time(&now_tm, "%H:%M:%S");

    std::string result = "[" + oss.str() + "] ";

    return result;
}

struct logger_t{
    std::vector<std::string> lines;

    void log(std::string line){
        lines.push_back(std::string(get_current_time_HH_MM_SS() + line));
    }

    void write_to_file(){
        std::string file_name = "log.txt";
        std::ofstream out_file(file_name);

        if (!out_file.is_open()) {
            std::cerr << "Error: Could not open file " << file_name << " for writing." << std::endl;
            return;
        }

        for (const auto& line : lines) {
            out_file << line << std::endl;
        }

        out_file.close();
    }
};



static const std::string FILE_EXTENSION = ".xlsx";
static logger_t logger;



std::string read_file(const std::string& file_path){
    std::ifstream file(file_path);

    std::stringstream buffer;
    buffer << file.rdbuf();
    std::string file_contents = buffer.str();

    file.close();

    return file_contents;
}

std::vector<std::string> split(const std::string& str, char delimiter){
    std::vector<std::string> tokens;
    std::string token;
    std::stringstream ss(str);

    while (std::getline(ss, token, delimiter)){
        tokens.push_back(token);
    }

    return tokens;
}

std::string extract_single_quoted_string(const std::string& input){
    size_t start = input.find('"');
    if (start == std::string::npos){
        return "";  // No opening quote found, return an empty string
    }

    size_t end = input.find('"', start + 1);
    if (end == std::string::npos){
        return "";  // No closing quote found, return an empty string
    }

    return input.substr(start + 1, end - start - 1);
}

std::string trim(const std::string& str){
    size_t start = str.find_first_not_of(" \t\n\r\f\v");
    size_t end = str.find_last_not_of(" \t\n\r\f\v");

    if (start == std::string::npos || end == std::string::npos){
        return ""; // Return an empty string if no non-whitespace characters are found
    }

    return str.substr(start, end - start + 1);
}

std::string extract_phone_number(const std::string& input){
    // Find the last occurrence of '-'
    size_t pos = input.rfind('-');
    
    // Check if '-' was found and it is not the last character
    if (pos != std::string::npos && pos + 1 < input.length()){
        // Extract the substring from the position after the last '-'
        std::string phone_number = input.substr(pos + 1);
        phone_number = trim(phone_number);
        if (phone_number == std::string("00000000000")){
            phone_number = "01234567891";
            logger.log("00000000000 Changed to 01234567891");
        }
        if (phone_number.length() != 11){
            phone_number = "01234567891";
            logger.log(std::string(phone_number + " Changed to 01234567891"));
        }
        return phone_number;
    }


    // If no '-' found or it's at the end of the string, return an empty string
    return "01234567891";
}

std::string take_all_before(const std::string& input, const char delimiter){
    // Find the last occurrence of '-'
    size_t pos = input.rfind(delimiter);
    
    // Check if '-' was found
    if (pos != std::string::npos) {
        // Extract the substring before the last '-'
        return input.substr(0, pos);
    }
    
    // If no '-' found, return the original string (or handle it as needed)
    return input;
}

std::string remove_numbers_from_string(const std::string& input){
    std::string result = input;
    result.erase(std::remove_if(result.begin(), result.end(), [](char c) {
        return std::isdigit(static_cast<unsigned char>(c));
    }), result.end());
    return result;
}

int get_random_number(int lower, int higher){
    std::random_device rd;
    std::mt19937 gen(rd());

    std::uniform_int_distribution<> distr(lower, higher);

    return distr(gen);
}


const std::map<std::string, std::string> LOCATION_MAP = {
    {"AB", "Aberdeen"},
    {"AL", "St Albans"},
    {"B",  "Birmingham"},
    {"BA", "Bath"},
    {"BB", "Blackburn"},
    {"BD", "Bradford"},
    {"BH", "Bournemouth"},
    {"BL", "Bolton"},
    {"BN", "Brighton"},
    {"BR", "Bromley"},
    {"BS", "Bristol"},
    {"CA", "Carlisle"},
    {"CB", "Cambridge"},
    {"CF", "Cardiff"},
    {"CH", "Chester"},
    {"CM", "Chelmsford"},
    {"CO", "Colchester"},
    {"CR", "Croydon"},
    {"CT", "Canterbury"},
    {"CV", "Coventry"},
    {"CW", "Crewe"},
    {"DA", "Dartford"},
    {"DE", "Derby"},
    {"DG", "Dumfries"},
    {"DH", "Durham"},
    {"DL", "Darlington"},
    {"DN", "Doncaster"},
    {"DT", "Dorchester"},
    {"DY", "Dudley"},
    {"E",  "London (East)"},
    {"EC", "London (Central)"},
    {"EH", "Edinburgh"},
    {"EN", "Enfield"},
    {"EX", "Exeter"},
    {"FK", "Falkirk"},
    {"FY", "Blackpool"},
    {"G",  "Glasgow"},
    {"GL", "Gloucester"},
    {"GU", "Guildford"},
    {"HA", "Harrow"},
    {"HD", "Huddersfield"},
    {"HG", "Harrogate"},
    {"HP", "Hemel Hempstead"},
    {"HR", "Hereford"},
    {"HS", "Outer Hebrides"},
    {"HU", "Hull"},
    {"HX", "Halifax"},
    {"IG", "Ilford"},
    {"IP", "Ipswich"},
    {"IV", "Inverness"},
    {"KA", "Kilmarnock"},
    {"KT", "Kingston upon Thames"},
    {"KW", "Kirkwall"},
    {"KY", "Kirkcaldy"},
    {"L",  "Liverpool"},
    {"LA", "Lancaster"},
    {"LD", "Llandrindod Wells"},
    {"LE", "Leicester"},
    {"LL", "Llandudno"},
    {"LN", "Lincoln"},
    {"LS", "Leeds"},
    {"LU", "Luton"},
    {"M",  "Manchester"},
    {"ME", "Medway"},
    {"MK", "Milton Keynes"},
    {"ML", "Motherwell"},
    {"N",  "London (North)"},
    {"NE", "Newcastle"},
    {"NG", "Nottingham"},
    {"NN", "Northampton"},
    {"NP", "Newport"},
    {"NR", "Norwich"},
    {"NW", "London (North West)"},
    {"OL", "Oldham"},
    {"OX", "Oxford"},
    {"PA", "Paisley"},
    {"PE", "Peterborough"},
    {"PH", "Perth"},
    {"PL", "Plymouth"},
    {"PO", "Portsmouth"},
    {"PR", "Preston"},
    {"RG", "Reading"},
    {"RH", "Redhill"},
    {"RM", "Romford"},
    {"S",  "Sheffield"},
    {"SA", "Swansea"},
    {"SE", "London (South East)"},
    {"SG", "Stevenage"},
    {"SK", "Stockport"},
    {"SL", "Slough"},
    {"SM", "Sutton"},
    {"SN", "Swindon"},
    {"SO", "Southampton"},
    {"SP", "Salisbury"},
    {"SR", "Sunderland"},
    {"SS", "Southend-on-Sea"},
    {"ST", "Stoke-on-Trent"},
    {"SW", "London (South West)"},
    {"SY", "Shrewsbury"},
    {"TA", "Taunton"},
    {"TD", "Galashiels"},
    {"TF", "Telford"},
    {"TN", "Tonbridge"},
    {"TQ", "Torquay"},
    {"TR", "Truro"},
    {"TS", "Teesside"},
    {"TW", "Twickenham"},
    {"UB", "Uxbridge"},
    {"W",  "London (West)"},
    {"WA", "Warrington"},
    {"WC", "London (West Central)"},
    {"WD", "Watford"},
    {"WF", "Wakefield"},
    {"WN", "Wigan"},
    {"WR", "Worcester"},
    {"WS", "Walsall"},
    {"WV", "Wolverhampton"},
    {"YO", "York"},
    {"ZE", "Lerwick"}
};











// Main Functions
void make_job_data_excel(const std::string& file_path){
    std::string filename = "JobData" + FILE_EXTENSION;
    lxw_workbook  *workbook  = workbook_new(filename.c_str());
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Sheet1");



    worksheet_write_string(worksheet, 0, 0, "Job reference", NULL);
    worksheet_write_string(worksheet, 0, 1, "Job type", NULL);
    worksheet_write_string(worksheet, 0, 2, "Contact name", NULL);
    worksheet_write_string(worksheet, 0, 3, "Contact postcode", NULL);
    worksheet_write_string(worksheet, 0, 4, "Contact address", NULL);
    worksheet_write_string(worksheet, 0, 5, "Job contact person phone", NULL);
    worksheet_write_string(worksheet, 0, 6, "Planned start time", NULL);
    worksheet_write_string(worksheet, 0, 7, "Job contact person", NULL);
    worksheet_write_string(worksheet, 0, 8, "Contact group", NULL);
    worksheet_set_column(worksheet, 0, 10, 25, NULL);


    std::vector<std::string> lines = split(read_file(file_path), '\n');
    for (int i = 1; i < lines.size(); i++) {
        if (lines[i].length() < 2){
            continue;
        }

        std::string address = extract_single_quoted_string(lines.at(i));

        std::vector<std::string> line_data = split(lines[i], ',');

        // Job reference
        worksheet_write_string(worksheet, i, 0, line_data[0].c_str(), NULL);

        // Job type
        worksheet_write_string(worksheet, i, 1, "Anyjunk Upload", NULL);


        // Contact name
        worksheet_write_string(worksheet, i, 2, std::string("Anyjunk " + line_data[3]).c_str(), NULL);


        // Contact postcode
        worksheet_write_string(worksheet, i, 3, line_data[3].c_str(), NULL);

        // Contact address
        // Limit the address field as needed
        int address_max_length = 100;
        if (address.length() > address_max_length) {
            address = address.substr(0, address_max_length - 5);
            logger.log(std::string("Address with postcode : " + line_data[3] + "has been shortened"));
        }
        worksheet_write_string(worksheet, i, 4, std::string("\"" + address +"\"").c_str(), NULL);

        
        // Job contact person phone
        std::string phone_number = extract_phone_number(line_data[line_data.size() - 1]);
        worksheet_write_string(worksheet, i, 5, phone_number.c_str(), NULL);


        // Planned start time
        std::string planned_start_time = take_all_before(line_data[2], '-');
        worksheet_write_string(worksheet, i, 6, planned_start_time.substr(0, planned_start_time.length() - 1).c_str(), NULL);
        

        // Job contact person
        std::string person_name = take_all_before(line_data[line_data.size() - 1], '-');
        worksheet_write_string(worksheet, i, 7, person_name.substr(0, person_name.length() - 1).c_str(), NULL);


        // Contact group
        worksheet_write_string(worksheet, i, 8, "ANYJUNK 2023", NULL);
    }
    workbook_close(workbook);
    logger.log(std::string(filename + " file has been generated!"));
}



void make_contact_data_excel(const std::string& file_path){
    
    std::string filename = "ContactData" + FILE_EXTENSION;
    lxw_workbook  *workbook  = workbook_new(filename.c_str());
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, "Sheet1");


    worksheet_write_string(worksheet, 0, 0, "Contact name", NULL);
    worksheet_write_string(worksheet, 0, 1, "Reference", NULL);
    worksheet_write_string(worksheet, 0, 2, "Group", NULL);
    worksheet_write_string(worksheet, 0, 3, "Address", NULL);
    worksheet_write_string(worksheet, 0, 4, "Postcode", NULL);
    worksheet_write_string(worksheet, 0, 5, "City", NULL);
    worksheet_write_string(worksheet, 0, 6, "Country", NULL);
    worksheet_write_string(worksheet, 0, 7, "Primary person - Mobile phone", NULL);
    worksheet_set_column(worksheet, 0, 10, 25, NULL);

    std::vector<std::string> lines = split(read_file(file_path), '\n');
    for (int i = 1; i < lines.size(); i++) {
        if (lines[i].length() < 2){
            continue;
        }

        std::string address = extract_single_quoted_string(lines.at(i));

        std::vector<std::string> line_data = split(lines[i], ',');

        // Contact name
        worksheet_write_string(worksheet, i, 0, std::string("Anyjunk " + line_data[3]).c_str(), NULL);


        // Reference
        worksheet_write_string(worksheet, i, 1, std::to_string(get_random_number(0, 1000000000)).c_str(), NULL);


        // Group
        worksheet_write_string(worksheet, i, 2, "ANYJUNK 2023", NULL);


        // Address
        // Limit the address field as needed
        int address_max_length = 100;
        if (address.length() > address_max_length) {
            address = address.substr(0, address_max_length - 5);
            logger.log(std::string("Address with postcode : " + line_data[3] + "has been shortened"));
        }
        worksheet_write_string(worksheet, i, 3, std::string("\"" + address +"\"").c_str(), NULL);


        // Postcode
        worksheet_write_string(worksheet, i, 4, line_data[3].c_str(), NULL);


        // City
        std::string postcode = remove_numbers_from_string(trim(take_all_before(line_data[3], ' ')));
        auto iterator = LOCATION_MAP.find(postcode);
        if (iterator == LOCATION_MAP.end()) {
            worksheet_write_string(worksheet, i, 5, "N/A", NULL);
            logger.log(
                std::string(
                    "Cant find valid city for postcode : " + postcode + 
                    " Inserting N/A at Row : " + std::to_string(i) + " Column 5"
                )
            );
        }
        else{
            worksheet_write_string(worksheet, i, 5, iterator->second.c_str(), NULL);
        }

        // Country
        worksheet_write_string(worksheet, i, 6, "United Kingdom", NULL);


        // Primary person - Mobile phone
        std::string phone_number = extract_phone_number(line_data[line_data.size() - 1]);
        worksheet_write_string(worksheet, i, 7, phone_number.c_str(), NULL);

    }
    workbook_close(workbook);
    logger.log(std::string(filename + " file has been generated!"));
}





int main() {
    std::string file_path;

    for (const auto& entry : std::filesystem::directory_iterator(".")) {
        if (entry.path().extension() == ".csv") {
            file_path = entry.path().filename();
        }
    }

    if (file_path.empty()){
        logger.log("No csv file can be found aborting");
        logger.write_to_file();
        exit(1);
    }

    make_job_data_excel(file_path);
    make_contact_data_excel(file_path);

    logger.write_to_file();
}