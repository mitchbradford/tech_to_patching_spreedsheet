# tech_to_patching_spreadsheet

Parses files containing one or more show tech of Cisco devices
and extracts port information to an Excel file to assist in creating
port patching spreedsheets for switch replacements.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

```
Python 3
xlwt
```

### Installing

Install the xlwt library

## Running the script
```
git clone https://github.com/mitchbradford/tech_to_patching_spreedsheet.git
cd tech_to_patching_spreadsheet
```
Copy files containing a 'show tech' from cisco switches into the folder containing this script
```
python tech_to_patching_spreadsheet <output_filename>.xls <input_file_1> <input_file_2> ... etc
```

## Version 1.0
Initial Release and code tidy up

## Version 0.1
Initial Fork

## Authors

* **Mitch Bradford** - [Github](https://github.com/mitchbradford)

## License

This project is licensed under the GNU General Public License v3.0.

## Acknowledgments
* Originally forked from https://github.com/angonz/tech2xl by [Andres Gonzelez](https://github.com/angonz)

