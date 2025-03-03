# Data Formatter CLI

<p align="center">
    <img title="Data Formatter CLI" height="100" src="https://raw.githubusercontent.com/laravel-zero/docs/master/images/logo/laravel-zero-readme.png" alt="Laravel Zero Logo" />
</p>

A powerful command-line tool for converting and manipulating Excel files to CSV format with extensive customization options. Built with Laravel Zero.

## Features

- **Excel to CSV Conversion**: Easily convert Excel (.xlsx) files to CSV format
- **Interactive Mode**: Guides you through the conversion process with interactive prompts
- **Value Mapping**: Map column values to new values during conversion
- **Column Management**: Selectively remove columns you don't need
- **Format Customization**: Configure delimiters, encoding, and more
- **Flexible Output**: Control headers, skip rows, and other output options

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/data-formatter-cli.git

# Navigate to the project directory
cd data-formatter-cli

# Install dependencies
composer install

# Make the command executable
chmod +x data-formatter-cli
```

## Usage

### Basic Usage

```bash
./data-formatter-cli exceltocsv
```

The command will guide you through the process interactively, prompting for:
- Input Excel file
- Output CSV location
- Delimiter character
- Column modifications
- Value mapping options

### Available Options

```bash
./data-formatter-cli exceltocsv [options]
```

Options:
- `--sheet=SHEET_NAME`: Specific sheet name to convert (default: first sheet)
- `--encoding=ENCODING`: Output file encoding (default: UTF-8)
- `--skip-rows=N`: Number of rows to skip from the beginning
- `--no-headers`: Don't include headers in output

### Example Workflow

1. Run the command
2. Select your Excel file
3. Choose your output CSV name
4. Configure your CSV delimiter
5. Optionally remove columns you don't need
6. Optionally map values for specific columns

The tool will handle the conversion and save the output file with your specified settings.

## Building a Standalone Application

You can build a standalone PHAR file using:

```bash
php data-formatter-cli app:build
```

## Credits

Data Formatter CLI is built on top of [Laravel Zero](https://laravel-zero.com), an elegant starting point for your Laravel console application by [Nuno Maduro](https://github.com/nunomaduro).

## License

This project is open-sourced software licensed under the MIT license.
