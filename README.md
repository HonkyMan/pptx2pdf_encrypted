# PDF Conversion and Protection Tool

This Python application facilitates the conversion of PPTX files to PDF format and applies protection to the resulting PDF files. It uses a YAML configuration file for setting up the conversion and protection parameters.

## Features

- **Conversion of PPTX to PDF**: Batch convert PPTX files within a specified directory to PDF format.
- **PDF Protection**: Apply encryption to the converted PDF files, setting an owner password and restricting permissions such as printing and editing.

## Requirements

To run this tool, you need Python 3.6 or later and the following packages:
- `pikepdf`
- `PyYAML`

Install them using pip:

```bash
pip install pikepdf PyYAML
```

# Configuration
Before running the application, you need to create a config.yaml file in the application directory with the following structure:

```
source_dir: "path/to/source/directory"
dist_dir: "path/to/destination/directory"
owner_password: "YourOwnerPassword"
```

- source_dir: Directory containing the PPTX files to be converted.
- dist_dir: Directory where the converted PDF files will be saved.
- owner_password: Password for PDF encryption and restriction settings.

# Usage

1. Ensure you have a valid config.yaml file in your application directory.
2. Run the application with Python:

```
python app.py
```

# How It Works

- The application reads the configuration from config.yaml.
- It then converts all the PPTX files found in source_dir to PDF format and saves them to dist_dir in the same folder structure as was in source_dir.
- After conversion, it applies protection to the PDF files using the specified owner_password.

# Contributing

Feel free to fork this repository and submit pull requests to contribute to this project. For major changes, please open an issue first to discuss what you would like to change.

# License

MIT