# Expocentr Parser

This script parses the exhibition links from https://icatalog.expocentr.ru/ru and saves them to a JSON file.

## Requirements

- Python 3.6+
- Required packages listed in `requirements.txt`

## Installation

1. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

Simply run the script:
```bash
python parser.py
```

The script will:
1. Fetch the webpage
2. Extract all links with class "list-group-item list-group-item-action"
3. Save the results to `expo_links.json`

## Output Format

The script creates a JSON file with the following structure:
```json
[
  {
    "text": "Exhibition Name",
    "url": "https://icatalog.expocentr.ru/..."
  },
  ...
]
``` 