# Certificake 🍰

A CLI tool to create event certicates automatically using PPTX templates and lists of participants CSV/XLSX file.

## Requirements

- [Libre Office](https://pt-br.libreoffice.org/) (for PDF generation)

## Install

```
$ pip install git+https://github.com/omixlab/certificake
```

## Usage

```
$ certificake \
  --template examples/template.pptx \
  --data examples/data.csv \
  --output "examples/{{ name }}" \
  --pdf 
```

## Instructions

- PPTX template files (`--template`) might be tested previously in Libre Office, as some elements 
might be rendered different on other programs (eg: Microsoft Powerpoint). 

- Avoid using spaces, special symbols or fields starting with numbers on the CSV/XLSX file containg the
data to be included on the certificates (`--data`). In case of field with more than one word, use underscores (eg: "full_address" instead of "full address").

