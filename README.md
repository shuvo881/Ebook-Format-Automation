# Ebook-Format-Automation
The project focuses on automating the formatting of Word documents to optimize them for digital use while adhering to specified formatting rules. The objective is to fix layout issues, standardize fonts, and ensure clean, professional output for ebooks.


## Installation

1. Clone the repository:
```bash
git https://github.com/shuvo881/Ebook-Format-Automation.git
cd Ebook-Format-Automation
```

3. Set up environment variables:

```bash
# Create a vertual enviroment
## for Mac and Linux
python -m venv venv

# Active the verual enviroment
## For Mac and Linux
source venv/bin/active
```

```bash
# Install Requirements
pip install -r requirements.txt
```

### Basic Usage

```bash
cd code
# run the main file:
python main.py
```

# File Structure 
```bash
Ebook-Format-Automation/
│
├── README.md
├── requirements.txt
│
├── data/
│   ├── Ebook/
│   │   ├── 278160.docx
│   │   ├── 90191.docx
│   │   ├── 429332.docx
│   │   ├── 416455.docx
│   │
│   ├── Processed_Ebooks/
│   │   ├── 278160.docx
│   │   ├── 90191.docx
│   │   ├── 429332.docx
│   │   ├── 416455.docx
│
├── code/
│   └── main.py

```
Main Project Files:

* README.md: Likely contains project documentation.
* requirements.txt: Contains Python dependencies for the project.
* main.py: Main script for the project, found in the code/ directory.

Data Files:

* Ebook: Contains original .docx files and an Excel file for font corrections.
* Processed_Ebooks: Contains .docx files that appear to have been processed.
