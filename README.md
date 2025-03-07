# 🎓 University Task Automation Scripts  

This repository contains two Python scripts designed to automate data processing tasks at **"1 Decembrie 1918" University of Alba Iulia**. These scripts help with organizing and managing student and high school records efficiently.  

## 📜 Overview  

### 1️⃣ High School Data Processing Script  
This script processes Excel files with student records, categorizing them by county and high school. It:  
- Extracts county names from bold-marked cells.  
- Normalizes and groups similar high school names using fuzzy matching.  
- Generates an Excel report with structured data and total counts per county.  

### 2️⃣ Excel Merging Script  
This script merges multiple student data Excel files while ensuring consistency. It:  
- Reads and validates input files.  
- Checks for missing or incorrect columns.  
- Standardizes data formatting.  
- Adds an index column for better organization.  

## ⚙️ Installation  

Ensure you have **Python 3.x** installed, along with the required dependencies. You can install them using:  

```bash
pip install pandas openpyxl fuzzywuzzy
