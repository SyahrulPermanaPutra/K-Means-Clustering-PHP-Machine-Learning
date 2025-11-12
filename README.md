# K-Means-Clustering-PHP-Machine-Learning

PHP Machine Learning Clustering with K-Means
This project implements K-Means Clustering using PHP-ML for wine quality dataset analysis with Excel export functionality.

Quick Setup (Windows)
1. Prerequisites Installation
Install PHP:
Option 1: XAMPP (Recommended for beginners)

Download from: https://www.apachefriends.org/download.html

Install with default settings

PHP will be available at: C:\xampp\php\

Option 2: Manual PHP Installation

Download PHP: https://windows.php.net/download/

Choose PHP 8.1+ Thread Safe ZIP package

Extract to: C:\php\

Setup PHP Environment PATH:
Open System Properties → Environment Variables

Under System Variables, edit Path

Add PHP path:

For XAMPP: C:\xampp\php

For manual: C:\php

Click OK to save

Verify PHP Installation:
cmd
php --version
Should display PHP version, not error

Install Composer:
Download Composer: https://getcomposer.org/download/

Run installer, check "Add to PATH" option

Select the installed PHP (auto-detected if PATH is correct)

Verify Composer:
cmd
composer --version
2. Project Setup
bash
# Create project directory
mkdir C:\phpml-kmeans
cd C:\phpml-kmeans
3. Install Dependencies
bash
# Install PHP-ML for machine learning
composer require php-ai/php-ml

# Install PhpSpreadsheet for Excel export
composer require phpoffice/phpspreadsheet
4. Prepare Dataset & Files
Download winequality-red.csv

Place it in project folder: C:\phpml-kmeans\

Create test-excel.php file and copy the provided clustering code

5. Run the Program
bash
cd C:\phpml-kmeans
php test-excel.php
Project Structure
text
phpml-kmeans/
├── vendor/                 # Dependencies (auto-generated)
├── winequality-red.csv     # Wine quality dataset
├── test-excel.php         # Main clustering program
├── composer.json          # Configuration
├── composer.lock         # Lock file
└── clustering_results_*.xlsx  # Output files (auto-generated)
