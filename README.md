# VBA-Dataset-Seperator
 A mini project using VBA to automate dataset of car sales into seperate files based on a category(regional sales & company wise car sales).

##Overview

### Table of Contents
* Overview
* Features
* Getting Started
* Prerequisites
* Installation
* Usage
* Configuration
* License
* Acknowledgments
* Features

## Overview

* Category-Based Separation: The script separates the dataset into multiple files based on user-defined categories.(Region,Company)
*  Automation: The process is automated through a VBA script, saving time and effort compared to manual separation.
*  Customizable: Users can easily customize the categories and file naming conventions according to their specific requirements.
* Analysis: Contains models to perform data cleaning , analysis on KPI's and basic visualization.

## Features

* MainProgram: performs the seperation of files according to company of car.
* Region_dealership: performs the seperation of files according to region of sales.

###  Bonus 
In addition to the above macro for seperating the data base according to different regions, macros for analysis on KPI's are included that perform the following funtions.

* Trim White Space
* Replace missing values
* Standardize text case
* Remove Duplicates

* Average customer income 
* Average sales price
* Average unit sales per day
* Average daily sales 
* Gender distribution of customers 
* Median customer income
* Total sales revenue

* Visualization of average sales


## Getting Started
This section provides instructions for getting started with the VBA Dataset Separator project.

Prerequisites
Before using this project, ensure that you have the following prerequisites installed on your system:

Microsoft Excel (version X or later)
Basic knowledge of VBA programming
### Installation
To install the VBA Dataset Separator project, follow these steps:

Download the project files:

* Clone this repository to your local machine using Git,

* Open the Excel file:

* Navigate to the directory where you extracted the project files.
* Open the DatasetSeparator.xlsm file in Microsoft Excel.

## Usage
Follow these steps to use the VBA Dataset Separator project:

### Prepare your dataset:

Ensure that your dataset is correctly formatted and located in the designated worksheet within the Excel file.

### Configure the script:

* Open the VBA editor in Excel by pressing ALT + F11.
* Navigate to the DatasetSeparatorModule module.
* Modify the category definitions and file naming convention according to your dataset and preferences (see the Configuration section for details).

* Run the script:

* Close the VBA editor.

* In the Excel file, press ALT + F8 to open the "Macro" dialog.
* Select the SeparateDataset macro from the list and click "Run".
* Follow the on-screen prompts to complete the separation process.

### Review the results:

Once the script has finished running, check the specified directory for the separated files, each containing data corresponding to a specific category.

### Configuration
* You can modify the code according to your requirements.Just make sure to change the file name and the name of the column to column you want to seperate the data based on.
* Specify the column range upto which you want your new datasets to be saved.


