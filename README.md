# spreadsheet-benchmark
Spreadsheet systems are used for storing and analyzing data
across domains by programmers and non-programmers alike.
While spreadsheet systems have continued to support
increasingly large datasets,
they are prone to hanging and
freezing while performing computations even on much smaller ones.
We have developed this benchmark
to evaluate and compare the performance of a spreadsheet system.
The benchmark a) measures the scalability of a spreadsheet systems
and b) investigates how a spreadsheet system stores data and whether it adopts optimizations
to speed up computation.
In our current release, we evaluate three popular systems, Microsoft Excel, LibreOffice Calc, and Google Sheets,
on a range of canonical spreadsheet computation operations.

# Design 
We construct two different kinds of benchmarks
to evaluate
these spreadsheet systems: _basic complexity testing (BCT)_, 
and _optimization opportunities testing (OOT)_. 


## Basic Complexity Testing (BCT) 
The BCT benchmark aims to assess the performance of
basic operations on spreadsheets. 
We construct a taxonomy of operations---encapsulating
opening, structuring, editing, and analyzing data---based on their
expected time complexity, and 
evaluate the relative performance of the spreadsheet
systems on a range of data sizes.

## Optimization Opportunities Testing (OOT) 
The OOT benchmark invesigates
whether
spreadsheet systems
take advantage of
techniques
such as 
indexes, 
incremental updates, 
workload-aware data layout, and 
sharing of computation.
The OOT benchmark 
constructs specific scenarios
to explore 
whether such optimizations are 
deployed by existing spreadsheet systems
while performing spreadsheet formula computation.

# Implementation Details
For all three spreadsheet systems, 
we implemented the experiments in their corresponding scripting language:
Visual basic (_VBA_) for Excel, 
Calc basic for Calc, and Google apps script (_GAS_) for Google Sheets. 
The file extension for VBA, Calc basic, and GAS scripts are _.cls_, _.bas_, and _.gs_, respectively.
All the experiments are single threaded. 

## Expeiment files and dataset
For each experiment in Excel, we first 
create an Excel Macro-Enabled Workbook (_.xlsm_) 
which can execute embedded macros programmed in VBA. 
Unlike Excel, LibreOffice Calc macros, programmed in Calc Basic, 
can be enabled and executed from the default workbook---OpenSpreadsheet Document (_.ods_). 
We create the Google App Scripts in 
[G Suite Developer Hub](https://developers.google.com/gsuite). 
Given an experiment, all three scripting languages can 
invoke a formula, e.g., _COUNTIF_, or operation, e.g., _SORT_, 
for their respective systems via an API call. 
We used default library functions of the corresponding 
scripting languages to measure the execution time of each experimental trial. 
For each experiment, we passed the file path of the 
relevant datasets as an argument for the scripts (macros) 
of the desktop-based systems, and a URL for GAS in Google Sheets. 
All the datasets used in the Excel and Calc-based experiments 
are in _xlsx_ and _ods_ format, 
respectively. 
The datasets used in the Google Sheets experiments are uploaded as _xlsx_ 
files and then manually converted to Google Sheets from the Google Drive menu.

## Execution time measurement
For each experiment, we run ten trials and report the average run time of
eight trials while removing the maximum and minimum reported time. 
Note that the Google Sheets experimental settings are 
limited by the daily quotas 
and hard limits imposed by 
Google Apps Script services 
on some features, 
like API calls and the number of spreadsheets 
created and accessed. 
As a result, for experiments with Google Sheets, 
we restrict the number of data points, i.e., row sizes,
to fit in the experiment trials for different test cases 
within the allocated daily quotas

# Benchmarking

To get started with the _spreadsheet-benchmark_, first clone or down the repository. 
To clone the repository use: `git clone https://github.com/dataspread/spreadsheet-benchmark.git`

The `/bct` and `/oot` directories contain the BCT and OOT benchmark experiments, respecticely.
Each benchmark is further categorized based on operations (BCT) or optimizations (OOT) tested. 
Following is the benchmark organization:

```bash
├── bct
│   ├── load
│   ├── query
│   └── update
├── oot
│   ├── data layout
│   ├── incremental update
│   ├── indexing
│   └── shared computation
├── .gitignore
├── randomized_setup.md
├── README.md
└── script.py
```

## Running an experiment 

First, create a `/data` directory locally (for Excel and Calc) or
in Google Drive (for Google Sheets) and save your experimental datasets there. 
The process of running an experiment varies with each spreadsheet system which we explain next.

### Excel

Create a `.xlsm` file and open the Excel visual basic editor:

```
Click the "Visual Basic" button on the "Developer" tab.

If the Developer tab is not present, go to File -> Options -> customize ribbon and tick Developer.

You can also open VBA in Excel using Alt + F11 keyboard shortcut.
```
Load a `.cls` file from the `Import File ...` option in the `File` menu.

Click `run` to launch an experiment.

### Calc

Create a `.ods` file and create a `module` for the macro to be placed into:

```
From menu open Tools->Macros->Organize Macros->Libre Office Basic.

On the following dialog select the library (such as Standard) where to create the module. 

Then click new & give it a name (based on your experiment. You can reuse the respective .bas file name)
```
Double click on the `module` to open the code editor. Then from menu select `File->Import Basic` and select your `.bas` file.

Click `run` to launch an experiment.

### Google Sheets

Create a new project in the Google Apps Script home. Copy the contents from a `.gs` file into the GAS script. 

To launch the experiment, go to `Run->Run function` in the menu and select the main function you want to run. You are unable to pass any parameters into the function through this method.

There are three methods we explored in running Google Sheets experiments:
1. All trials run during a script execution
2. One trial per script execution via Trigger
3. One trial per script execution via API

#### Method 1
No changes need to be made for method 1.

#### Method 2
To randomize the order of dataset sizes that would run the experiment on and avoid hitting the timeout limit, we changed the script to select a random dataset size and run only one trial of the experiment. To avoid having to manually execute the script for each trial, go to `Edit->Current project's triggers` then click `+ Add Trigger` and configure the settings. 

A configuration of `hello_world`, `Head`, `Time-driven`, `Minutes Timer`, and `Every minute` will run the current version of the `hello_world` function every minute. It is recommended that you configure the time interval to be greater than the runtime of your script to avoid overlap of script execution, which could lead to concurrency issues if you're using the same data.

#### Method 3 (randomized_setup.md)
Method 3 follows the same incentive with the addition of predetermining the order of the trials and increasing efficiency. In order to run the randomized trial GS scripts, refer to `randomized_setup.md`.