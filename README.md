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

## Experiments: Datasets and APIs
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