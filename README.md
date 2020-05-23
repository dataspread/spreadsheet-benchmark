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
Our goal is to understand the impact 
of the type of operation,
the size of data being operated on, 
and the spreadsheet system used, on the
latency. 
Moreover, we want to quantify 
when each spreadsheet system
fails to be interactive 
for a given operation, 
violating the $500$ms mark
widely regarded as the bound for interactivity~\cite{liu2014effects}.

## Optimization Opportunities Testing (OOT) 
Spreadsheet systems
have continued to increase their size 
limits over the past few decades~\cite{excel-limit, gs-limit}. 
On the other hand, research on data management
has, over the past four decades, 
identified a wealth of techniques for 
optimizing the processing of large datasets. 
We wanted to understand
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
Our goal is to identify
new opportunities for
improving the design
of spreadsheet systems
to support computation on
large datasets.

# Implementation
The \excel-based experiments 
were conducted with Microsoft Excel 2016 running on  Windows, 
while the \calc-based experiments 
were conducted on LibreOffice Calc 6.0.3.2 running on Ubuntu. 
The \gs-based experiments were run on a 
university allocated G Suite account. 
For all three spreadsheet systems, 
we implemented the experiments in their corresponding scripting language, 
\ie Visual basic (\emph{VBA}) for \excel, 
Calc basic for \calc, and Google apps script (\emph{GAS}) for \gs. 
All the experiments were single threaded. 
Note that \excel 2016 can be configured 
to support multi-threaded recalculation of formulae~\cite{excel-mtc}. 
However, the default setting is to evaluate a formula 
on the main thread of \excel. 
\techreportAP{To ensure that all experiments operated 
on the entire dataset, we selected the desired data within the spreadsheet via a macro command.} 

# Benchmarking
For each experiment in \excel, we first 
created an Excel Macro-Enabled Workbook (\emph{xlsm})~\cite{xlsm} 
which can execute embedded macros programmed in VBA. 
Unlike \excel, LibreOffice \calc macros, programmed in \calc Basic, 
can be enabled and executed from the default workbook---OpenSpreadsheet Document (\emph{ods})~\cite{ods}. 
We created the Google App Scripts in 
G Suite Developer Hub~\cite{GAS}. 
Given an experiment, all three scripting languages can 
invoke a formula, \eg \code{COUNTIF}, or operation, \eg sort, 
for their respective systems via an API call. 
We used default library functions of the corresponding 
scripting languages to measure the execution time of each experimental trial. 
For each experiment, we passed the file path of the 
relevant datasets as an argument for the scripts (macros) 
of the desktop-based systems, and a URL for GAS in \gs. 
All the datasets used in the \excel and \calc-based experiments 
were in \emph{xlsx} and \emph{ods} format, 
respectively. 
The datasets used in the \gs experiments were uploaded as \emph{xlsx} 
files and then manually converted to Google Sheets from the Google Drive menu.

For each experiment, we ran ten trials and measured the running time. 
We report the average run time of eight trials while removing the maximum and minimum reported time. 
\techreportAP{Note that the \gs experimental settings were 
limited by the daily quotas 
and hard limits imposed by 
Google Apps Script services 
on some features, 
like API calls and the number of spreadsheets 
created and accessed. 
As a result,}\onlypapertext{Note that} for experiments with \gs, 
we restricted the maximum size of the data to $90$k 
rows to fit in the experiment trials for different test cases 
within the allocated daily quotas

