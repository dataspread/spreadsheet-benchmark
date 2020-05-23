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


## Basic Complexity Testing (BCT). 
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

# Optimization Opportunities Testing (OOT). 
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
VBA BASIC GAS

# Benchmarking
load files

