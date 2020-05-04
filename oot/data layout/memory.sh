#!/bin/bash
array=( 10000 30000 50000 )
for j in "${array[@]}" # for each file size
do
    for i in {1..10} # for 10 trials
    do  
        file = "/path/to/file" #file directory
        extension = ".ods"
        filepath = $file$j$extension #filename with path
        libreoffice --calc "${filepath}" #open file
        pid=$(ps aux | grep soffice.bin | head -n 1 | awk '{print $2}') # get soffice process id
        pmap -x $pid | tail -n 1 | awk '{print $5}' #get memomry consumption
        kill -9 $pid # close file
    done
done