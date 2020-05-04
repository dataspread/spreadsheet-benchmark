#!/bin/bash
#open each file in libre office manually
#TODO: automate the process
pid=$(ps aux | grep soffice.bin | head -n 1 | awk '{print $2}') # get soffice process id
pmap -x $pid | tail -n 1 | awk '{print $5}' #get memomry consumption


