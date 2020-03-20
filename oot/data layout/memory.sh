pid=$(ps aux | grep soffice.bin | head -n 1 | awk '{print $2}')
pmap -x $pid | tail -n 1 | awk '{print $5}'

