# top_to_excel
This small script converts the output from TOP linux command to EXCEL file. It doesn't include the usage for each process since I have no idea what's the best way to demonstrate them yet. But it gives you a brief understanding about the CPU/memory usages.


Under DUT to get the log of top.
while [ true ]; do
	sleep 5
	echo "." >> ./test.txt
	top -b -c -n 1 >> ./test.txt
	ectool temps 0 >> ./test.txt
	ectool temps 1 >> ./test.txt
	ectool temps 2 >> ./test.txt
	ectool temps 3 >> ./test.txt
	ectool pwmgetfanrpm all >> ./test.txt
done

Run this script:
$python parse_top.py -i ./test.txt -o ./


