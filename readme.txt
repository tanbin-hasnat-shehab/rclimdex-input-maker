inputs are
		1.start year of the base period
		2.end year of the base period
		3.name of the station(say Dhaka,Bogra,Barisal,Cox's Bazar etc. just as in the excel files****remember its case and space sensitive)
		


if any year data missing there will pop up a warning
if everything ok...then a success pop up will appear

in case of erorr missing data pop ups :
1.read the messege which file has missed data max.xlsx or min.xlsx or rain.xlsx
2.open that xlsx file
3.the cell A1 has got a hyperlink click it which will take you to the start year of the base period
4.find out the missing year,remember it then close the file save or don't save take any option

example:

base period 1961-2018 say
but database at somewhere like,
	....
	year	month	1	2	3	4  ........
	1973	12	25	20	22	20......
	1975	1	22	26	32	15.....
	....
5.then start the program again input base period 1961 to 1973
6.after successfully created input for 1961 to 1973 just restart the program having base period 1975 to 2018

thats it


  