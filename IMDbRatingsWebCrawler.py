import imdb
import openpyxl
import datetime

try:
	# Create the object that will be used to access the IMDb's database.
	ia = imdb.IMDb() # by default access the web.
	wb = openpyxl.load_workbook('MovieDataClean.xlsx')
	sheet = wb.get_sheet_by_name('Sheet1')
	
	for i in range(2,1500):
	movie_name = sheet['A' + str(i)].value
	movie_date = sheet['B' + str(i)].value.year
	
	# Search for a movie (get a list of Movie objects).
	s_result = ia.search_movie(movie_name, movie_date)
	
	the_unt = s_result[0]
	ia.update(the_unt)
	
	rating = the_unt['rating']
	
	sheet['E' + str(i)] = rating
	print movie_name, rating
	wb.save('MovieDataClean.xlsx')
	
except KeyError:
	pass
