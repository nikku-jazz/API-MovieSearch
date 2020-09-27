import wget
import msvcrt
import requests
import xlsxwriter
from requests.exceptions import ConnectionError


main_url='https://api.themoviedb.org/3/search/movie?'
api='api_key=1dcf69b9b95240032c80e5d374ca2bee'
search_query='&query='
no_of_pages='&page='

movie_name=None
movie_data=None
total_pages=None
total_results=None
keystroke=None
lines_XLSX=2

Movie_Name=[]
Movie_Overview=[]

def MovieSearch():
    global movie_name, movie_data, total_pages, total_results
    try:
        movie_name=movie_name.replace(' ','+')
        requested_data = requests.get(url =main_url+api+search_query+movie_name)
        movie_data = requested_data.json()
        total_pages=int(movie_data['total_pages'])
        total_results=int(movie_data['total_results'])
        DataManipualation(currPage=1, currResults=20)
    except ConnectionError:
        print ('Connection Error...Press Space to Retry or Enter to Exit')
        keypress()
        while keystroke:
            MovieSearch()
        exit()

def DataManipualation(currPage, currResults):
    global total_pages, total_results, keystroke
    if total_pages==0:
        print('No Results Found')
    elif total_pages==1:
        currResults=total_results
        ShowData(currResults)
    elif (total_pages>1):
        for currPage in range(2,total_pages):
            ShowData(currResults)
            print('\nPress Space to Load Next Page or Enter to Exit')
            keypress()
            if keystroke:
                RequestData(currPage)
            else:
                SearchAgain()
        currPage+=1
        currResults=total_results-20*(total_pages-1)
        RequestData(currPage)
        ShowData(currResults)
    SearchAgain()

def RequestData(currPage):
    global movie_data, total_pages, total_results
    requested_data = requests.get(url =main_url+api+search_query+movie_name+no_of_pages+str(currPage))
    movie_data = requested_data.json()

def ShowData(currResults):
    global movie_data
    for i in range(currResults):
        movieName=movie_data['results'][i]['original_title']
        overview=movie_data['results'][i]['overview']
        image=movie_data['results'][i]['poster_path']

        print('\n\n Name Of Movie ',i+1,': '+movieName)
        Movie_Name.append(movieName)

        if(overview!=''):
            print(' Overview Of Movie ',i+1,': '+overview,'\n')
            Movie_Overview.append(overview)
        else:
            Movie_Overview.append('NA')
            print(' Overview Of Movie: NA')
        
        if(image!=None):
            print('\nDownloading Poster...')
            image_url='https://image.tmdb.org/t/p/w500/'+image
            wget.download(image_url)
        else:
            print('Image: NA')
    SaveData(currResults)

def SaveData(currResults):
    global lines_XLSX
    workbook = xlsxwriter.Workbook('movie.xlsx') 
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Movie Name') 
    worksheet.write('B1', 'Movie Description')
    for i in range(2,lines_XLSX+currResults):
        wb='A'+str(i)
        wbb='B'+str(i)
        worksheet.write(wb, Movie_Name[i-2]) 
        worksheet.write(wbb, Movie_Overview[i-2]) 
    lines_XLSX+=currResults
    workbook.close()

def SearchAgain():
    global movie_name
    searchAgain=input('\nSearch Again or Press Enter to Exit: ')
    if searchAgain=='':
        exit()
    movie_name=searchAgain
    MovieSearch()

def keypress():
    global keystroke
    if msvcrt.getch()==b' ':
        keystroke=True
    else:
        keystroke=False

movie_name=input('Enter Movie Name: ')
if movie_name=='':
    exit()
MovieSearch()

# def comments():
    #/////////////////////////////////////////////////////////////////////////////////////////////
    # image_url='https://image.tmdb.org/t/p/w500/'+data['results'][i]['poster_path']
    # wget.download(image_url)
    #/////////////////////////////////////////////////////////////////////////////////////////////
    
    #Example Format
    # {
    #   'poster_path': '/IfB9hy4JH1eH6HEfIgIGORXi5h.jpg',
    #   'adult': false,
    #   'overview': 'Jack Reacher must uncover the truth behind a major government conspiracy in order to clear his name. On the run as a fugitive from the law, Reacher uncovers a potential secret from his past that could change his life forever.',
    #   'release_date': '2016-10-19',
    #   'genre_ids': [
    #     53,
    #     28,
    #     80,
    #     18,
    #     9648
    #   ],
    #   'id': 343611,
    #   'original_title': 'Jack Reacher: Never Go Back',
    #   'original_language': 'en',
    #   'title': 'Jack Reacher: Never Go Back',
    #   'backdrop_path': '/4ynQYtSEuU5hyipcGkfD6ncwtwz.jpg',
    #   'popularity': 26.818468,
    #   'vote_count': 201,
    #   'video': false,
    #   'vote_average': 4.19
    # }
    #//////////////////////////////////////////////////////////////////////////////////////////////

    # movie_name=[]
    # movie_overview=[]
    # movie_release=[]
    # sending get request and saving the response as response object 
    # r = requests.get(url = link) 
    # data = r.json() 
    # results=data['total_results']
    # for i in range(0,5):
    #     movie_name.append(data['results'][i]['original_title'])
    #     movie_release.append(data['results'][i]['release_date'])
    #     movie_overview.append(data['results'][i]['overview'])

    # workbook = xlsxwriter.Workbook('movie.xlsx') 
    # worksheet = workbook.add_worksheet() 
    # worksheet.write('A1', 'Movie Name') 
    # worksheet.write('B1', 'Movie Release Date') 
    # worksheet.write('C1', 'Movie Description') 
    # for i in range(2,5):
    #     wb='A'+str(i)
    #     wbb='B'+str(i)
    #     wbbb='C'+str(i)
    #     worksheet.write(wb, movie_name[i-2]) 
    #     worksheet.write(wbb, movie_release[i-2]) 
    #     worksheet.write(wbbb, movie_overview[i-2]) 
    # Finally, close the Excel file 
    # via the close() method. 
    # workbook.close()
