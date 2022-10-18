import requests,openpyxl
from bs4 import BeautifulSoup

wb=openpyxl.Workbook()
ws=wb.active
ws.title='Most Popular TV Shows'
ws.append(['Rank','TV Show','IMDB Rating','Date of Release'])




try:
    source=requests.get("https://www.imdb.com/chart/tvmeter/?ref_=nv_tvv_mptv")
    source.raise_for_status()

    soup=BeautifulSoup(source.text,'html.parser')
    shows=soup.find('tbody',class_='lister-list').find_all('tr')

    for show in shows:
        name=show.find('td',class_='titleColumn').a.text
        rank=show.find('td',class_='posterColumn').find_all('span')[0]['data-value']
        rating = show.find('td', class_='posterColumn').find_all('span')[1]['data-value']
        year=show.find('td',class_='titleColumn').span.text.strip("()")


                                                                                #if rank.span['name']=='rk':
                                                                                #print(rank.span['data-value'])
                                                                                #if rank.span['name']=='ir':
                                                                                #   print(rank.span['data-value'])
                                                                                        #print(show)
        ws.append([rank,name,rating,year])

except Exception as e:
    print(e)
wb.save('Most Popular TV Shows.xlsx')
