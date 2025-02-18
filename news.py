import requests	

def NewsFromBBC():
	query_params = {
	"source": "bbc-news",
	"sortBy": "top",
	"apiKey": "561497bcd1414c08a580b39e4c6670d4"
	}
	main_url = " https://newsapi.org/v1/articles"

	res = requests.get(main_url, params=query_params)
	open_bbc_page = res.json()
	article = open_bbc_page["articles"]

	# empty list which will
	# contain all trending news
	results = []
	
	for ar in article:
		results.append(ar["title"])
		
	for i in range(len(results)):
		
		print(i + 1, results[i])

	from win32com.client import Dispatch
	speak = Dispatch("SAPI.Spvoice")
	speak.Speak(results)				

if __name__ == '__main__':
	
	# function call
	NewsFromBBC()
by Mr.Engineer0510
