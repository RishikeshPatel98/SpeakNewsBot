import requests
import json
import win32com.client as wincom
query = input("What type of news are you interested in? ")
url = f"https://newsapi.org/v2/everything?q={query}&from=2023-03-20&sortBy=publishedAt&apiKey=585fb71d5daa4c7db1775479b162e0f5"
r = requests.get(url)
news = json.loads(r.text)
for article in news["articles"]:
    print(article["title"])
    print(article["description"])
    break
title = article["title"]
description = article["description"]
speak = wincom.Dispatch("SAPI.SpVoice")
text = f"The news related your search result are title{title}description{description}"
speak.Speak(text)
