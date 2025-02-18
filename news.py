import requests
import json
from win32com.client import Dispatch

def fetch_news_from_bbc(api_key):
    query_params = {
        "sources": "bbc-news",
        "sortBy": "top",
        "apiKey": api_key
    }
    main_url = "https://newsapi.org/v2/top-headlines"
    
    try:
        res = requests.get(main_url, params=query_params)
        res.raise_for_status()  # Check for request errors
    except requests.exceptions.RequestException as e:
        print(f"Error fetching news: {e}")
        return None

    return res.json()

def speak_text(text):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

def news_from_bbc(api_key):
    news_data = fetch_news_from_bbc(api_key)
    if not news_data:
        return

    articles = news_data.get("articles", [])
    if not articles:
        print("No news articles found.")
        return

    results = []
    for ar in articles:
        title = ar.get("title", "No Title")
        description = ar.get("description", "No Description")
        author = ar.get("author", "Unknown Author")
        published_at = ar.get("publishedAt", "Unknown Date")

        news_item = f"Title: {title}\nDescription: {description}\nAuthor: {author}\nPublished At: {published_at}\n"
        results.append(news_item)
    
    for i, result in enumerate(results, 1):
        print(f"{i}. {result}")
        speak_text(result)
        speak_text(".... Next news item ....")

if __name__ == '__main__':
    api_key = "561497bcd1414c08a580b39e4c6670d4"  # Replace with your NewsAPI API key
    news_from_bbc(api_key)
