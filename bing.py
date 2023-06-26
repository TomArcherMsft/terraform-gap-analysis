import os
#import requests
import json
import time

def bing_web_search(subscription_key, query):
    # set parameters
    search_url = "https://api.bing.microsoft.com/v7.0/search"
    headers = {"Ocp-Apim-Subscription-Key": subscription_key}
    params = {
        "q": query,
        "textDecorations": True,
        "count":50,
        "textFormat": "HTML"}

    # get response
    time.sleep(1)
    
    #DO NOT USE THIS FUNCTION !!!
    #response = requests.get(search_url, headers=headers, params=params)
    #response.raise_for_status()
    #return response.json()
    return None

def extract_web_pages(search_results):
    web_pages = []

    for k in search_results:
        if k == 'webPages':
            for v in search_results[k]['value']:
                endpoint = v.get('url','')
                if endpoint:
                    web_pages.append(endpoint)

    return web_pages

def find_articles(search_term):
    subscription_key = os.getenv('BING_SEARCH_KEY')
    assert subscription_key

    search_site = 'learn.microsoft.com/en-us'
    search_pattern = f"\"{search_term}\" site:{search_site}"

    content = bing_web_search(subscription_key, search_pattern)

    # The following line is useful for debugging.
    #open('bing.json','w',encoding='utf-8').write(json.dumps(content))

    web_pages = extract_web_pages(content)
    web_pages.sort()
    return web_pages

if __name__ == '__main__':
    search_term = 'azurerm_cdn_endpoint'
    print(f"\nSearching for: {search_term}")
    for web_page in find_articles(search_term):
        print(f"\t{web_page}")