from duckduckgo_search import DDGS
import time

def find_articles(search_term):
	time.sleep(1)

	search_site = 'learn.microsoft.com/en-us'
	search_pattern = f"\"{search_term}\" site:{search_site}"

	found_articles = []
	with DDGS() as ddgs:
		for r in ddgs.text(search_pattern, region='wt-wt', safesearch='Off', timelimit='y'):
			article_url = r['href']
			if not article_url in found_articles:
				found_articles.append(article_url)

	found_articles.sort()
	return found_articles

if __name__ == '__main__':
	for article_url in find_articles('azurerm_cdn_endpoint'):
		print(f"\t{article_url}")