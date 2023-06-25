from pathlib import Path
import json
import bing
from excel import ExcelWriter
from colorama import Fore, Back, Style
  
INPUT_FILE = 'azurerm.json'

path = Path(INPUT_FILE)
raw = path.read_text()
content = json.loads(raw)

az_services = []

# TODO: Specify Azure service names to skip - useful for debugging.
service_excludes = []
service_excludes = ['API Management', ]

article_excludes = [
	'https://learn.microsoft.com/en-us/answers',
	'https://learn.microsoft.com/en-us/azure/developer/terraform/provider-version-history-azurerm',
	]

class AzureService:
	def __init__(self, name):
		self.name = name
		self.search_results = {}
		self.articles = []

	def is_article_excluded(self, article_url, exact_match=True):
		article_is_excluded = False
		if exact_match and article_url in article_excludes:
			article_is_excluded = True
		else:
			for article_exclude in article_excludes:
				if article_url.startswith(article_exclude):
					article_is_excluded = True
					break
		
		return article_is_excluded
	
	def add_search_results(self, tf_resource_name, article_urls):
		self.search_results[tf_resource_name] = article_urls

		# If any article URL is not in the list of articles, add it.
		for article_url in article_urls:
			if article_url not in self.articles:
				if not self.is_article_excluded(article_url, exact_match=False):
					self.articles.append(article_url)
		
	def __str__(self):
		return f"{self.name}\n{self.search_results}\n{self.articles}"

def dump_azure_services(*az_service_names):
	for az_service in az_services:
		if az_service.name in az_service_names:
			print(f"\n{az_service}")

def write_to_excel():
	excelWriter = ExcelWriter(az_services)
	excelWriter.write_data(article_excludes)

	file_name = f"{Path(__file__).stem}.xlsx"
	excelWriter.save(file_name)
	print(Fore.GREEN + f"\nGap analysis report generated: '{file_name}'")

def main():
	count_az_services = 0

	print(Fore.WHITE)

	for az_service_name in content:
		if not az_service_name in service_excludes:
			az_service = AzureService(name=az_service_name)

			print(f"\nProcessing '{az_service_name}'...")

			for i, tf_resource_name in enumerate(content[az_service_name]):
				print(f"\tSearching for '{tf_resource_name}'...", end='')

				found_articles = bing.find_articles(tf_resource_name)
				print(f"{len(found_articles)} search result(s).")
				az_service.add_search_results(tf_resource_name=tf_resource_name, article_urls=found_articles)

			az_services.append(az_service)

			count_az_services += 1
			
			# TODO: Use this to speed up testing.
			if count_az_services == 11:
				break
		else:
			print(Fore.BLUE + f"\nSkipping '{az_service_name}'")
			print(Fore.WHITE)
			
	write_to_excel()

	print(Fore.WHITE)
main()