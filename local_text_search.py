from pathlib import Path
from openpyxl import Workbook
import json

tf_resource_search_results = {}
    
# Read 'bing_search_results_by_service.csv' and create a dictionary of TF resource names and their associated search results.
def init(file_name):

	print(f"Reading {file_name}")
	
	path = Path(file_name)
	raw = path.read_text()

	lines = raw.split('\n')
	for line in lines:
		# Skip blank lines.
		if line:
			# Split line into TF resource name and search result.
			tf_resource, search_result = line.split(',', maxsplit=1)

			# Remove quotes from search result.
			search_result = search_result.replace('"', '')

			# Add search result to dictionary.
			if tf_resource in tf_resource_search_results:
				tf_resource_search_results[tf_resource].append(search_result)
			else:
				tf_resource_search_results[tf_resource] = [search_result]

def find_articles(tf_resource_name):
	return tf_resource_search_results.get(tf_resource_name, [])

init('bing_search_results_by_service.csv')

if __name__ == '__main__':
	for result in find_articles('azurerm_app_service_certificate'):
		print(result)