import json
from pathlib import Path

az_services = []

def init(file_name):
	'''Read services_and_related_resources.csv'''
	print(f"Reading {file_name}")

	path = Path(file_name)
	raw = path.read_text()

	lines = raw.split('\n')
	for line in lines:
		# Skip blank lines.
		if line:
			# Split line into Azure service name and TF resource name.
			az_service_name, tf_resource_name = line.split(',', maxsplit=1)

			# Remove quotes from values.
			az_service_name = az_service_name.replace('"', '')
			tf_resource_name = tf_resource_name.replace('"', '')
			
			az_service = get_az_service(az_service_name)
			az_service['keywords'].append(tf_resource_name)

def get_az_service(az_service_name):
	for az_service in az_services:
		if az_service['AzureService'] == az_service_name:
			return az_service

	# Create new Azure service dictionary.
	az_service = {}	
	az_service['AzureService'] = az_service_name
	az_service['keywords'] = []

	az_services.append(az_service)

	return az_service

def print_az_services():
	for az_service in az_services:
		print(f"{az_service['AzureService']}")
		for tf_resource in az_service['keywords']:
			print(f"\t{tf_resource}")

init('services_and_related_resources.csv')

if __name__ == '__main__':
	dict = {}
	dict['TerraformServices'] = az_services

	open(file='delora.json',OpenTextMode='w',encoding='utf-8').write(json.dumps(dict))