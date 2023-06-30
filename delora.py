import json
from pathlib import Path

'''Module to create a JSON file of Azure services and related Terraform resources.'''

# List of Azure services and related Terraform resources.
az_services = []

def init(file_name):
	'''Initialize the Azure services and related resources.'''

	# Read the Azure services and related resources from the CSV file.
	path = Path(file_name)
	raw = path.read_text()

	# Split raw text into lines.
	lines = raw.split('\n')

	# For each line in the CSV file...
	for line in lines:

		# Skip blank lines.
		if line:

			# Split line into Azure service name and TF resource name.
			az_service_name, tf_resource_name = line.split(',', maxsplit=1)

			# Remove quotes from values.
			az_service_name = az_service_name.replace('"', '')
			tf_resource_name = tf_resource_name.replace('"', '')
			
			# Get the Azure service dictionary for Azure service name 
			# or get a new Azure service dictionary if it doesn't already exist.
			az_service = get_az_service_by_name(az_service_name)

			# Add TF resource name to Azure service dictionary if it doesn't 
			# already exist.
			add_resource(az_service, tf_resource_name)

	# Sort the TF resources for each Azure service.
	sort_tf_resources()

def add_resource(az_service, tf_resource_name):
	'''Add TF resource name to Azure service dictionary if it doesn't already exist.'''

	# If TF resource name is not already in Azure service dictionary...
	if tf_resource_name not in az_service['keywords']:

		# Add TF resource name to Azure service dictionary.
		az_service['keywords'].append(tf_resource_name)

def get_az_service_by_name(az_service_name):
	'''Get Azure service dictionary for specified Azure service name.'''

	# For each Azure service...
	for az_service in az_services:

		# If Azure service name matches specified Azure service name...
		if az_service['AzureService'] == az_service_name:

			# Return Azure service dictionary.
			return az_service

	# Create new Azure service dictionary.
	az_service = {}	
	az_service['AzureService'] = az_service_name
	az_service['keywords'] = []

	# Add Azure service dictionary to list of Azure services.
	az_services.append(az_service)

	# Return Azure service dictionary.
	return az_service

def sort_tf_resources():
	'''Sort the TF resources for each Azure service.'''

	# For each Azure service...
	for az_service in az_services:
		# Sort the TF resources for each Azure service.
		az_service['keywords'].sort()

def print_az_services():
	'''Print all Azure services and related resources.'''

	# For each Azure service...
	for az_service in az_services:

		# Print the Azure service.
		print_az_service(az_service)

def print_az_service(az_service):
	'''Print a specific Azure service and related resources.'''

	# Print the Azure service name.
	print(f"{az_service['AzureService']} resources:")

	# For each TF resource for the Azure service...
	for tf_resource in az_service['keywords']:

		# Print the TF resource name.
		print(f"\t{tf_resource}")

# Initialize the Azure services and related resources.
init('services_and_related_resources.csv')

if __name__ == '__main__':

	# Create dictionary to hold all services and related resources.
	# Reporting requires a dictionary with a single key
	# called 'TerraformServices' that contains a list of all services.
	dict = {}
	dict['TerraformServices'] = az_services

	# Write all services and related resources to JSON file.
	#open('delora.json','w',encoding='utf-8').write(json.dumps(dict))
	print_az_services()

	# Test to write a specific service and related resources to JSON file.
	#az_service = get_az_service_by_name('App Service (Web Apps)')
	#print_az_service(az_service)