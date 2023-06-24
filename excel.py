import openpyxl as excel

class ExcelWriter:

	def __init__(self, az_services):
		self.wb = excel.Workbook()
		self.az_services = az_services

	def save(self, file_name):
		self.wb.save(file_name)

	def write_sheet_azure_services(self):
		# Get workbook active sheet from the active attribute.
		sheet = self.wb.active
		sheet.title = 'Azure services'

		# Write header.
		row = 1
		c = sheet.cell(row = row, column = 1)
		c.value = 'Azure service'
		c = sheet.cell(row = row, column = 2)
		c.value = 'Article count'
		c = sheet.cell(row = row, column = 3)
		c.value = f"Terraform (azurerm) articles for Azure service"

		row = 2
		for az_service in self.az_services:
			# Write Azure service name.
			c = sheet.cell(row = row, column = 1)
			c.value = f"{az_service.name}"

			# Write article count.
			c = sheet.cell(row = row, column = 2)
			c.value = f"{len(az_service.articles)}"

			found_articles = ''
			for article_url in az_service.articles:
				if len(found_articles):
					found_articles += '\n' #'\015'
				found_articles += f"{article_url}"

			c = sheet.cell(row = row, column = 3)
			c.value = f"{found_articles}"

			row += 1

	def _write_tf_resource_row(self, sheet, row, az_service, tf_resource_name, article_url=None):
		# TODO: Use a packed param here and enumerate each column

		# Write Azure service name.
		c = sheet.cell(row = row, column = 1)
		c.value = f"{az_service.name}"

		# Write article count.
		c = sheet.cell(row = row, column = 2)
		c.value = f"{tf_resource_name}"

		# Write article URL.
		if article_url:
			c = sheet.cell(row = row, column = 3)
			c.value = f"{article_url}"

		# Write article (Y/N).
		c = sheet.cell(row = row, column = 4)
		if article_url and not az_service.is_article_excluded(article_url):
			c.value = f"Y"
		else:
			c.value = f"N"

	def write_sheet_terraform_resources(self):
		# Get workbook active sheet from the active attribute.
		sheet = self.wb.create_sheet(title='Terraform resources')

		# Write header.
		row = 1
		c = sheet.cell(row = row, column = 1)
		c.value = 'Azure service'
		c = sheet.cell(row = row, column = 2)
		c.value = 'azurerm resource name'
		c = sheet.cell(row = row, column = 3)
		c.value = 'Article that contains the azurerm resource name'
		c = sheet.cell(row = row, column = 4)
		c.value = 'Article (Y/N)'

		row = 2
		for az_service in self.az_services:
			if len(az_service.search_results):
				for tf_resource_name, article_urls in az_service.search_results.items():
					if len(article_urls):
						for article_url in article_urls:
							self._write_tf_resource_row(sheet, row, az_service, tf_resource_name, article_url)
							row += 1
					else:
						self._write_tf_resource_row(sheet, row, az_service, tf_resource_name)
						row += 1
			else:
				self._write_tf_resource_row(sheet, row, az_service, tf_resource_name)
				row += 1

	def write_sheet_excluded_articles(self, excluded_articles):
		# Write excluded articles to a new sheet.
		sheet = self.wb.create_sheet(title='Excluded articles')

		c = sheet.cell(row = 1, column = 1)
		c.value = f"Excluded articles"
		
		for row, article_url in enumerate(excluded_articles, start=2):
			
			c = sheet.cell(row = row, column = 1)
			c.value = f"{article_url}*"