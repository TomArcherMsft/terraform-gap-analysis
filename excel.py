import openpyxl as Excel
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

class ExcelWriter:

	def __init__(self, az_services):
		self.wb = Excel.Workbook()
		self.az_services = az_services

		# Get workbook active sheet from the active attribute.
		sheet = self.wb.active
		# The first sheet will be the data summary sheet.
		sheet.title = 'Summary'

	def _auto_size_columns(self):
		for sheet_name in self.wb.sheetnames:
			for column_cells in self.wb[sheet_name].columns:
				#new_column_width = max(len(str(cell.value)) for cell in column_cells)
				new_column_width = 0
				for cell in column_cells:
					#print(f"Type:'{cell.data_type}' - Value:'{cell.value}'")
					if cell.data_type == 's':
						lines = cell.value.splitlines()
						for line in lines:
							new_column_width = max(len(str(line)), new_column_width)
					elif cell.data_type == 'n':
						new_column_width = max(len(str(cell.value)), new_column_width)

				new_column_letter = (get_column_letter(column_cells[0].column))
				if new_column_width > 0:
					self.wb[sheet_name].column_dimensions[new_column_letter].width = new_column_width #* 1.23

	def save(self, file_name):
		self._auto_size_columns()
		self.wb.save(file_name)

	def write_data(self, article_excludes):
		self._write_sheet_azure_services()
		self._write_sheet_terraform_resources()
		self._write_sheet_excluded_articles(article_excludes)
		self._write_sheet_summary()

	def _write_sheet_summary(self):
		sheet = self.wb['Summary']

		# Write summary data.
		data = [
			['Total Azure services', '=COUNTA(AzureServicesArticles[Azure service])',],
			['Azure services with articles', '=COUNTIF(AzureServicesArticles[Article count],\">0\")',],
			['% Completion', '=B2/B1',],
		]

		for row in data:
			sheet.append(row)

		sheet['B3'].number_format = '0.0%'

	def _write_sheet_azure_services(self):
		'''Write Azure services worksheet sheet.'''

		# Create worksheet.
		sheet = self.wb.create_sheet(title='Azure services')

		# Write header.
		row_number = 1
		c = sheet.cell(row_number, column = 1)
		c.value = 'Azure service'
		c = sheet.cell(row_number, column = 2)
		c.value = 'Article count'
		c = sheet.cell(row_number, column = 3)
		c.value = f"Terraform (azurerm) articles for Azure service"

		# Set row number to 2 to skip header.
		row_number = 2

		# For each Azure service.
		for az_service in self.az_services:

			# Write Azure service name.
			c = sheet.cell(row_number, column = 1)
			c.value = f"{az_service.name}"

			# Write article count.
			c = sheet.cell(row_number, column = 2)
			c.alignment = Alignment(horizontal='center')
			c.value = len(az_service.articles)

			# Format found articles into single string with newline delimiter.
			found_articles = ''
			for article_url in az_service.articles:
				if len(found_articles):
					found_articles += '\n' #'\015'
				found_articles += f"{article_url}"

			# Write article URLs.
			c = sheet.cell(row_number, column = 3)
			c.alignment = Alignment(wrapText=True)
			c.value = f"{found_articles}"

			# Increment row number.
			row_number += 1

		# Create table.
		table = self._create_table('AzureServicesArticles', row_number - 1)

		# Add table to sheet.
		sheet.add_table(table)

	def _create_table(self, table_name, row_count):
		'''Create and return a table with the given name and row count.'''

		# Create table.
		table = Table(displayName=table_name, ref=f"A1:C{row_count}")

		# Create a style with striped rows and banded columns
		style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
													showLastColumn=False, showRowStripes=True, showColumnStripes=True)

		# Apply table style.
		table.tableStyleInfo = style

		return table		

	def _write_sheet_terraform_resources(self):
		'''Write Terraform resources worksheet sheet.'''

		def _write_tf_resource_row(sheet, row_number, az_service, tf_resource_name, article_url=None):
			# Write Azure service name.
			c = sheet.cell(row_number, column = 1)
			c.value = f"{az_service.name}"

			# Write article count.
			c = sheet.cell(row_number, column = 2)
			c.value = f"{tf_resource_name}"

			# Write article URL.
			if article_url:
				c = sheet.cell(row_number, column = 3)
				c.value = f"{article_url}"

			# Write article (Y/N).
			c = sheet.cell(row_number, column = 4)
			if article_url and not az_service.is_article_excluded(article_url):
				c.value = f"Y"
			else:
				c.value = f"N"

		# Create worksheet.
		sheet = self.wb.create_sheet(title='Terraform resources')

		# Write header.
		row_number = 1
		c = sheet.cell(row_number, column = 1)
		c.value = 'Azure service'
		c = sheet.cell(row_number, column = 2)
		c.value = 'azurerm resource name'
		c = sheet.cell(row_number, column = 3)
		c.value = 'Article that contains the azurerm resource name'
		c = sheet.cell(row_number, column = 4)
		c.value = 'Article (Y/N)'

		row_number = 2
		for az_service in self.az_services:
			if len(az_service.search_results):
				for tf_resource_name, article_urls in az_service.search_results.items():
					if len(article_urls):
						for article_url in article_urls:
							_write_tf_resource_row(sheet, row_number, az_service, tf_resource_name, article_url)
							row_number += 1
					else:
						_write_tf_resource_row(sheet, row_number, az_service, tf_resource_name)
						row_number += 1
			else:
				_write_tf_resource_row(sheet, row_number, az_service, tf_resource_name)
				row_number += 1

		# Create table.
		table = self._create_table('TerraformServicesSearchResults', row_number - 1)

		# Add table to sheet.
		sheet.add_table(table)

	def _write_sheet_excluded_articles(self, excluded_articles):
		# Write excluded articles to a new sheet.
		sheet = self.wb.create_sheet(title='Excluded articles')

		c = sheet.cell(row = 1, column = 1)
		c.value = f"Articles excluded from search results."
		
		for row_number, article_url in enumerate(excluded_articles, start=2):
			c = sheet.cell(row_number, column = 1)
			c.value = f"{article_url}*"