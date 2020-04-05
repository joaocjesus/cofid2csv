import glob
import csv
import sys

from os.path import join

csv_path = './csv_exports'
csv_files = glob.glob(join(csv_path, '*.csv'))
csv_count = len(csv_files)
CSVs = []
columns = []
rows_count = 0


class CSV:
	def __init__(self, filename):
		self.filename = filename
		self.indexes = []
		self.content = self.process_data()

	def process_data(self):
		global rows_count
		content = []
		with open(self.filename, 'r') as f:
			lines = sum(1 for line in f)
			if lines > rows_count:
				rows_count = lines
		with open(self.filename, 'r') as f:
			reader = csv.reader(f)
			header = next(reader)
			print('Creating object for ' + self.filename + '...')
			for col_index, header_cell in enumerate(header):
				if header_cell not in columns:
					self.indexes.append(col_index)
					columns.append(header_cell)
			for row in reader:
				content.append(row)
		return content


def print_percent(percent):
	text = 'Processing merge... ' + str('%.2f' % percent + '%')
	sys.stdout.write('\r')
	sys.stdout.flush()
	sys.stdout.write(text)
	sys.stdout.flush()


def get_csv(filename):
	for csv_file in CSVs:
		if csv_file.filename == filename:
			return csv_file


print('Creating objects from ' + str(csv_count) + ' csv files...')

for file in csv_files:
	CSVs.append(CSV(file))

output = []
output_row = []

print('Processing merge...')

# Process header
for column in columns:
	output_row.append(column)
output.append(output_row)

# Process rest of rows
for row_index in range(rows_count - 1):
	output_row = []
	for csvObj in CSVs:
		for column_index, cell in enumerate(csvObj.content[row_index]):
			if column_index in csvObj.indexes:
				output_row.append(cell)
	output.append(output_row)

print('Writing file...')

with open('merged_csvs.csv', 'w') as fw:
	filewriter = csv.writer(fw)
	filewriter.writerows(output)
