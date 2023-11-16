import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup


CREDENTIALS_FILE = 'viarmo-task-405222-69a05e66fe0d.json'
spreadsheet_id = '1_Dj50lOHGbkxv_blxjUYalyqkuqfongXZ_ZoQOGWURo'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
	CREDENTIALS_FILE,
	['https://spreadsheets.google.com/feeds',
	 'https://www.googleapis.com/auth/spreadsheets',
	 'https://www.googleapis.com/auth/drive.file',
	 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

def get_html(URL):
	try:
		headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36", "content-type": "text"}
		r = requests.get(URL, headers=headers)
		r.raise_for_status()
		return r
	except(requests.RequestException, ValueError):
		print('Server error')
		return False

def get_characteristics(HTML):
	soup = BeautifulSoup(HTML, 'html.parser')
	characteristic_keys = []
	characteristic_keys_dirty = list(soup.find_all('span', class_='_2NZVF _15CQ5 _32rOe _25vcL'))
	for key in characteristic_keys_dirty:
		characteristic_keys.append(key.getText())

	characteristic_values = []
	characteristic_values_dirty = soup.find_all('div', class_='_198Aj cXkP_ _3wss4 _1XOOj')
	for i in range(len(characteristic_keys)):
		value = characteristic_values_dirty[i].getText().split(f'{characteristic_keys[i]}')[1]
		characteristic_values.append(value)
	
	return characteristic_keys, characteristic_values

def insert_into_table(keys, values, row):
	available_keys = service.spreadsheets().values().get(
		spreadsheetId=spreadsheet_id,
		range='1:1',
		majorDimension='ROWS').execute()['values'][0]
	len_available_keys = len(available_keys)
	for characteristic_id in range(len(keys)):
		col_to_put = -1
		for i in range(len(available_keys)):
			if keys[characteristic_id] == available_keys[i]:
				col_to_put = i
				break
		if col_to_put == -1:
			column = chr(65 + len_available_keys) \
				if (65 + len_available_keys) <= 90 \
				else (chr((65 + len_available_keys // 26 - 1)) \
					+ chr((65 + len_available_keys % 26))) # будет актульно при доступном количестве столбоцов > 26
			service.spreadsheets().values().batchUpdate(
				spreadsheetId = spreadsheet_id,
				body = {
					"valueInputOption": "USER_ENTERED",
					"data": [
						{"range": f"{column}1",
						"values": [[keys[characteristic_id]]]}, 
						{"range": f"{column}{row}",
						"values": [[values[characteristic_id]]]}]}).execute()
			len_available_keys += 1
		else:
			column = chr(65 + col_to_put) \
				if (65 + col_to_put) <= 90 \
				else (chr((65 + col_to_put // 26 - 1)) \
					+ chr((65 + col_to_put % 26))) # будет актульно при доступном количестве столбоцов > 26
			service.spreadsheets().values().batchUpdate(
					spreadsheetId = spreadsheet_id,
					body = {
						"valueInputOption": "USER_ENTERED",
						"data": [
							{"range": f"{column}{row}",
							"values": [[values[characteristic_id]]]}]}).execute()


links = service.spreadsheets().values().get(
	spreadsheetId=spreadsheet_id,
	range='A:A',
	majorDimension='COLUMNS').execute()['values'][0][1:]
for row in range(len(links)):
	if not links[row]:
		continue
	HTML = get_html(links[row])
	characteristic_keys, characteristic_values = get_characteristics(HTML.text)
	insert_into_table(characteristic_keys, characteristic_values, row + 2)
