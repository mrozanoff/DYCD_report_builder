import pandas as pd 
from pathlib import Path
import os

def get_attendance_COMPASS():
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\The Attendance Report.xlsx'

	df = pd.read_excel(file, skiprows=41, usecols=['Participant', 'Days Attended'])

	df = df.dropna()

	df = df[df['Days Attended'] !=0]

	os.remove(file)

	return len(df)

def get_attendance_Beacon():
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\The Attendance Report.xlsx'

	df = pd.read_excel(file, skiprows=41, usecols=['Participant', 'Actual Hours'])

	df = df.dropna()

	df = df[df['Actual Hours'] !=0]

	os.remove(file)

	return len(df)


def get_attendance_aLit():
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\The Attendance Report.xlsx'

	df = pd.read_excel(file, skiprows=41, usecols=['Participant', 'Actual Hours '])

	df = df.dropna()

	df = df[df['Actual Hours '] !=0]

	os.remove(file)

	return len(df)

def get_enrollment_Beacon():
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Official Enrollment Report.xlsx'

	df = pd.read_excel(file, skiprows=25, usecols=['Workscope', 'Total Enrollment'])

	df = df.dropna()

	# df = df[df['Actual Hours '] !=0]

	os.remove(file)

	return df['Total Enrollment'].iloc[0]

def get_ROP_CES(date):
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Rate of Participation  - Compass Elementary School.xlsx'

	df = pd.read_excel(file, skiprows=24, usecols=['Date (Monday)', 'ROP Weekly Average %'])

	df = df.dropna()

	df = df[df['Date (Monday)'] == date]

	os.remove(file)

	return df['ROP Weekly Average %'].iloc[0]

def get_ROP_CMS(date):
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Rate of Participation  - Compass Middle School.xlsx'

	df = pd.read_excel(file, skiprows=22, usecols=['Date (Monday)', 'Cumulative ROP (%)'])

	df = df.dropna()

	df = df[df['Date (Monday)'] == date]

	os.remove(file)

	return df['Cumulative ROP (%)'].iloc[0]

def get_ROP_CYEPe(date):
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Rate of Participation  - Compass Youth Empowered Programs.xlsx'

	df = pd.read_excel(file, skiprows=25)

	df.columns = [x.replace("\n", "") for x in df.columns.to_list()]

	df_1 = df[['Date(Monday)', 'Cumulative ROP (%)']]

	df_2 = df[['Date(Monday).1', 'Cumulative ROP(%)']]

	df_2 = df_2.rename(columns={'Date(Monday).1':'Date(Monday)', 'Cumulative ROP(%)':'Cumulative ROP (%)'})

	df = pd.concat([df_1, df_2])

	df = df.dropna()

	df = df[df['Date(Monday)'] == date]

	os.remove(file)

	return df['Cumulative ROP (%)'].iloc[0]


def get_ROP_CYEPhs(date):
	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Rate of Participation  - Compass Youth Empowered Programs.xlsx'

	df = pd.read_excel(file, skiprows=25)

	df.columns = [x.replace("\n", " ") for x in df.columns.to_list()]

	df_1 = df[['Date (Monday)', 'Cumulative ROP (%)']]

	df_2 = df[['Date(Monday)', 'Cumulative  ROP (%)']]

	df_2 = df_2.rename(columns={'Date(Monday)':'Date (Monday)', 'Cumulative  ROP (%)':'Cumulative ROP (%)'})

	df = pd.concat([df_1, df_2])

	df = df.dropna()

	df = df[df['Date (Monday)'] == date]

	os.remove(file)

	return df['Cumulative ROP (%)'].iloc[0]

def get_ROP_B(sheet, date):
	date = pd.to_datetime(date)

	downloads_path = str(Path.home() / "Downloads")

	file = str(downloads_path) + r'\Rate of Participation - Beacon.xlsx'

	df = pd.read_excel(file, sheet, skiprows=7, usecols=['Date (Monday)', 'Rop % (Begins Week  10)']) 

	df = df.dropna()

	df = df[df['Date (Monday)'] == date]

	if sheet == 'Sheet3':
		os.remove(file)

	return df['Rop % (Begins Week  10)'].iloc[0]

