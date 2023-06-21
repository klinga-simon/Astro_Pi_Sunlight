# Astro_Pi_Sunlight
# This code can tell you whether the ISS was in sunlight or darkness, based on the latitude, longitude, date and time of the measurement. It takes the required data from Excel and then gives you back a csv file.

import openpyxl
import csv
import datetime as dt
from suntime import Sun
from pytz import timezone

# Load data from Excel file
workbook = openpyxl.load_workbook('iss_data.xlsx')
sheet = workbook.active

# Set the timezone for the timestamp (Greenwich Mean Time with an offset of +1)
timezone_str = 'Etc/GMT+1'

# Create a list to store the result rows
result_rows = []

# Iterate through each row in the data
for row in sheet.iter_rows(min_row=2, values_only=True):
    date = row[0].date()
    timestamp = row[1]
    latitude = row[2]
    longitude = row[3]

    # Combine date and time
    timestamp = dt.datetime.combine(date, timestamp)

    # Set the timezone for the timestamp
    timestamp = timezone(timezone_str).localize(timestamp)

    # Calculate sunrise and sunset times
    sun = Sun(latitude, longitude)
    sunrise = sun.get_sunrise_time(timestamp)
    sunset = sun.get_sunset_time(timestamp)

    # Convert sunrise and sunset times to the given timezone
    sunrise = sunrise.astimezone(timezone(timezone_str))
    sunset = sunset.astimezone(timezone(timezone_str))

    # Check if the ISS is in sunlight or darkness
    if sunrise <= timestamp <= sunset:
        status = "In sunlight"
    else:
        status = "In darkness"

    # Create a row for the result
    result_row = [date, timestamp, latitude, longitude, status]

    # Add the row to the result list
    result_rows.append(result_row)

# Define the output CSV file path
output_csv_file = "iss_sunlight_data.csv"

# Write the result to a CSV file
with open(output_csv_file, 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Date', 'Timestamp', 'Latitude', 'Longitude', 'Status'])
    writer.writerows(result_rows)

print(f"Data exported to {output_csv_file} successfully.")

