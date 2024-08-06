import os
import openpyxl
import sqlite3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
import time

# Load test cases from Excel file
def load_test_cases(file_path):
    print(f"Checking if the file exists at: {file_path}")
    if os.path.exists(file_path):
        print("File found, loading workbook...")
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        test_cases = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            time_value = row[2].strftime("%H:%M") if isinstance(row[2], datetime.time) else str(row[2])
            test_cases.append({
                'test_case_id': row[0],
                'name': row[1],
                'time': time_value,
                'expected_result': row[3]
            })

        return test_cases
    else:
        raise FileNotFoundError(f"The file {file_path} does not exist.")

# Initialize SQLite database
def init_db():
    db_path = r'C:/Users/HP/OneDrive/Desktop/New12[1]/New123/appointments.db'
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS appointments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                patient_name TEXT NOT NULL,
                doctor_name TEXT NOT NULL,
                specialty TEXT NOT NULL,
                appointment_time TEXT NOT NULL
            )
        ''')
        conn.commit()
        return conn
    except sqlite3.OperationalError as e:
        print(f"Database error: {e}")
        raise

# Book appointment in the database
def book_appointment(conn, patient_name, doctor_name, specialty, appointment_time):
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO appointments (patient_name, doctor_name, specialty, appointment_time)
        VALUES (?, ?, ?, ?)
    ''', (patient_name, doctor_name, specialty, appointment_time))
    conn.commit()
    return True

# Initialize WebDriver
driver = webdriver.Chrome()

# Open the HTML file
driver.get(r"file:///C:/Users/HP/OneDrive/Desktop/New12[1]/New123/appointment.html")

# Load test cases
file_path = "C:/Users/HP/OneDrive/Desktop/New12[1]/New123/appointment_test_cases.xlsx"
try:
    test_cases = load_test_cases(file_path)
except FileNotFoundError as e:
    print(e)
    driver.quit()
    exit()

# Initialize the database
conn = init_db()

# Execute test cases
for case in test_cases:
    try:
        print(f"Executing test case {case['test_case_id']}...")
        name_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "name")))
        time_select = Select(WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "time"))))
        result_div = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "result")))

        name_input.clear()
        name_input.send_keys(case['name'])

        time_select.select_by_value(case['time'])
        time.sleep(1)  # Add a delay of 1 second

        # Submit the form
        driver.find_element(By.ID, "appointmentForm").submit()
        time.sleep(2)  # Add a delay of 2 seconds

        # Check the result
        actual_result = result_div.text
        if actual_result == case['expected_result']:
            print(f"{case['test_case_id']}: Passed")
        else:
            print(f"{case['test_case_id']}: Failed - Expected: {case['expected_result']}, Got: {actual_result}")

        # Book appointment in database if the time is valid
        if 'successfully' in actual_result.lower():
            doctor_name = "Dr. Ramesh"
            specialty = "General Physician"
            book_appointment(conn, case['name'], doctor_name, specialty, case['time'])

        time.sleep(3)  # Add a delay of 3 seconds before moving to the next test case

    except Exception as e:
        print(f"Error in test case {case['test_case_id']}: {e}")

conn.close()
driver.quit()
