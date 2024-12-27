import os
import time
from random import randint
import datetime

Unix = int(time.time())
dir_path = os.getcwd()
# Define the full path for the text file
filepath = os.path.join(dir_path, 'License.txt')

def create_license_file():
    with open(filepath, 'w') as file:
        # Initial state of License key
        random_initial_state = str(''.join(["{}".format(randint(0, 9)) for num in range(0, 64)]))

        # this is date of license expiration, when creating new license, it's set on May 2024
        random_initial_state = random_initial_state[:5] + "1714985405" + random_initial_state[15:]

        # this is timestamp of last use of script, when creating new license, date is circa 2010.
        random_initial_state = random_initial_state[:25] + "1279069322" + random_initial_state[35:]

        file.write(random_initial_state)
        file.close()

def update_license_file(string):
    with open(filepath, 'w') as file:
        # Initial state of License key
        updated_license = str(string)
        file.write(updated_license)
        file.close()

def License():

    # Create the text file if it doesn't exist
    if not os.path.exists(filepath):
        create_license_file()

    with open(filepath, 'r') as file:
        license_key_string = file.read()

    if not len(license_key_string) == 64:
        with open(filepath, 'w') as file:
            pass
        create_license_file()
        current_license_expiry = str(input_License())

    # protection against changing time back on PC
    current_license_unix = license_key_string[25:35]
    if int(current_license_unix) > int(Unix):
        print("License computer timestamp tampering detected!")
        exit()

    # changing old unix time to actual in license string
    license_key_string = license_key_string[:25] + str(Unix) + license_key_string[35:]
    update_license_file(license_key_string)


    current_license_expiry = license_key_string[5:15]
    if int(current_license_expiry) < int(Unix):
        print("Invalid License key.")
        current_license_expiry = int(input_License())
        if current_license_expiry < int(Unix):
            print("Invalid License key")
            input()
            exit()
        elif int(current_license_expiry) > int(Unix):
            license_key_string = license_key_string[:5] + str(current_license_expiry) + license_key_string[15:]
            update_license_file(license_key_string)

    # Convert Unix timestamp to datetime object
    dt_object = datetime.datetime.utcfromtimestamp(int(current_license_expiry))
    # Extract the year and month
    year = dt_object.year
    month = dt_object.month
    day = dt_object.day

    print("License active till " + str(day) + "." + str(month) + "." + str(year))

def input_License():
    user_license_input = input("Enter your license key:")

    #License_keys = License key : Date of expiry
    License_keys = {
    "zXmUjeWlMAR25" : 1740783600,
    "zDEsUvjKJUN25" : 1748728800,
    "lXyxms0aSEP25" : 1756677600,
    "1RfMMy1FDEC25" : 1764543600,
    "developer-key" : 2524604400,
    }
    try:
        expiry_date = License_keys[user_license_input]
    except KeyError:
        print("Invalid License key")
        input()
        exit()

    return expiry_date

if __name__ == "__main__":
    License()
