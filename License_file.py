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
        random_initial_state = str(''.join(["{}".format(randint(0, 9)) for num in range(0, 512)]))

        # this is date of license expiration, when creating new license, it's set on May 2024
        random_initial_state = random_initial_state[:225] + "1714985405" + random_initial_state[235:]

        # this is timestamp of last use of script, when creating new license, date is circa 2010.
        random_initial_state = random_initial_state[:385] + "1279069322" + random_initial_state[395:]

        file.write(random_initial_state)
        file.close()

def update_license_file(string):
    with open(filepath, 'w') as file:
        # Initial state of License key
        updated_license = str(string)
        file.write(updated_license)
        file.close()

def randomize_unused_digits():
    with open(filepath, 'r') as file:
        license_key_string = file.read()
        first_random = str(''.join(["{}".format(randint(0, 9)) for num in range(0, 225)]))
        second_random = str(''.join(["{}".format(randint(0, 9)) for num in range(0, 150)]))
        third_random = str(''.join(["{}".format(randint(0, 9)) for num in range(0, 117)]))
        current_license_unix = license_key_string[385:395]
        current_license_expiry = license_key_string[225:235]
        randomized_string = str(first_random) + str(current_license_expiry) + str(second_random) + str(current_license_unix) + str(third_random)
        file.close()
        with open(filepath, 'w') as file2:
            file2.write(randomized_string)
            file2.close()

def License():

    # Create the text file if it doesn't exist
    if not os.path.exists(filepath):
        create_license_file()

    with open(filepath, 'r') as file:
        license_key_string = file.read()

    if not len(license_key_string) == 512 or not license_key_string.isdigit():
        with open(filepath, 'w') as file:
            pass
        print("License file is corrupted or verification failed.")
        print("Resetting License file.")
        os.remove(filepath)
        time.sleep(0.2)
        create_license_file()
        input()
        exit()

    # protection against changing time back on PC
    current_license_unix = license_key_string[385:395]
    if int(current_license_unix) > int(Unix):
        print("License computer timestamp tampering detected!")
        print("Resetting License file.")
        os.remove(filepath)
        create_license_file()
        input()
        exit()

    # changing old unix time to actual in license string
    license_key_string = license_key_string[:385] + str(Unix) + license_key_string[395:]
    update_license_file(license_key_string)

    current_license_expiry = license_key_string[225:235]
    if int(current_license_expiry) < int(Unix):
        print("Invalid License key.")
        create_license_file()
        current_license_expiry = int(input_License())
        if current_license_expiry < int(Unix):
            create_license_file()
            print("Invalid License key")
            input()
            exit()
        elif int(current_license_expiry) > int(Unix):
            license_key_string = license_key_string[:225] + str(current_license_expiry) + license_key_string[235:]
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
    "zDEsUvjkJAN25" : 1738364400,
    "lXyxms0aFEB25" : 1740783600,
    "1RfMMy1fMAR25" : 1743458400,
    "vOAO6VQ5APR25" : 1746050400,
    "dFrPP565MAY25" : 1748728800,
    "v9NPQOubJUN25" : 1751320800,
    "BOKIzrjaJUL25" : 1753999200,
    "ZPbzHbZsAUG25" : 1756677600,
    "mxT2rr64SEP25" : 1759269600,
    "PpnCN4CqOCT25" : 1761951600,
    "TP56FIXlNOV25" : 1764543600,
    "e4xLz0amDEC25" : 1767222000,
    "developer-key" : Unix + 600,
    }
    try:
        expiry_date = License_keys[user_license_input]
    except KeyError:
        print("Invalid License key")
        input()
        exit()

    return expiry_date

if __name__ == "__main__":
    try:
        License()
    except:
        print("License file is corrupted or verification failed.")
        print("Resetting License file.")
        os.remove(filepath)
        time.sleep(0.2)
        create_license_file()
        input()
        exit()
