import os

def readdirectory(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            with open(file_path, "r") as file:
                contents = file.read()
                print(f"Contents of {filename}: {contents}")

readdirectory("/path/to/directory")
