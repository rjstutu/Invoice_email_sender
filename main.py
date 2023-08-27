from utils import PDf_email_generator
import time
from config import File_listener_folder, Output_File_Location, sleep_time


def main():
    PDf_email_generator(File_listener_folder,Output_File_Location).file_runner()


if __name__ == '__main__':
    while True:
        try:
            main()
            time.sleep(sleep_time)  # Wait for 500 seconds before the next iteration
        except Exception as e:
            print(f"An error occurred: {e}")
            time.sleep(sleep_time)  # If there's an error, wait for 500 seconds before retrying


