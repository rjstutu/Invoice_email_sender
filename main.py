from utils import PDf_email_generator


File_listener_folder = "C:/Users/renjo/OneDrive/Desktop/Personal/Side_Business/LY/Business/Inovices/Incoming"
Output_File_Location = "C:/Users/renjo/OneDrive/Desktop/Personal/Side_Business/LY/Business/Inovices"


def main():
    PDf_email_generator(File_listener_folder,Output_File_Location).file_runner()


if __name__ == "__main__":
    main()



