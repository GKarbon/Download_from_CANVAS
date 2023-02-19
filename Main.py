import zipfile
import shutil
import os
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver


def treemove(src, dst):
    """Move a directory tree to a destination directory, skipping existing files.

    Args:
        src (file path): Path to the source directory.
        dst (file path): Path to the destination directory.
    """
    with os.scandir(src) as itr:
        entries = list(itr)
    for srcentry in entries:
        srcname = os.path.join(src, srcentry.name)
        dstname = os.path.join(dst, srcentry.name)
        try:
            if srcentry.is_dir():
                if os.path.exists(dstname):
                    treemove(srcname, dstname)
                else:
                    os.mkdir(dstname)
                    treemove(srcname, dstname)
                continue
            else:
                # Skip existing files
                if os.path.exists(dstname):
                    continue
                shutil.copy(srcname, dstname)
        except OSError as e:
            print(e)
            pass


class Download_Extract_Move:
    """Main class for downloading, extracting and moving files."""

    def __init__(self, username, password):
        self.username = username  # Could replace with your own username
        self.password = password  # Could replace with your own password
        self.courselink_dic = {}
        self.course_names = []
        self.file_list = []
        self.extract_path = r"home\user\Desktop\temp"     # Replace with your own path
        self.download_path = r"home\user\Downloads"       # Replace with your own path (Default browser's download path)
        self.final_path = r"home\user\Desktop\testing"    # Replace with your own path (Have to be existed)
        self.driver = webdriver.Edge(EdgeChromiumDriverManager().install())

    def login(self):
        """Get and sign in to the canvas website."""
        self.driver.get("https://canvas.cityu.edu.hk/")
        self.driver.implicitly_wait(10)
        self.driver.find_element(By.XPATH, "//input[@name='username']").send_keys(
            self.username
        )
        self.driver.find_element(By.XPATH, "//input[@name='password']").send_keys(
            self.password
        )
        self.driver.find_element(By.XPATH, "//input[@name='password']").submit()
        self.driver.implicitly_wait(10)

    def get_courselink_list_w_name(self):
        """Get the course links and names."""
        self.driver.find_element(By.ID, "global_nav_courses_link").click()
        course_list = self.driver.find_elements(
            By.XPATH, "//div[@class='tray-with-space-for-global-nav']/div/ul/li/a"
        )
        index = 1
        for course in course_list:
            courselink = course.get_attribute("href")
            if courselink in [
                # Exclude the courses that you don't want to download
                # (Courses that don't have files, the legacy courses, and the All courses button)
                r"https://canvas.cityu.edu.hk/courses/35721",   # Legacy course (Don't have files, replace with your own course link)
                r"https://canvas.cityu.edu.hk/courses/46576",   # Legacy course (Don't have files, replace with your own course link)
                r"https://canvas.cityu.edu.hk/courses",         # All courses button (Must keep)
            ]:
                continue
            self.courselink_dic.update({index: courselink})
            self.course_names.append(course.get_attribute("text"))
            index += 1

    def download(self, index, link):
        """Function for downloading files from the course page.

        Args:
            index (int): Count for window_handles
            link (str): Link for the course page
        """
        # open a new tab
        self.driver.execute_script("window.open('');")
        # switch to the new tab
        self.driver.switch_to.window(self.driver.window_handles[index])
        self.driver.get(link)
        self.driver.implicitly_wait(10)
        self.driver.find_element(By.XPATH, "//a[contains(text(),'Files')]").click()
        self.driver.implicitly_wait(10)
        filerows = self.driver.find_elements(By.XPATH, "//div[@class='ef-item-row']")
        ActionChains(self.driver).key_down(Keys.CONTROL).perform()
        for file in filerows:
            file.click()
        ActionChains(self.driver).key_up(Keys.CONTROL).perform()
        # press the download button
        self.driver.find_element(By.XPATH, "//button[@title='Download']").click()
        WebDriverWait(self.driver, 360).until_not(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[@class='alert alert-info']",
                )
            )
        )

    def waiting_download(self):
        # check the filname in the download folder with .download extension
        # if the filename is not in the download folder, then the download is finished
        while True:
            if os.path.exists(
                self.download_path
                + r"\course_files_export ({}).zip".format(
                    self.courselink_dic.__len__() - 1
                )
            ):
                break

    def rename(self):
        """Rename the downloaded files."""
        for index in range(len(self.course_names)):
            # The first file need to be renamed is called course_files_export
            # Other files are called course_files_export(1)、course_files_export(2)...
            if index == 0:
                # Rename as first element of course_names
                os.rename(
                    self.download_path + r"\{}.zip".format("course_files_export"),
                    self.download_path + r"\{}.zip".format(self.course_names[index]),
                )
            else:
                # Rename as the index element of course_names
                os.rename(
                    self.download_path
                    + r"\{} ({}).zip".format("course_files_export", index),
                    self.download_path + r"\{}.zip".format(self.course_names[index]),
                )
        self.file_list = [
            os.path.join(self.download_path, "{}.zip".format(course_name))
            for course_name in self.course_names
        ]

    def extract(self):
        """Extract the downloaded files to the extract_path."""
        for file in self.file_list:
            with zipfile.ZipFile(file, "r") as zip_ref:
                zip_ref.extractall(
                    self.extract_path
                    + r"\{}".format(os.path.splitext(os.path.basename(file))[0])
                )
            os.remove(file)

    def move(self):
        """Move the extracted files to the final_path."""
        for _ in self.course_names:
            treemove(
                self.extract_path,
                self.final_path,
            )
        shutil.rmtree(self.extract_path)

    def start(self):
        """Start the program."""
        self.login()
        self.get_courselink_list_w_name()
        for index, link in self.courselink_dic.items():
            self.download(index, link)
        self.waiting_download()
        self.rename()
        self.extract()
        self.move()
        self.driver.quit()


if __name__ == "__main__":
    # If you set the username and password in the code, you can delete these two lines
    username = input("Please enter your username: ")
    password = input("Please enter your password: ")
    downloader = Download_Extract_Move(username, password)
    downloader.start()
