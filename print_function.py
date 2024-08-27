from selenium import webdriver
import pyautogui
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import os
import time
import create_absolute_path
from core_function import *;
import pandas as pd
import shutil
import re

class printFunction:

    def __init__(self,ab_path,ch_loc,inp_loc,out_loc):
        self.wd = ab_path
        self.chrome_location = ch_loc
        self.input_folder_path = inp_loc
        self.output_folder_path = out_loc

    def print_urlpages_default(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")
        
        #chrome_options = load_chrome_options(chrome_location)
        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        for url in urls:
            try:
                driver=load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title

                custom_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                custom_icon_path = os.path.join(self.wd, "images", "customicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                apply_icon_path = os.path.join(self.wd, "images", "applyicon.png")
                p = pyautogui.locateOnScreen(apply_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                print_icon_path = os.path.join(self.wd, "images", "printicon.png")
                p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(5)

    """def Preview_load_time(): #it is working and finds time for default preview also
        loading_times = []  # List to store the loading times
        urls_list = []  # List to store the URLs

        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")

        urls_path = os.path.join(input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        for url in urls:
            try:
                driver = load_chrome_options(chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title

                # Coordinates of the pixel to check (replace with actual coordinates)
                x, y = 100, 200

                # Get the initial color of the pixel
                initial_color = pyautogui.pixel(x, y)
                print(f"Initial color of the pixel at ({x}, {y}): {initial_color}")

                extension_icon_path = os.path.join(wd, "images", "extensionicon.png")
                v = pyautogui.locateOnScreen(extension_icon_path, confidence=0.8)

                # Start the timer
                start_time = time.time()

                pyautogui.click(x=v[0], y=v[1], clicks=1, interval=0.0, button="left")

                time.sleep(20)

                # Wait for the print button to be enabled
                print_button_path = os.path.join(wd, "images", "printicon.png")
                while not pyautogui.locateOnScreen(print_button_path, confidence=0.8):
                    time.sleep(0.1)  # Check every 100 milliseconds

                # End the timer
                end_time = time.time()

                # Calculate the time taken for the extension to load
                loading_time = end_time - start_time
                loading_times.append(loading_time)  # Store the time in the list
                urls_list.append(url.strip())  # Store the URL in the list

                time.sleep(10)

                print(f"Print button is enabled.")

                full_name = str(title).replace("|", "") + ".png"
                screenshot_path = os.path.join(output_folder_path, full_name)
                ss = pyautogui.screenshot()
                ss.save(screenshot_path)
            except Exception as e:
                print(f"Error for URL: {url.strip()} - {str(e)}")
                loading_times.append(None)  # Store None for the failed URL
                urls_list.append(url.strip())  # Store the URL in the list
            finally:
                driver.quit()
                time.sleep(5)

            # Append the current URL and loading time to the Excel file
            df = pd.DataFrame({
                "URL": [url.strip()],
                "Loading Time (s)": [loading_times[-1]]
            })
            output_excel_path = os.path.join(output_folder_path, "loading_times.xlsx")
            with pd.ExcelWriter(output_excel_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

        print(f"Loading times have been saved to {output_excel_path}")"""


    def Preview_load_time(self): #it is working and finds time for default preview also
        loading_times = []  # List to store the loading times
        urls_list = []  # List to store the URLs

        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")

        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        for url in urls:
            try:
                driver = load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title

                # # Coordinates of the pixel to check (replace with actual coordinates)
                # x, y = 100, 200

                # # Get the initial color of the pixel
                # initial_color = pyautogui.pixel(x, y)
                # print(f"Initial color of the pixel at ({x}, {y}): {initial_color}")

                extension_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                v = pyautogui.locateOnScreen(extension_icon_path, confidence=0.8)

                # Start the timer
                start_time = time.time()

                pyautogui.click(x=v[0], y=v[1], clicks=1, interval=0.0, button="left")

                time.sleep(20)

                # Wait for the print button to be enabled
                print_button_path = os.path.join(self.wd, "images", "printicon.png")
                while(not pyautogui.locateOnScreen(print_button_path, confidence=0.8)):
                    time.sleep(0.1)
                    
                try:
                    while not pyautogui.locateOnScreen(print_button_path, confidence=0.8):
                        print("Waiting for the print button to appear...")
                        pyautogui.sleep(0.5)  # Wait for 1 second before retrying
                except pyautogui.ImageNotFoundException:
                    print("Could not locate the print button on the screen.")
                # End the timer
                end_time = time.time()

                # Calculate the time taken for the extension to load
                loading_time = end_time - start_time
                loading_times.append(loading_time)  # Store the time in the list
                urls_list.append(url.strip())  # Store the URL in the list

                time.sleep(10)

                print(f"Print button is enabled.")

                # full_name = str(title).replace("|", "") + ".png"
                # screenshot_path = os.path.join(output_folder_path, full_name)
                # ss = pyautogui.screenshot()
                # ss.save(screenshot_path)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(5)
            print(loading_times)

        # Create a DataFrame from the URLs and loading times
        df = pd.DataFrame({
            "URL": urls_list,
            "Loading Time (s)": loading_times
        })

        print(urls_list)

        # Save the DataFrame to an Excel file
        output_excel_path = os.path.join(self.output_folder_path, "loading_times.xlsx")
        df.to_excel(output_excel_path, index=False)

        print(f"Loading times have been saved to {output_excel_path}")

    def print_urlpages_fontsize(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")

        
        #chrome_options = load_chrome_options(chrome_location)
        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        for url in urls:
            try:
                driver=load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title

                custom_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                font_icon_path = os.path.join(self.wd, "images", "fonts_icon.png")
                p = pyautogui.locateOnScreen(font_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                font_icon_path = os.path.join(self.wd, "images", "font_click.png")
                p = pyautogui.locateOnScreen(font_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                apply_icon_path = os.path.join(self.wd, "images", "font_change.png")
                p = pyautogui.locateOnScreen(apply_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                print_icon_path = os.path.join(self.wd, "images", "printicon.png")
                p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(5)




    def print_urlpages_custom(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")

        
        #chrome_options = load_chrome_options(chrome_location)
        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        for url in urls:
            try:
                driver=load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title

                custom_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                custom_icon_path = os.path.join(self.wd, "images", "customicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                apply_icon_path = os.path.join(self.wd, "images", "applyicon.png")
                p = pyautogui.locateOnScreen(apply_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)

                print_icon_path = os.path.join(self.wd, "images", "printicon.png")
                p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(20)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(5)

    def prn_urlpages_default(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")
    
        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")
    
        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()
        
        #create a new folder to save the PRN files
        
        prn_folder_path = os.path.join(self.output_folder_path, "prn_files")
        if not os.path.exists(prn_folder_path):
            os.makedirs(prn_folder_path)
        
        #create a new file to save the PRN files path
        prn_files_path = os.path.join(self.output_folder_path, "prn_files_list.txt")
        prn_files_count = 0

        for index, url in enumerate(urls, start=1):
            driver = None  # Initialize driver to None
            try:
                driver = load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                driver.get(url.strip())
                title = driver.title
    
                custom_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                p = pyautogui.locateOnScreen(custom_icon_path, confidence=0.8)
                pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                time.sleep(30)
    
                print_icon_path = os.path.join(self.wd, "images", "print_icon.png")
                
                p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                try:
                    pyautogui.click(x=p[0], y=p[1], clicks=1, interval=0.0, button="left")
                except Exception as e:
                    print("Preview not loaded: ", e)
                    continue
                time.sleep(25)
    
                # Automate the print process to save the PRN file
                # pyautogui.hotkey('ctrl', 'p')
                # time.sleep(2)  # Wait for the print dialog to open
    
                # Save the PRN file with a unique name
                prn_file_name = f"recipe{index}_{title.replace('|', '')}.prn"
                prn_file_path = os.path.join(prn_folder_path, prn_file_name)
                pyautogui.typewrite(prn_file_path)
                pyautogui.press('enter')
                time.sleep(5)  # Wait for the file to be saved

                # Append the path of the PRN file to the PRN files list file
                with open(prn_files_path, "a") as prn_files_list_file:
                    prn_files_list_file.write(f"{prn_file_path}\n")
                prn_files_count += 1

                print(f"PRN file saved as {prn_file_path}")
    
            except TimeoutException as e:
                print("Timeout")
            finally:
                if driver is not None:  # Check if driver was initialized
                    driver.quit()
                time.sleep(5)
        print(f"Number of PRN files saved: {prn_files_count}")

    def sanitize_filename(self,filename):
        # Replace any non-alphabetic characters (excluding underscores) with an underscore
        sanitized_name = re.sub(r'[^a-zA-Z_]', '_', filename)
        
        # Ensure the file has a valid extension
        name, ext = os.path.splitext(sanitized_name)
        
        # Return the sanitized name with the original extension
        return f"{name}{ext}"

    def store_good_preview_page(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")


        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        #create a folder to store the default preview pdfs
        default_preview_folder = os.path.join(self.output_folder_path, "good_previews")
        if os.path.exists(default_preview_folder):
            #delete the folder and recreate it
            shutil.rmtree(default_preview_folder)
        
        os.makedirs(default_preview_folder)
        preview_not_loaded_count = 0

        #create files to store the file_name
        default_preview_files_path = os.path.join(self.output_folder_path, "urls_title_names.txt")
        if os.path.exists(default_preview_files_path):
            os.remove(default_preview_files_path)


        for url in urls:
            try:
                driver=load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                #set chrome size as 75%
                # driver.set_window_size(1024, 768)
                driver.get(url.strip())
                title = driver.title

                extension_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                v = pyautogui.locateOnScreen(extension_icon_path, confidence=0.8)
                pyautogui.click(x=v[0], y=v[1], clicks=1, interval=0.0, button="left")
                time.sleep(60)

                print_icon_path = os.path.join(self.wd, "images", "print_icon.png")
                urls_title = str(title).replace("|", "").replace("-","").replace(" ","_").replace(",","").replace(":","") 
                try:
                    p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                    with open(default_preview_files_path, "a") as f:
                        f.write(f"{urls_title}\n")

                except Exception as e:
                    print("Preview not loaded: ", e)
                    preview_not_loaded_count += 1
                    with open(default_preview_files_path, "a") as f:
                        f.write("preview not loaded\n")
                    continue
                
                pyautogui.click(x=p[0], y=p[0], clicks=1, interval=0.0, button="right")
                time.sleep(2)
                default_pp_save_icon_path = os.path.join(self.wd, "images", "default_pp_save_icon.png")
                d = pyautogui.locateOnScreen(default_pp_save_icon_path, confidence=0.8)
                pyautogui.click(x=d[0], y=d[1], clicks=1, interval=0.0, button="left")
                time.sleep(2)
                default_pp_pdf_name = urls_title+"_good.pdf"
                default_pp_pdf_path = os.path.join(default_preview_folder, default_pp_pdf_name)
                try:
                    pyautogui.typewrite(default_pp_pdf_path)
                except Exception as e:
                    print("Error: ", e)
                pyautogui.press("enter")
                time.sleep(7)
            
            except Exception as e:
                print("Error: ", e)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(3)
        print(f"Total number of urls: {len(urls)}")
        print(f"Number of previews not loaded: {preview_not_loaded_count}")


    def store_actual_preview_page(self):
        if close_browsers_if_running(browsers_to_check):
            print("One or more browsers were running and have been closed.")
        else:
            print("No specified browsers were running.")


        urls_path = os.path.join(self.input_folder_path, "url_lists.txt")

        with open(urls_path, "r") as urlsFile:
            urls = urlsFile.readlines()

        #create a folder to store the default preview pdfs
        default_preview_folder = os.path.join(self.output_folder_path, "actual_previews")
        if os.path.exists(default_preview_folder):
            #delete the folder and recreate it
            shutil.rmtree(default_preview_folder)
        
        os.makedirs(default_preview_folder)
        preview_not_loaded_count = 0

        #create files to store the file_name
        default_preview_files_path = os.path.join(self.output_folder_path, "urls_title_names_actual.txt")
        if os.path.exists(default_preview_files_path):
            os.remove(default_preview_files_path)


        for url in urls:
            try:
                driver=load_chrome_options(self.chrome_location)
                driver.set_page_load_timeout(60)
                driver.maximize_window()
                #set chrome size as 75%
                # driver.set_window_size(1024, 768)
                driver.get(url.strip())
                title = driver.title

                extension_icon_path = os.path.join(self.wd, "images", "extensionicon.png")
                v = pyautogui.locateOnScreen(extension_icon_path, confidence=0.8)
                pyautogui.click(x=v[0], y=v[1], clicks=1, interval=0.0, button="left")
                time.sleep(60)

                print_icon_path = os.path.join(self.wd, "images", "print_icon.png")
                urls_title = str(title).replace("|", "").replace("-","").replace(" ","_").replace(",","").replace(":","") 
                try:
                    p = pyautogui.locateOnScreen(print_icon_path, confidence=0.8)
                    with open(default_preview_files_path, "a") as f:
                        f.write(f"{urls_title}\n")

                except Exception as e:
                    print("Preview not loaded: ", e)
                    preview_not_loaded_count += 1
                    with open(default_preview_files_path, "a") as f:
                        f.write("preview not loaded\n")
                    continue
                
                pyautogui.click(x=p[0], y=p[0], clicks=1, interval=0.0, button="right")
                time.sleep(2)
                default_pp_save_icon_path = os.path.join(self.wd, "images", "default_pp_save_icon.png")
                d = pyautogui.locateOnScreen(default_pp_save_icon_path, confidence=0.8)
                pyautogui.click(x=d[0], y=d[1], clicks=1, interval=0.0, button="left")
                time.sleep(2)
                default_pp_pdf_name = urls_title+"_actual.pdf"
                default_pp_pdf_path = os.path.join(default_preview_folder, default_pp_pdf_name)
                try:
                    pyautogui.typewrite(default_pp_pdf_path)
                except Exception as e:
                    print("Error: ", e)
                pyautogui.press("enter")
                time.sleep(7)
            
            except Exception as e:
                print("Error: ", e)
            except TimeoutException as e:
                print("Timeout")
            finally:
                driver.quit()
                time.sleep(3)
        print(f"Total number of urls: {len(urls)}")
        print(f"Number of previews not loaded: {preview_not_loaded_count}")

    
    