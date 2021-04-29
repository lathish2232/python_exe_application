
#---------------------------------------------------------------------------------------------------------
#       Author:-Sreenath,Lathish(lathish2232@gmail.com)
#       Date:- 04/22/2021
#       Contact us :- contact@unificater.com
#       Language :-python3.9 
#---------------------------------------------------------------------------------------------------------
import os
import time
import webbrowser

import tkinter as tk
import pandas as pd

from os.path import join
from tkinter import filedialog, messagebox
from pandas.core.indexes.base import Index
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException


root = tk.Tk()

root.geometry("1100x700")
root.title('Unificater Services')
root.pack_propagate(False)
root.iconbitmap(join(os.getcwd(),'logo.ico'))

#root.resizable(0, 0) 

#menu bar
menu_bar = tk.Menu(root)

file_menu=tk.Menu (menu_bar, tearoff=0)
file_menu.add_command(label='Exit', command=root.destroy)
menu_bar.add_cascade (label='File',menu =file_menu)

#variable assignment

url = "https://unificater.com/#/"
url1="http://qa.unificater.com/#/dashboard"
url2= "https://www.gebiz.gov.sg/"
msg='Waited for 30 seconds to navigate to the desired page and closed the browser'
browser_msg="console info:- \n Browser waiting time is only 60 Seconds. please select desired page to process your request ...\n Please Don't close the chrome Browser...."
user_cancle_msg= ' '*10+"Process cancelled as per Your Request"+' '*10

def openweb():
    webbrowser.open(url)
def openweb1():
    webbrowser.open(url1)

about_menu =tk.Menu (menu_bar, tearoff=0)
about_menu.add_command(label='About us', command=openweb)
about_menu.add_command(label='other Tools', command=openweb1)
menu_bar.add_cascade (label='About',menu =about_menu)

root.config(menu=menu_bar)


ouput_path=tk.StringVar()

df=pd.DataFrame()
#------------------------------------------ input File---------------------------------
file_frame = tk.LabelFrame(root, text="select Input File")
file_frame.place(height=100, width=450, rely=0.01, relx=0.03)

label_file = tk.Label(file_frame , text="No File Selected")
label_file.place(rely=0, relx=0)

button = tk.Button(file_frame, text="Browse Input File", command=lambda: input_file_browser())
button.place(rely=0.6, relx=0.4)

#------------------------------------------ Output File---------------------------------
file_frame1 = tk.LabelFrame(root, text="Save output File")
file_frame1.place(height=100, width=450,x=100, rely=0.01, relx=0.37)

label_file1 = tk.Label(file_frame1 , text="No File Selected")
label_file1.place(rely=0, relx=0)

button1 = tk.Button(file_frame1, text="Save output File into", command=lambda: save())
button1.place(rely=0.6, relx=0.4)
#------------------------------------------Buttions ---------------------------------
button2 = tk.Button(root, text="MOM Open Browser",width=25, command=lambda: mom_service(df))
button2.place(rely=0.25, relx=0.20)

button3 = tk.Button(root, text="Gebiz Open Browser",width=25, command=lambda: gebiz_open_browser(df))
button3.place(rely=0.25, relx=0.50)




file_frame3 = tk.LabelFrame(root, text="Console Info")
file_frame3.place(height=299, width=920, rely=0.30, relx=0.03)

txt = tk.Text(file_frame3, height=14,width =100,font=("Bookman Old Style", 11))
txt.place(rely=0.001, relx=0.01)
scroll_bar = tk.Scrollbar(file_frame3,orient="vertical", command=txt.yview)
txt['yscroll']=scroll_bar.set
txt.insert('1.0',"User Info:-\n Please select Input and Output File Paths Bassed on Requirement.... \n")
scroll_bar.pack(side="right", fill="y")




def mom_service(df):
    input_file=label_file["text"]
    if os.path.isfile(input_file):
        chrome_path= join(os.getcwd(),'chromedriver.exe')
        driver = webdriver.Chrome(executable_path=chrome_path)
        df=pd.read_excel(input_file)
        df.columns=['licenseNo', 'EAName','Address']
        for i, licenseNum in enumerate(df.licenseNo):
            df.loc[df['licenseNo']==licenseNum,['Address','Telephone']]=extract_text(licenseNum,driver)
        output_path=label_file1["text"]
        if os.path.isfile(output_path):
            output_path=output_path if output_path.endswith('.csv') else output_path+'.csv'
            df.to_csv(output_path, index=False)
            driver.quit()
            txt.delete('1.0', tk.END)
            txt.insert('1.0', f'User Info:-\n File Downloaded into {output_path}')
            messagebox.showinfo("Info", f"File Downloaded into {output_path}",icon=messagebox.INFO)
        else:
            messagebox.showinfo(title="caution", message='please select output file before proceeding', icon=messagebox.ERROR)
    else:
        messagebox.showinfo(title="caution", message='please select Input file before proceeding', icon=messagebox.ERROR)

def gebiz_open_browser(df):
    output_path=label_file1["text"]
    if os.path.isfile(output_path):
        chrome_path= join(os.getcwd(),'chromedriver.exe')
        driver = webdriver.Chrome(executable_path=chrome_path)
        driver.get(url2)
        time.sleep(0.02)
        txt.delete('1.0', tk.END)
        txt.insert('1.0', browser_msg)
        confirm=messagebox.askokcancel(
                                        title="Information", message=' '*10+"Click Ok to Resume Process"+' '*20, icon=messagebox.INFO)
        if confirm:
            try:
                Last=WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Last']")))
                try:
                    selected=driver.find_element_by_xpath("//div[@class='headerMenu2_MENU-BUTTON-DIV-SELECTED']/span/a")
                    text_element=selected.text
                except:
                    text_element='Opportunities'
                if text_element =='Supplier Directory':
                    print('i am in 123 ')
                    #----------------------------------------SUPPLIER DIRECTORY----------------------------
                    LastId=Last.get_attribute('id')
                    nxt=driver.find_element_by_xpath("//input[@value='Next']")
                    NextId=nxt.get_attribute('id').split('_')
                    for Nxt in range(int(LastId.split('_')[-1])):
                        NextId[-1]=str(Nxt+2)
                        nextId='_'.join(NextId)
                        try:
                            Next=WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, nextId))
                            )
                        except:
                            messagebox.showinfo(title="Abort", message=msg, icon=messagebox.ERROR)
                            driver.quit()
                        labels = driver.find_elements_by_xpath("//div[@class='formRow_HIDDEN-LABEL']/span/a/..")                    
                        lable_list=[l.get_attribute('id') for l in labels]
                        for i,label_lst in enumerate(lable_list):
                            lids = driver.find_elements_by_xpath("//div[@class='formRow_HIDDEN-LABEL']/span/a/..")
                            lid_lst=[l.get_attribute('id') for l in lids]
                            try:
                                link=WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, f'''{lid_lst[i]}'''))
                                )
                            except:
                                messagebox.showinfo(title="Abort", message=f'Error occurred in the page {Nxt+1}', icon=messagebox.ERROR)
                                continue
                            x=driver.find_element_by_xpath(f'''//span[@id="{lid_lst[i]}"]/a''')
                            #print('X: ', x)
                            CompanyName=x.text
                            x.click()
                            try:
                                element=WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@class='form2_MAIN'][position()<=2]//div[@class='formColumns_COLUMN-TABLE']/div[position()!=1]"))
                                )
                            except:
                                messagebox.showinfo(title="Abort", message=msg, icon=messagebox.ERROR)
                                driver.quit()
                            row_labels=driver.find_elements_by_xpath("//div[@class='form2_MAIN'][position()<=2]//div[@class='formColumns_COLUMN-TABLE']/div[position()!=1]//span")
                            row_values=driver.find_elements_by_xpath("//div[@class='form2_MAIN'][position()<=2]//div[@class='formColumns_COLUMN-TABLE']/div[position()!=1]//div[@class='formOutputText_VALUE-DIV ']")

                            keys=[keys.text for keys in row_labels]
                            values=['00' + values.text[1:] if values.text.startswith('+') else values.text for values in row_values]
                            keys.insert(0,'Company Name')
                            values.insert(0, CompanyName)
                            df=df.append(pd.DataFrame([dict(zip(keys, values))]), ignore_index=True)
                            driver.find_element_by_xpath("//input[@value='Back to Search Results']").click()
                    
                        Next=driver.find_element_by_id(nextId)
                        isDisable=Next.get_attribute('disabled')
                        if not isDisable:
                            Next.click()
                    df.to_csv(output_path, index=False)
                    txt.delete('1.0', tk.END)
                    txt.insert('1.0', f'User Info:-\n File Downloaded into {output_path}')
                    messagebox.showinfo("Info", f"File Downloaded into {output_path}",icon=messagebox.INFO)
                    driver.quit()   
                else:
                    print('i am in 666 ')
                    #----------------------------------------Advanced Search----------------------------
                    LastId=Last.get_attribute('id')
                    nxt=driver.find_element_by_xpath("//input[@value='Next']")
                    NextId=nxt.get_attribute('id').split('_')
                    for Nxt in range(int(LastId.split('_')[-1])):
                        NextId[-1]=str(Nxt+2)
                        nextId='_'.join(NextId)
                        try:
                            Next=WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.ID, nextId))
                            )
                        except Exception as ex:
                            messagebox.showinfo(title="Abort", message=msg, icon=messagebox.ERROR)
                            driver.quit()

                        labels = driver.find_elements_by_xpath("//a[@class='commandLink_TITLE-BLUE']")
                        lids1 = driver.find_elements_by_xpath('//span/input[@class="commandButton_MAIN"]')
                        lid_lst1=[l.get_attribute('id') for l in lids1]
                        for i,label_lst in enumerate(lid_lst1):
                            lids = driver.find_elements_by_xpath('//span/input[@class="commandButton_MAIN"]')
                            lid_lst=[l.get_attribute('id') for l in lids]
                            try:
                                link=WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.ID, f'''{lid_lst[i]}'''))
                                )
                            except Exception as ex:
                                messagebox.showinfo(title="Warning", message=f'Error occurred in the page {Nxt+1} for the link text "{label_lst}"', icon=messagebox.WARNING)
                                continue
                            x=driver.find_element_by_xpath(f'''//input[@id="{lid_lst[i]}"]/following-sibling::a''')
                            x.click()
                            try:
                                element=WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//div[@class='col-md-9  formColumn_MAIN formColumns_COLUMN-TD']/div[@class='formColumns_COLUMN-TABLE']//div[@class='row  form2_ROW-TABLE']//div[@class='form2_ROW-LABEL-DIV' or @class='formTooltip_LABEL']//span"))
                                )
                            except Exception as ex:
                                messagebox.showinfo(title="Abort", message=msg, icon=messagebox.ERROR)
                                driver.quit()

                            row_labels=driver.find_elements_by_xpath("//div[@class='col-md-9  formColumn_MAIN formColumns_COLUMN-TD']//div[@class='row  form2_ROW-TABLE']//label/span")
                            row_values=driver.find_elements_by_xpath("//div[@class='col-md-9  formColumn_MAIN formColumns_COLUMN-TD']/div[@class='formColumns_COLUMN-TABLE']//div[@class='row  form2_ROW-TABLE']//div[@class='formOutputText_VALUE-DIV ']")

                            keys=[keys.text for keys in row_labels]
                            values=['00' + values.text[1:] if values.text.startswith('+') else values.text for values in row_values]
                            df=df.append(pd.DataFrame([dict(zip(keys, values))]), ignore_index=True)
                            driver.find_element_by_xpath("//input[@value='Back to Search Results']").click()
                    
                        Next=driver.find_element_by_id(nextId)
                        isDisable=Next.get_attribute('disabled')
                        if not isDisable:
                            Next.click()
                output_path=output_path if output_path.endswith('.csv') else output_path+'.csv'
                df.to_csv(output_path, index=False)
                driver.quit()
                txt.delete('1.0', tk.END)
                txt.insert('1.0', f'User Info:-\n File Downloaded into {output_path}')
                messagebox.showinfo("Info", f"File Downloaded into {output_path}",icon=messagebox.INFO)
            except Exception as  ex:
                print('i am in exception--------')
                messagebox.showinfo(title="Abort", message=ex, icon=messagebox.ERROR)
                driver.quit()
                txt.delete('1.0', tk.END)
                raise ex
    else:
        messagebox.showinfo(title="caution", message='please select output file before proceeding', icon=messagebox.ERROR)
         
    return None

def save():
    output_path=filedialog.asksaveasfile(filetypes =[('csv Files','*.csv')])
    if output_path:
        label_file1["text"] = output_path.name
        txt.delete('1.0', tk.END)
        txt.insert('1.0', 'User Info:-\n please choose Mom service or Gebiz service based on your Requirement ....')
    else:
        label_file1
    return None

def input_file_browser():
    input_path=filedialog.askopenfilename(filetypes =[('Excel Files','*.xlsx')])
    if input_path:
        label_file["text"] = input_path
        txt.delete('1.0', tk.END)
        txt.insert('1.0', 'User Info:-\n please choose Mom service or Gebiz service based on your Requirement ....')
    else:
        label_file1
    return None

def extract_text(license,driver):
    driver.get('https://service2.mom.gov.sg/eadirectory/')
    try:
        element=WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "fwTab"))
        )
    except Exception as ex:
        messagebox.showinfo(title="Abort", message=ex, icon=messagebox.ERROR)
        driver.quit()
    driver.find_element_by_id("fwTab").click()
    driver.find_element_by_id("eadHomeForm:searchEaNameFilter2").send_keys(license)
    SearchButton=driver.find_element_by_id("eadHomeForm:searchEaNameFilterBtn2")

    SearchButton.click()
    try:
        element=WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, "eadAllAgenciesForm:allAgenciesList:0:j_id_18_52"))
        )
    except:
        return "Not Available", "Not Available"

    try:  
        driver.execute_script('arguments[0].click()', element)
    except NoSuchElementException:
        return "Not Available", "Not Available"
    
    try:  
        address=WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "eadAgencyDetailsForm:j_id_18_6k"))
        )
    except NoSuchElementException:
        return "Not Available", "Not Available"
    try:  
        phonenum=driver.find_element_by_xpath("//a[@id='eadAgencyDetailsForm:j_id_18_6k']/following-sibling::p/span")
    except NoSuchElementException:
        return address.text, "Not Available"
    return address.text, phonenum.text




root.mainloop()

