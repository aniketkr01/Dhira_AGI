from langchain.chains import LLMChain, SimpleSequentialChain
from langchain.llms import OpenAI
from langchain.prompts import PromptTemplate
from langchain import LLMMathChain
from langchain.agents import AgentType, tool, Tool, initialize_agent, load_tools, ZeroShotAgent, AgentExecutor
from langchain.chat_models import ChatOpenAI
from langchain.tools import BaseTool
from langchain.memory import ConversationBufferMemory, ReadOnlySharedMemory
from langchain import OpenAI, LLMChain, PromptTemplate
from langchain.utilities import GoogleSearchAPIWrapper
import wikipedia
import re

import spotipy
import auth
import json
import time
import re
import subprocess
import threading
import win32gui, win32con
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import ast #Abstract Syntax Tree
import pyautogui as auto

import customtkinter
from PIL import Image
import speech_recognition as sr
import pyttsx3
import os
import threading
import playsound
import cv2
import pyautogui as auto
import time
import openai
import tiktoken as tk
from google.cloud import vision
from typing import Sequence
import io
import win32gui, win32con
import ast #Abstract Syntax Tree

def open_browser():
    subprocess.run(["C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe","--remote-debugging-port=9980","--user-data-dir=D:\\Assistant\\chrome"])

class SelectMusic(BaseTool):
    name = "select music"
    description = "useful for when you are asked to play music or song. The input to this tool would be the type of song, singer name or melodious song that you want to hear. Think of artist or genere, For example `kishore kumar song` would be the input if you want to hear kishore song"
    
    def _run(self, query: str) -> str:
        """use the tool"""
        pass
        
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")
    
class MusicIdentify(BaseTool):
    name = "music identify"
    description = "useful for when you are asked to play music or song. The input to this tool would be the type of song, singer name or melodious song that you want to hear. Think of artist or genere, For example `kishore kumar song` would be the input if you want to hear kishore song"
    return_direct = True
    
    
    def _run(self, query: str) -> str:
        """use the tool"""
        tools = [SelectMusic()]
        llm = ChatOpenAI(temperature=0.4)
        music = initialize_agent(tools, llm, agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION).run(query)
        play_music(music)
        return "Task completed !"
        
    async def _arun(self, query: str)-> str:
        raise NotImplementedError("BingSearchRun does not support async")

        
def play_music(query: str):
    keys = auth.keys["spotify"]
    oauth = spotipy.SpotifyOAuth(client_id=keys["clientID"],client_secret=keys["clientSecret"],redirect_uri=keys["redirect_uri"])
    token_dict = oauth.get_cached_token()
    token = token_dict["access_token"]
    spotifyObject = spotipy.Spotify(auth=token)
    res = spotifyObject.search(query, 1, 0, "track")
    song = res["tracks"]["items"][0]["external_urls"]["spotify"]
    threading.Thread(target=open_browser).start()
    opt = Options()
    opt.add_experimental_option("debuggerAddress","localhost:9980")
    service = Service("D://chromedriver")
    driver = webdriver.Chrome(service=service,options=opt)
    win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE)
    driver.maximize_window()
    time.sleep(4)
    driver.execute_script("window.focus();")
    driver.get(song)
    unupdate = True
    while(unupdate):
        try:
            pause = driver.find_elements(By.CSS_SELECTOR,"div[class='os-content'] div[class='os-padding'] div[data-testid|='action-bar-row'] button[data-testid|='play-button']")[0]
            unupdate = False
        except Exception as e:
            pass
    chain = ActionChains(driver)
    chain.click(on_element=pause)
    time.sleep(4)
    chain.perform()

class EmailAccess(BaseTool):
    name = "email access"
    #description = "useful when your email content are ready to send and email service provider. Strictly follow input format to this tool as ['recipient=abcd@gmail.com' ,'subject=leave for two days' ,'body=Dear sir,\nI hope this email finds you well. Due to fever I will be absent for two days.\n Thanking You,\nYours faithfully\n aman']. Your Task completed."
    description = f"""
                    Think step by step:
                    1. useful for when you want to send an email to a person otherwise don't use this tool.
                    2. useful for when you want to use email service.
                    3. You should write content of email effectively.
                    4. Sender details are as follows: {str(auth.keys["personal"]).replace("{","").replace("}","")}
                    5. Input to this tool would be python list: [recipient, subject, body] for example: 
                       input would be `["abcd@gmail.com" ,"leave for two days" ,"Dear sir,\nI hope this email finds you well. Due to fever I will be absent for two days. Be assure that all the due work wil be completed by me in time.\n Thanking You,\nYours faithfully\n aman"]` if sender name aman want leave for two days.
                       input would be `["raghav@gmail.com" ,"congratulation for getting job" ,"Dear Raghav,\nI hope you are doing weel. I am glad to hear that you are selected as bank manager of SBI bank. Keep on progressing and may you success a lot in your carrer.\n Thanking You,\nYours faithfully\n rohit"]` if sender name rohit want to cngratulate raghav.
                    """

    #return_direct = True
    
    def _run(self, query: str) -> str:
        """use the tool"""
        
        print("query is : "+query)
        
        query = query.replace("`","")
        query = query.replace("\"","\"\"\"")
        processed_data = ast.literal_eval(query)
        
        load_ques = re.findall("\[[A-Za-z ]*\]",processed_data[2])
        
        if(len(load_ques)>0): return "please fill all details including sender and recipient inside body. Think again wisely !"
        
        processed_data[2] = processed_data[2].replace("\\n","\n")
        try:
            self.open_email(processed_data[0], processed_data[1], processed_data[2])
            return "successfully send, Task Completed."
        except Exception as e:
            print(e)
            return "Failed ! to send. Check the input format and body format. There might be EOL error"
    
    def get_elmnt(self, driver, css_sel):
        found = False
        et = None
        while(not found):
            try:
                et = driver.find_elements(By.CSS_SELECTOR, css_sel)
                found = True
            except Exception as e:
                pass
        return et[0]
        
    def open_email(self, recipient, subject, body):
        print(recipient)
        threading.Thread(target=open_browser).start()
        opt = Options()
        opt.add_experimental_option("debuggerAddress","localhost:9980")
        service = Service("D://chromedriver")
        driver = webdriver.Chrome(service=service,options=opt)
        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE)
        driver.maximize_window()
        driver.get("https://mail.google.com/mail")
        found = False
        compose = self.get_elmnt(driver, "div.nH div.T-I.T-I-KE.L3") #div.nH.aqk.aql.bkL div.aeN.WR.nH.oy8Mbf div.aic 
        print(compose)
        compose.click()
        to = self.get_elmnt(driver, "input[role='combobox']")  #div[aria-label='To'] div.aGb div.afx div.aH9 
        print(to)
        to.send_keys(recipient)
        sub = self.get_elmnt(driver, "form input[name='subjectbox']") #div.aoD.az6 
        sub.send_keys(subject)
        content = self.get_elmnt(driver, "tr td.Ap div.Ar.Au div.Am.Al.editable.LW-avf.tS-tW")
        content.send_keys(body)
            
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingeSearchRun does not support async")

class EmailWriter(BaseTool):
    name="email writer"
    #description="useful when you want to write an email, you must use sender detail as follows "+str(auth.keys["personal"]).replace("{","").replace("}","")+". You must return whole email as the Final Answer."
    description = f"""
    useful when you want to write an email
    Think step by step:
        1. useful when you want to write an email
        2. Don't use this tool if user does not want to write or send an email
        3. Is user want to write an email ?
        4. You must use sender detail as follows: {str(auth.keys["personal"]).replace("{","").replace("}","")}
        5. Check sender detail filled ! before sending mail 
    """
    def _run(self, query: str) -> str:
        "use the tool"
        print("query writer: ", query)
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

class GoogleWrapper(BaseTool, GoogleSearchAPIWrapper):
    name = "web search"
    description = """
                    Think step by step:
                    1. Use this tool if you want to search about any data or website in web or internet.
                    2. Use this tool for general search or to answer ambiguous questions.
                    2. Use this tool if you want to use Search Engine like Google
                    3. Use this tool to learn about definition, meaning, working or functioning of anything.
                    4. Use this tool to get a technical help from internet.
                  """
    
    def _run(self, query: str) -> str:
        """use the tool"""
        return GoogleSearchAPIWrapper().run(query)
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1RpZAdm060IUR2tBWwmK4auHNZppI1KX4iFhVdigLssc"

def authorize():
    creds = None
    if(os.path.exists('token.json')):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as cred_file:
            cred_file.write(creds.to_json())
    return creds

class SheetValue(BaseTool):
    name = "sheet value"
    description = """
                    Think step by step:
                    1. Use this tool when you want to print values of sheets.
                    2. Use this tool when you want to read values from sheets.
                    3. Use this tool when you want to check formula.
                    3. Input to this tool would be cell range.
                    for example:
                        Input to this tool would be `A1:A2` when you want to read/print values from A1 to A2
                  """
    def sheet_values(self, sheetId: str = SPREADSHEET_ID, cell_range: str = 'A1:A2', sheet_num: str = 'Sheet1'):
        creds = authorize()
        try:
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets() #sheet is representing access to spreadsheets service
            result = sheet.values().get(spreadsheetId=sheetId, range=f"{sheet_num}!{cell_range}").execute()
            values = result.get('values', [])
        
            for i in values:
                print(i)
            return "Done !. Now Stop !"
        except HttpError as e:
            print(e)
            return "Error !"
    
    def _run(self, query: str) -> str:
        """use the tool"""
        print("query: ", query)
        query = query.replace("`","")
        query = query.replace("'","")
        try:
            self.sheet_values(cell_range=query)
            return "Done !"
        except:
            return None
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")
    
class CreateNewSheet(BaseTool):
    name = "create new sheet"
    description = """
                    useful when you want to create new sheets or google spreadsheets. 
                    Input to this tool would be appropriate title, 
                    
                    for example: `Budget` would be the input to this tool if you are creating a budget.
                 """
    
    def create_sheet(self, title: str):
        creds = authorize()
        try:
            service = build('sheets', 'v4', credentials=creds)
            spreadsheet = {
                'properties':{
                    'title': title
                }
            }
            spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
            print(f"spreadsheetId is {spreadsheet.get('spreadsheetId')}")
            return spreadsheet.get('spreadsheetId')
        except HttpError as e:
            print(e)
            
    def _run(self, query: str) -> str:
        global SPREADSHEET_ID
        query = query.replace("`","")     #Need to handle unpredictable behavior of llm
        SPREADSHEET_ID = self.create_sheet(query)
        threading.Thread(target=open_browser).start()
        service = Service("D:\\chromedriver")
        opt = Options()
        opt.add_experimental_option("debuggerAddress","localhost:9980")
        win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_MINIMIZE)
        driver = webdriver.Chrome(service=service, options=opt)
        driver.maximize_window()
        driver.get(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid=0")
        return "new sheet created !"
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

class sheet_input(BaseTool):
    name = "sheet update"
    description = """
                    Think step by step:
                    1. Use this tool only to input or insert data in spreadsheets.
                    2. Use this tool to write formula on multiple cells.
                    3. Input to this tool would be range and value.
                    4. Use maths to ensure that count of value matches range.
                    
                    for example:
                        Input to this tool would be `['A1:D2', [['Name', 'Class', 'Skills','Age'], ['ABCD', 'TY', 'C#',20],]]` when you want
                        to insert data [['Name', 'Class', 'Skills','Age'], ['ABCD', 'TY', 'C#',20],] from A1 to D2
                  """
    def sheet_update(self, sheet_id: str, ranges: str, value: list[list] =  [["Name", "Class", "Skills","Age"], ["ABCD", "TY", "C#",20],], value_input_option: str = "USER_ENTERED"):
        """
            value_input_option: RAW | USER_ENTERED
            ranges: A1:C5,D1:D4...
        """
        creds = authorize()
        try:
            
            values = [value]
            ranges = ranges.split(",")
            
            if(ranges[0][0]==ranges[0][3]):  #Handling case if llm unable to do maths properly
                values = []
                for l in value:
                    for k in l:
                        values.append([k])
                values = [values]
                
            data = [
                {
                    "values": values[0],
                    "range": ranges[0],
                },
            ]
        
            body = {
                "valueInputOption" : value_input_option,
                "data": data,
            }
        
            service = build("sheets", "v4", credentials=creds)
            exec("result = service.spreadsheets().values().batchUpdate(spreadsheetId = sheet_id, body = body).execute()")
            #print(f"{result.get('totalUpdatedCells')}") #count of updated cells
            #return f"{result.get('totalUpdatedCells')}"
        except HttpError as e:
            print(e)
            raise e
            
    def _run(self, query: str) -> str:
        #lst = []
        query = query.replace("`","")
        try:
            #exec("lst="+query+"\nself.sheet_update(SPREADSHEET_ID, lst[0], lst[1])")
            lst = ast.literal_eval(query)
            self.sheet_update(SPREADSHEET_ID, lst[0], lst[1])
            return "inserted data successfully !. Task Completed !"
        except Exception as e:
            print(e)
            return "failed to add data !. Make sure not to include ellipsis. Follow format !"
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

class WriteFormula(BaseTool):
    name = "apply formula"
    description = """
                    Think step by step:
                    1. Use sheet formula tool before using this tool else don't use it first.
                    2. Use this tool when you have a formula ready with you else don't use.
                    3. Use this tool only when you want to apply formula on a cell else don't use.
                    4. Input to this tool would be 'range_name' and 'value' as list.
                    5. Use maths to ensure count of value matches range.
                    
                    for example:
                        Input to this tool would be `['C1:C', '=A1:A-B1:B']` when you have a formula '=A1:A-B1:B' and storing result in column C starting from C1.
                        Input to this tool would be `[['C1:C', '=A1:A+B1:B'], ['D1:D', '=A1:A-B1:B']]` when you want to first add column A and B, store result in C then subtract column A and B, store result in D. 
                  """
    def apply_sheet_formula(self, sheet_id: str, range_name: str, value: str):
        creds = authorize()
        try:
            service = build("sheets", "v4", credentials=creds)
            values = [
                [
                  value  
                ],
            ]
        
            body = {
                "values": values,
            }
            result = service.spreadsheets().values().update(spreadsheetId=sheet_id,range=range_name,
                                                        body=body,valueInputOption="USER_ENTERED").execute()
            print(result)
        except HttpError as e:
            print(e)
            raise e
    
    def _run(self, query: str) -> str:
        print("Query is : ", query)
        #lst = []
        query = query.replace("`","")
        try:
            #exec("lst="+query+"\nself.apply_sheet_formula(SPREADSHEET_ID, lst[0], lst[1])")
            lst = ast.literal_eval(query)
            for formula in lst[1:]:
                if type(formula)==list:
                    self.apply_sheet_formula(SPREADSHEET_ID, formula[0], formula[1])
                else:
                    self.apply_sheet_formula(SPREADSHEET_ID, lst[0], formula)
            return "successfully completed task !"
        except Exception as e:
            print(e)
            return "failed to write data !. Make sure not to use ellipsis. Follow format !"
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

class sheet_formula(BaseTool):
    name = "find formula"
    description = """
                    Think step by step:
                    1. Use this tool to know sheet formula and functions and generate data else don't use.
                    2. Use this tool to write google spreadsheets formula
                    3. Input to this tool will be specific formula that you want
                    4. Output of this tool contains only formula
                    
                    for Example:
                        1.for input 'calculate profit from A and B' output will be '=A1:A-B1:B'
                        
                  """
    def _run(self, query: str) -> str:
        pass
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")
class highlight_cell(BaseTool):
    name = "highlight cell"
    description = """Think step by step:
                     1. Use this tool if you want to highlight or color cells in the sheets else do not use.
                     2. Input to this tool would be [condition-type, threshold, color, range]
                     3. You must use one of the following colors: red, green, blue
                     
                     Rules for deciding condition-type:
                     `NUMBER_GREATER` : when the cell's value must be greater than the condition's value.
                     `NUMBER_GREATER_THAN_EQ` : when the cell's value must be greater than or equal to the condition's value.
                     `NUMBER_LESS` : when the cell's value must be less than the condition's value.
                     `NUMBER_LESS_THAN_EQ`: when the cell's value must be less than or equal to the condition's value.
                     `NUMBER_EQ` : when the cell's value must be equal to the condition's value.
                     `NUMBER_NOT_EQ` : when the cell's value must be not equal to the condition's value.
                     `NUMBER_BETWEEN` : when the cell's value must be between the two condition values.
                     `TEXT_CONTAINS` : when the cell's value must contain the condition's value.
                     `TEXT_IS_EMAIL` : when the cell's value must be a valid email address.
                     `TEXT_IS_URL` : when the cell's value must be a valid URL.
                     
                     for example: 
                     `['NUMBER_GREATER_THAN_EQ', '90', 'green', 'A1:A5']` would be the input if you want to highlight cells with green color from A1 to A5 having number greater than or equal to 90.
                     `['NUMBER_LESS', '80', 'red', 'A1:B5']` would be the input if you want to highlight cells with red color from A1 to B5 having number less than 80.
                     `['TEXT_IS_EMAIL', '', 'green', 'A4:A8']` would be the input if you want to highlight cells with green color from A4 to A8 having valid email or mail.
                  """
    
    def numeric(self, lst):
        while '' in lst:
            lst.remove('')
        if len(lst)==0: return 1000
        return int(lst[0]) - 1

    def alpha(self, lst):
        c = 0
        for i in range(len(lst)):
            c = c + (ord(lst[i])-64)*pow(26,len(lst)-1-i)
            
        return c-1
    
    def range_cal(self, range_name: str):
        try:
            myRange = {
                "sheetId": 0,
                "startRowIndex":0,
                "endRowIndex": 100,
                "startColumnIndex": 0,
                "endColumnIndex": 100
            }
        
            lst = range_name.split(":")
            
            rs = re.findall("[0-9]*", lst[0])
            cs = re.findall("[A-Z]", lst[0])
            r_end = re.findall("[0-9]*", lst[1])
            ce = re.findall("[A-Z]", lst[1])
            
            myRange["startColumnIndex"] = self.alpha(cs)
            myRange["endColumnIndex"] = self.alpha(ce) + 1
            myRange["startRowIndex"] = self.numeric(rs)
            myRange["endRowIndex"] = self.numeric(r_end)
            
            if myRange["startRowIndex"]==1000: myRange["startRowIndex"] = 0
            
            return myRange
        except Exception as e:
            print(e)
            return "error"
    
   
    def highlight(self, rule: str, qty: str, color: str, range_name: str):
        creds = authorize()
        try:
            my_range = self.range_cal(range_name)
            if my_range=='error':
                return "failed !.Invalid format of range."
            
            requests = [
                {
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [my_range],
                            "booleanRule": {
                                "condition": {
                                    "type": rule,
                                    "values": [
                                        {
                                            "userEnteredValue": qty
                                        }
                                    ]
                                },
                                "format": {
                                    "backgroundColor": {
                                        color : 0.8
                                    }
                                }
                            }
                        },
                        "index": 0
                    }
                }
            ]
            
            body = {
                "requests": requests,
            }
            service = build("sheets", "v4", credentials=creds)
            result = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
            print("Response is ", result)
            print(f"response: {result.get('replies')}")
        except HttpError as e:
            print(e)
            raise e
    
    def _run(self, query: str) -> str:
        query = query.replace("`","")
        lst = ast.literal_eval(query)
        try:
            self.highlight(lst[0], lst[1], lst[2], lst[3])
            return "Task completed successfully."
        except Exception as e:
            print(e)
            return "Failed !.Check input format."
        
    async def _arun(self, query: str) -> str:
        raise NotImplemetedError("BingSearchRun does not support async")

class user_input(BaseTool):
    name = "user input"
    description = "Useful when you want to ask user for clarification, you must use this tool rarely when you cannot decide on your own."
    
    def _run(self, query: str) -> str:
        """use the tool"""
        return "user input is: " + input(query)
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

class chat_gpt(BaseTool):
    name = "general question"
    description = """useful for when you want to greet or give answer to general question.
    for example:
    `Hello` would be the input to this tool if you want the answer of `Hello how are you` or any greetings.
    """
    #return_direct = True
    
    def _run(self, query: str) -> str:
        llm = ChatOpenAI()
        return llm.call_as_llm(query)
    
    async def _arun(self, query: str) -> str:
        NotImplementedError("BingSearchRun does not support async")

# class Summarize(BaseTool):
#     name = "Summary"
#     description = "useful for when you summarize a conversation or understand the context of user. The input to this tool should be a string, representing who will read this summary."
    
#     def _run(self, query: str) -> str:
#         """use the tool"""
#         llm = OpenAI(temperature=0.4)
#         return llm("summarize the conversation below:\n\n{chat_history}".format(**agent.memory.load_memory_variables({})))
    
#     async def _arun(self, query: str) -> str:
#         raise NotImplementedError("BingSearchRun does not support async")

class WordWriter(BaseTool):
    name = "word writer"
    description = """useful when you want to use microsoft word or ms word in order to write a content. Input to this tool would be content topic that you want to write.
    for example: Input would be `Population` when you want to write about population.
    """
    
    def _run(self, query: str) -> str:
        """use the tool"""
        llm = OpenAI(temperature=0.4)
        content = llm("You are expert in writing article in word, Write an article on {topic} and summarize within 1000 words".format(topic=query))
        auto.hotkey("win")
        time.sleep(2)
        auto.write("word", interval=0.2)
        auto.press("enter")
        time.sleep(5)
        auto.press("enter")
        auto.write(content)
        auto.hotkey("ctrl","s")
        time.sleep(2)
        bx, by = auto.locateCenterOnScreen("D:\\Assistant\\resources\\images\\save.png")
        auto.click(bx, by)
        return "Success!. Task Completed !. Save and Close"
    
    async def _arun(self, query: str) -> str:
        raise NotImplementedError("BingSearchRun does not support async")

