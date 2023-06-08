import os
os.environ["OPENAI_API_KEY"] = "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"
os.environ["GOOGLE_API_KEY"] = "AIzaSyAiu0_7H1LNEXwt4JL_1yUVECVMCBjLF5E"
os.environ["GOOGLE_CSE_ID"] = "641ce55caab6148af"


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
import whisper
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
import tiktoken
from google.cloud import vision
from typing import Sequence
import io
import win32gui, win32con
import ctypes
from ai_tools import *
from auth import *

SPEECH_MODEL = whisper.load_model("base")
CURR_TASK = ""

with open("./privacy/security-key/openai/openAIKey.txt") as openai_key:
    openai.api_key = openai_key.read()

def hide():
    """
        This function simply hides gui window after triggering
    """
    window = ""
    time.sleep(2)
    try:
        window = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(window, win32con.SW_MINIMIZE)
        return window
    except Exception as e:
        print(str(e))

def show(window):
    """
        This function simply shows gui window
    """
    time.sleep(2)
    win32gui.ShowWindow(window, win32con.SW_NORMAL)

class CustomThread(threading.Thread):
    def __init__(self, target, args=None):
        threading.Thread.__init__(self)
        self.args = args
        self.target = target

    def run(self):
        try:
            self.target(*self.args)
        except Exception as e:
            print(e)
            print("Stopped !")
        print("Exited !")
              
    def get_id(self):
        if hasattr(self,"_thread_id"):
            return self._thread_id
        for id, thread in threading._active.items():
            if thread is self:
                return id

    def stop_exec(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(SystemExit))
        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)
            print("Exception raise failure")

class MemoryManager:
    def __init__(self):
        self.gm = GeneralManager()
        self.memory = self.gm.agent.memory
    
    def _call(self, query: str, insert):
        try:
            self.optimize_space(query)
            self.gm.ask_gui(query, insert)
        except Exception as e:
            print(e)
            insert("e", ".")
    
    def optimize_space(self, query: str):
        """This method will optimize space i.e. keep tokens count less than 1000"""
        count = len(tiktoken.encoding_for_model("gpt-3.5-turbo").encode(query))
        if count>1000:
            raise Exception("Token out of limit !")
            
    def clear_memo(self):
        self.memory.chat_memory.messages = []

class GeneralManager:
    def __init__(self):
        self.rObj = sr.Recognizer()
        self.mic = sr.Microphone()
        with self.mic as src:
            self.rObj.adjust_for_ambient_noise(src, 1)
        engine = pyttsx3.init("sapi5")
        engine.setProperty("voice", engine.getProperty("voices")[1].id)
        self.engine = engine
        self.tools = [ WriteFormula(), sheet_formula(), sheet_input(), WordWriter(),
        SheetValue(), GoogleWrapper(), CreateNewSheet(), highlight_cell(), #EmailWriter()
        EmailAccess(), MusicIdentify(), chat_gpt()]+load_tools(["wikipedia", "llm-math"], llm = ChatOpenAI())
        self.llm = ChatOpenAI(temperature=0.4)
        self.agent = initialize_agent(tools=self.tools, llm = self.llm, agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION, verbose = True, memory=ConversationBufferMemory(memory_key="chat_history"))
        
    def speak(self, data="ABCD"):

        self.engine.say(data)
        self.engine.runAndWait()
    
    def ask_gui(self, query: str, insert):
        ans = self.agent.run(query)
        insert("dhira", ans)
        self.speak(ans)

    
class Graphics:
    def __init__(self, memo_manager):
        self.memo = memo_manager
        self.dhira = self.memo.gm
        self.row_count = 0
        self.window = ""
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("blue")
        self.app = customtkinter.CTk()
        self.app.title("Dhira")
        self.app.geometry("500X400")
        self.app.resizable(False, False)
        self.app.minsize(400, 400)
        self.frame = customtkinter.CTkFrame(self.app)
        self.frame.pack(side=customtkinter.BOTTOM)

        self.text_box = customtkinter.CTkTextbox(self.frame, width=380, height=50)

        self.search_im = customtkinter.CTkImage(Image.open("./resources/images//search.png"))
        self.search = customtkinter.CTkButton(self.frame, image=self.search_im,
                                              text="", height=50,
                                              width=50, border_width=1,
                                              border_color="black", command=self.search_callback)
        self.mic_im = customtkinter.CTkImage(Image.open("./resources/images//mic.png"))
        self.speak_btn = customtkinter.CTkButton(self.frame, image=self.mic_im,
                                                 text="", height=50,
                                                 width=50, border_width=1,
                                                 border_color="black", command=self.ask_callback)

        self.text_box.grid(row=0, column=0)
        self.search.grid(row=0, column=1)
        self.speak_btn.grid(row=0, column=2)

        self.chat_frame = customtkinter.CTkScrollableFrame(self.app, width=480, height=290)
        self.chat_frame.pack(side=customtkinter.BOTTOM, pady=20)

        self.insert_dialog("dhira", "Hi, How can I help you ?")

    def add_row(self):
        self.row_count += 1
        return self.row_count

    def search_callback(self):
        global CURR_TASK
        query = self.text_box.get(1.0, "end")
        query = query.strip()
        if query == "\n" or query == "": return ""
        if query == "stop":
            CURR_TASK.stop_exec()
            return ""
        self.insert_dialog("user", query)
        self.text_box.delete(1.0, "end")

        #CURR_TASK = threading.Thread(target = self.memo._call, args=(query, self.insert_dialog))
        CURR_TASK = CustomThread(target = self.memo._call, args=(query, self.insert_dialog))
        CURR_TASK.start()

    def ask_with_mic(self, rObj, mic, _call, insert_dialog):
        global CURR_TASK
        global SPEECH_MODEL
        rec = rObj
        with mic as src:
            data = rec.listen(src)
            with open("test.mp3", "wb") as f:
                f.write(data.get_wav_data())
            res = SPEECH_MODEL.transcribe("test.mp3")
            query = res["text"]
       
        query.strip()
        if query == "\n" or query == "": return ""
        insert_dialog("user", query)
        #CURR_TASK = threading.Thread(target=self.memo._call, args=(query, self.insert_dialog))
        try:
            _call(query, insert_dialog)
        except Exception as e:
            print(e)

    def ask_callback(self):
        global CURR_TASK
        rec = self.dhira.rObj
        CURR_TASK = CustomThread(target = self.ask_with_mic, args=(self.dhira.rObj, self.dhira.mic, self.memo._call, self.insert_dialog))
        CURR_TASK.start()

    def insert_dialog(self, role, content):
        color = "#002699" if role == "dhira" else "#b32400"
        anchor = "w" if role == "dhira" else "e"
        col = 0 if role == "dhira" else 1
        box = customtkinter.CTkLabel(self.chat_frame, text=content,
                                     width=235, anchor=f"{anchor}",
                                     bg_color=color, wraplength=235,
                                     justify="left")
        box.grid(row=self.add_row(), column=col)

if __name__ == "__main__":
    memo_manager = MemoryManager()
    graphics = Graphics(memo_manager)
    graphics.app.mainloop()

