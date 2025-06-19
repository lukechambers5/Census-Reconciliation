import os
from dotenv import load_dotenv

load_dotenv()

ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
VIEW_ID = os.getenv("VIEW_ID")
TABLEAU_SERVER = os.getenv("TABLEAU_SERVER")
TOKEN_NAME = os.getenv("TOKEN_NAME")
SITE_ID = os.getenv("SITE_ID")
