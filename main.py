#imports/setup
import os
from google import genai
from google.genai import types

# Run the following line in powershell before running the code for the first time
# set GEMINI_API_KEY = "AIzaSyCppkYuwHki_23FLvtg74MS282Y80UBWHU"
# The client automatically picks up the GEMINI_API_KEY environment variable
# No code change needed to use the secure method
try:
    client = genai.Client()
except Exception as e:
    print(f"Error initializing the client: {e}")
    print("Please ensure the GEMINI_API_KEY environment variable is set correctly.")
    exit()


client = genai.Client()

# Send a prompt and get a response
model_name = "gemini-2.5-flash"

# Format to use ai:
#prompt = "Explain how AI works in simple terms."
#
#response = client.models.generate_content(
#    model=model_name,
#    contents=prompt
#)
#
#print(response.text)

#Making Combined File


#Evaluating File

