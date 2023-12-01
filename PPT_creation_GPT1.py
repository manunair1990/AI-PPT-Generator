import streamlit as st
import openai
import os
#import json
from dotenv import load_dotenv
from pptx import Presentation	 
#from pptx.util import Inches, Pt 
# Creating presentation object 
ppt = Presentation() 
first_slide_layout = ppt.slide_layouts[0] 

from docx import Document
document = Document()

# Load environment variables
load_dotenv()

# Configure OpenAI API
openai.api_key = os.getenv('OPENAI_API_KEY')

if 'slides_info' not in st.session_state:
    st.session_state.slides_info = ''

# Chat UI title
st.header("PowerPoint Presentations Generator")

Editor_prompt = """You are an efficient Content Editor and Content Creator with 20 years of experience. 
You will be provided with Power point presentation content and the instruction from the user. 
You need to edit the content and make changes based on the instruction given by the user. 
Return the edited content.
"""

Code_checker_prompt = """You are an efficient Senior python programmer with 30 years of experience. 
You are a code checker who checks and test the code, and edit the code if needed.
You need to ensure that, the result must be a python code only, without any extra texts. 
Result must be a perfect optimised python code. 
"""
    
ppt_generator_prompt = """ You are an efficient Creative Python programmer with 20 year of experience.
Your job is the write a creative python code to create a power point presentation with the given information. 
The slides must contain catchy colors, templates, shapes and everything to make the slide interesting. 
The result must be a pure python code that can be saved to a file. 
If there is text lines other than code, it must be commented with #
"""    

slides_generator_prompt = """You are an efficient Business Analyst with 20 years of experience.
    Your job is to create a compelling texts for power point slides for the topic given by the user.
    Do not exceed the number of slides more than 10.  
    """

def generator(system_prompt,user_prompt):
    
    completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    temperature = 0.5,
    max_tokens=2000,
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt}
    ]
    )
    
    return (completion.choices[0].message['content'])
        
with st.sidebar:
    topic_name=st.text_input('Enter the topic')
    #topic_name = st.text_area("Details for the blog")
    if topic_name:
        st.write(topic_name)
    analyze = st.button("Generate Content")
    #Edit = st.text_input('Edit')
    #Edit_button = st.button("Document Edit")
    ppt_generate = st.button("Generate the PPT")
    
    #save = st.button("Save the blog")
    
if topic_name and analyze:
    #st.write("Generating the blog")
    
    with st.spinner('Generating the slides'):
        st.session_state.slides_info = generator(slides_generator_prompt,"Create a power point deck on "+topic_name+"\n Answer: ")
    st.write(st.session_state.slides_info)
    
#st.write(slides_info)
if not st.session_state.slides_info == '' and ppt_generate:    
    print("Entered")
    with st.spinner('Generating the code to create ppt'):    
        code_info = generator(ppt_generator_prompt,"Create a python code with this information:\n "+st.session_state.slides_info+"\nAnswer: ")
    with st.spinner('Rechecking the code'):    
        code_perfect = generator(Code_checker_prompt,"Check this python code and return an optimised perfect code.\nPython code: "+code_info+"\nAnswer: ")
    
    with st.spinner('Generating the ppt'):    
        # Writing the code to a Python file
        fname = 'presentation_code.py'
        with open(fname, 'w') as file:
            file.write(code_perfect)
        os.system(f'python {fname}')
        
    st.write("PPT Generated")   
