import os
from dotenv import load_dotenv
load_dotenv()
from openai import OpenAI
from typing import List, Dict
from PyPDF2 import PdfReader


class PDFChatbot:
    def __init__(self, api_key: str, base_url: str, pdf_path: str):
       
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.pdf_content = self._extract_pdf_content(pdf_path)
        self.chat_history: List[Dict[str, str]] = []
        self.system_prompt = self._build_system_prompt()
        
    def _extract_pdf_content(self, pdf_path: str) -> str:
     
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                content = ""
                for page in pdf_reader.pages:
                    content += page.extract_text() + "\n"
                return content
        except Exception as e:
            raise Exception(f"Error reading PDF: {str(e)}")
    
    def _build_system_prompt(self) -> str:
     
        return f"""You are an expert assistant specializing in Monenco Company's risks and strategies. Your role is to provide accurate, comprehensive answers based solely on the provided document.

DOCUMENT CONTENT:
{self.pdf_content}

INSTRUCTIONS:
1. Answer user questions completely and accurately based ONLY on the document content above
2. Focus your answers directly on what the user is asking
3. Before providing your final answer, internally verify the information for accuracy
4. Structure your responses clearly and professionally
5. If the document doesn't contain information to answer a question, politely state this
6. At the end of each response, suggest ONE relevant follow-up question that the user might find interesting, where the answer IS available in the document
7. Maintain context from previous messages in the conversation

RESPONSE FORMAT:
[Your comprehensive answer here]

پرسش پیشنهادی برای کاربر: [A relevant question the user might want to ask next]

Remember: 
- Be thorough but concise
- Stay focused on the user's specific question
- Always double-check your answer against the source material
- Only suggest questions that can be answered from the document"""

    def chat(self, user_message: str) -> str:
        self.chat_history.append({
            "role": "user",
            "content": user_message
        })
       
        messages = [
            {"role": "system", "content": self.system_prompt}
        ] + self.chat_history
        
        try:
           
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=messages,
                temperature=0.1,
                max_tokens=1000
            )
            
            assistant_message = response.choices[0].message.content
           
            self.chat_history.append({
                "role": "assistant",
                "content": assistant_message
            })
            
            return assistant_message
            
        except Exception as e:
            return f"Error: {str(e)}"
    
    def clear_history(self):
        self.chat_history = []
        
    def get_history(self) -> List[Dict[str, str]]:
        return self.chat_history


def main():
    API_KEY = os.getenv("OPENAI_API_KEY")
    BASE_URL = os.getenv("OPENAI_API_BASE")
    PDF_PATH = r"C:\Users\ahmadi.mohammadreza\Documents\GitHub\ChatBot_RT\Assets\Risk.pdf"  
    try:
        chatbot = PDFChatbot(
            api_key=API_KEY,
            base_url=BASE_URL,
            pdf_path=PDF_PATH
        )
        print("=" * 60)
        print("Monenco Iran Consulting Engineers Chatbot")
        print("=" * 60)
        print("Assistant: Ask me anything about Monenco Company's Risk Glossary!")
        print("Type 'exit' to exit, 'clear' to clear history")
        print("=" * 60)
        print()
        
        while True:
            user_input = input("").strip()     # input("You: ").strip()
            
            if not user_input:
                continue
                
            if user_input.lower() == 'exit':
                print("Thank you for using Monenco Chatbot!")
                break
                
            if user_input.lower() == 'clear':
                chatbot.clear_history()
                print("Chat history cleared!")
                continue
            
            print("\n", end="")      # print("\nAssistant: ", end="")
            response = chatbot.chat(user_input)
            print(response)
            print("\n" + "-" * 60 + "\n")
            
    except Exception as e:
        print(f"Error initializing chatbot: {str(e)}")


if __name__ == "__main__":
    main()