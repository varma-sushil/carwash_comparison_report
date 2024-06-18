import requests
from dotenv import load_dotenv
import os

load_dotenv()


class tgConfig:
    env_vars    = os.environ
    BOT_TOKEN   = env_vars.get("BOT_TOKEN")
    CHAT_ID     = env_vars.get("CHAT_ID")

class telegramBot():
    def __init__(self) -> None:
        self.BOT_TOKEN = tgConfig.BOT_TOKEN
        self.CHAT_ID   = tgConfig.CHAT_ID
    
    def formt_with_msg_info(self,message:dict):
        "wnat to say something like error message "
        location_name = message.get("location")
        msg_info = message.get("msg")
        message_str=f"""{location_name} {msg_info}"""
        
        return message_str
    
    def update_messgage_format(self,message:dict):
        "updatin message info"
        location_name = message.get("location")
        new_value = message.get("new_value")
        diff      = message.get("diff")
        
        message_str = f"""{location_name}  current : {new_value} change : {diff}"""
        
        return message_str
    
    def single_message(self,msg_list):
        
        return "\n".join([single_msg for single_msg in msg_list])
    
    def send_api(self,message_str)->bool:
        status=False
        url = f'https://api.telegram.org/bot{self.BOT_TOKEN}/sendMessage'
        payload = {
            'chat_id': self.CHAT_ID,
            'text': message_str
        }
        try:
            response = requests.post(url, data=payload)
            if response.status_code == 200:
                print('Message sent successfully')
                status=True
            else:
                print('Failed to send message:', response.text)
        except Exception as e:
            print("Exception while sending with telegram api {e}")
        
        return status
    
    def send_message(self,msg_lst:list):
        "This function will send message"
        status =False
        # is_msg = message.get("msg")
        
        # if is_msg:
        #     message_str = self.formt_with_msg_info(message)
        # else:
        #     message_str= self.update_messgage_format(message)
        
        message_str = self.single_message(msg_lst)

        return self.send_api(message_str)
        
        