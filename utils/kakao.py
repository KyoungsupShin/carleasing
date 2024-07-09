from PyKakao import Message
import json

class kakaomsg():
    def __init__(self):
        TOKEN_FILE = "../data/tokens.json" 
        with open(TOKEN_FILE, 'r') as f:
            data = json.load(f)
            access_token = data.get('access_token')

        self.API = Message(service_key = "7205d2e339b1e499f44ca77518252fb7")
        self.API.set_access_token(access_token)

    def send_pdf_url(self, googleurl):
        link = {
            "web_url": googleurl,
            "mobile_web_url": googleurl,
        }
        button_title = "바로 확인" # 버튼 타이틀
        self.API.send_message_to_friend(
            message_type="text", 
            text=googleurl,
            link=link,
            button_title = button_title,
            receiver_uuids=["THREdUZ1RHxNYVNgVGNUYVZhUn5Pfk5_SXBHIA"]
        )

    def send_info(self, content_data):
        content = {
                    "title": " ",
                    "link": {},
                }
        item_content = {
                    "title_image_text" :"리스견적서 요청 건",
                    "items" : content_data,
                    "sum" :"Total",
                }
        print(content)
        print(item_content)

        self.API.send_message_to_friend(
            message_type="feed",
            content=content, 
            item_content=item_content, 
            receiver_uuids=["THREdUZ1RHxNYVNgVGNUYVZhUn5Pfk5_SXBHIA"]
        )