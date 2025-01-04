from PyKakao import Message
import json
import asyncio
import telegram

token = '6457267364:AAG_cnlHMo5pvQAFzbfZUok05q72tsTwgO8'
chat_id = '-1002163072869' #이한빛

async def send_pdf_url(googleurl):
    bot = telegram.Bot(token)
    async with bot:
        await bot.send_message(text=googleurl, chat_id=chat_id)

async def send_info(content_data):
    bot = telegram.Bot(token)
    async with bot:
        await bot.send_message(text=content_data, chat_id=chat_id)

async def main():
    bot = telegram.Bot(token)
    async with bot:
        updates = await bot.getUpdates()
        for u in updates:
            print(u.message)        
        print(await bot.get_me())

# class kakaomsg():
#     def __init__(self):
#         TOKEN_FILE = "../data/tokens.json" 
#         with open(TOKEN_FILE, 'r') as f:
#             data = json.load(f)
#             access_token = data.get('access_token')

#         self.API = Message(service_key = "7205d2e339b1e499f44ca77518252fb7")
#         self.API.set_access_token(access_token)

#     def send_pdf_url(self, googleurl):
#         self.googleurl = googleurl
#         link = {
#             "web_url": self.googleurl,
#             "mobile_web_url": self.googleurl,
#         }
#         button_title = "바로 확인" # 버튼 타이틀
#         self.API.send_message_to_friend(
#             message_type="text", 
#             text=self.googleurl,
#             link=link,
#             button_title = button_title,
#             receiver_uuids=["TH1LekN7SXFAdlppW21YYVZkXWpGd0Z2R3FIfw0"]
#             # receiver_uuids= ["THlLfE17T3dCblxvXGRXZ1JqWXVEdUV0QntMPQ"]
#         )

#     def send_info(self, content_data):
#         content = {
#                     "title": " ",
#                     "link": {},
#                 }
#         item_content = {
#                     "title_image_text" :"리스견적서 요청 건",
#                     "items" : content_data,
#                     "sum" :"Total",
#                 }
#         print(content)
#         print(item_content)
        
#         self.API.send_message_to_me(
#             message_type="feed",
#             content=content, 
#             item_content=item_content, 
#             receiver_uuids=["TH1LekN7SXFAdlppW21YYVZkXWpGd0Z2R3FIfw0"]
#             # receiver_uuids= ["THlLfE17T3dCblxvXGRXZ1JqWXVEdUV0QntMPQ"]
#         )
# if __name__ == '__main__':
    # km = kakaomsg()
    # km.send_info(input_data['contents'])
if __name__ == '__main__':
    asyncio.run(main())
    # bot = telegram.Bot(token=token)
    # updates = bot.getUpdates()
    # for u in updates:
    #     print(u.message)