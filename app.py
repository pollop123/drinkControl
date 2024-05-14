import random
from readline import insert_text
from this import s
from flask import Flask, request, abort
import requests
import json
import urllib
import re
from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,MemberJoinedEvent,FollowEvent,StickerSendMessage,MemberLeftEvent,ImageSendMessage,JoinEvent,AudioSendMessage,FlexSendMessage

)


app = Flask(__name__)

line_bot_api = LineBotApi('TPPSWL0iQ0xieaDP9Z+3fnZ1t7wF02z7/KlIQBzVSVZNyZleMCyz5AIWVVyQCLei4yD0sE/zWQUTY3D8VB3pjeljI4auX7BQgXgL+H2RQEuCaBPdE1OtJdziKDrWSEqnu2JrmcR/PKYhO/pCgte8GgdB04t89/1O/w1cDnyilFU=')
handler = WebhookHandler('52d0745099ebb304a9c7561bcb74e32b')


@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)
    return 'OK'
  

@handler.add(JoinEvent)#進群打招呼
def join(event):
    sticker=StickerSendMessage(package_id='789',sticker_id='10855')
    line_bot_api.reply_message(event.reply_token,sticker)
@handler.add(FollowEvent)#加好友介紹功能
def follow(event):
    uid = event.source.user_id
    profile = line_bot_api.get_profile(uid)
    name = profile.display_name
    message = TextSendMessage(text=f'{name}歡迎')
    mes2=TextSendMessage(text="可以問我去哪裡吃飯或吃甚麼喔")
    mes3=TextSendMessage(text="更可以問我什麼是什麼東西喔")
    ca=json.load(open('ca.json','r',encoding='utf-8'))
    mes4=FlexSendMessage('profile',ca)
    line_bot_api.reply_message(event.reply_token, [message,mes2,mes3,mes4])
@handler.add(MemberJoinedEvent)#入群歡迎詞
def welcome(event):
    uid = event.joined.members[0].user_id
    gid = event.source.group_id
    profile = line_bot_api.get_group_member_profile(gid, uid)
    name = profile.display_name
    message = TextSendMessage(text=f'{name}歡迎加入')
    line_bot_api.reply_message(event.reply_token,message )
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    Sendstring=''
    if"去哪吃" in event.message.text:
        Sendstring="當然是"+places()+"啊"
    elif"吃什麼" in event.message.text:
        Sendstring="當然是"+foods()+"啊"
    elif"要不要" in event.message.text:
        Sendstring="我覺得"+ok()
    elif"什麼時候" in event.message.text:
        Sendstring=when()
    elif"什麼是" in event.message.text:
        cmd=event.message.text.split("是")
        try:
         q_string = {'tbm': 'isch', 'q': cmd[1]}
         url = f"https://www.google.com/search?{urllib.parse.urlencode(q_string)}/"
         headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36'}

         req = urllib.request.Request(url, headers = headers)
         conn = urllib.request.urlopen(req)

         print('fetch conn finish')

         pattern = 'img data-src="\S*"'
         img_list = []

         for match in re.finditer(pattern, str(conn.read())):
          img_list.append(match.group()[14:-1])

         random_img_url = img_list[random.randint(0, len(img_list)+1)]
         print('fetch img url finish')
         print(random_img_url)

         line_bot_api.reply_message(
         event.reply_token,
         ImageSendMessage(
         original_content_url=random_img_url,
         preview_image_url=random_img_url)
         )
        except:
         line_bot_api.reply_message(
         event.reply_token,
         TextSendMessage(text="我....我也不知道"))
    elif"天氣" in event.message.text:
        weather=event.message.text.split(" ")
        city = weather[1]
        city = city.replace('台','臺')
        line_bot_api.reply_message(event.reply_token, TextSendMessage)(
        text=city+"的我不知道"
        )
    else:
        Sendstring=event.message.text
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=Sendstring))
def places():
    placeslist=["南樓","北樓","中正樓"]
    return placeslist[random.randint(0,len(placeslist)-1)]
def foods():
    foodslist=["炒飯","肉燥飯","肉燥麵","紅油炒手","水餃","關東煮",]
    return foodslist[random.randint(0,len(foodslist)-1)]
def ok():
    oklist=["要","不要"]
    return oklist[random.randint(0,len(oklist)-1)]
def when():
    whenlist=["現在","等一下","五分鐘後","一小時後","不久的將來","未來","不要做啦"]
    return whenlist[random.randint(0,len(whenlist)-1)]
@app.route('/')
def index():
     return 'hello world'
import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
