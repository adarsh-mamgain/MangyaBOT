from __future__ import print_function
import os, yagmail

import discord
from discord.ext import commands
from discord import Embed

import json
from datetime import datetime
import pytz

import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from flask import Flask
from threading import Thread

TOKEN = os.environ['DISCORD_TOKEN']

client = discord.Client()
bot = commands.Bot(command_prefix='<')

@bot.event
async def on_guild_join(guild):
  embed = Embed(title="<add",description="Type '<add' to register your server and upload your class Time-Table.",colour=0x6900d3)
  await guild.system_channel.send(embed=embed)

@bot.command()
async def administrator(ctx):
  if ctx.guild.id == 784651679227969536:
    SERVER_ID = "1RKoIECHtN-DY7ale9a_gWm1ETc3qmkRyL8ZiylfoMrQ"
    result = sheet.values().get(spreadsheetId=SERVER_ID,range="Sheet1!A2:Z").execute()
    values = result.get('values', [])
    values = [['784651679227969536', 'Adarsh Mamgain', '784651679227969539', 'general', '760153623723900982', 'Mangya'], ['841612776567996456', 'Testing Server', '841612776567996459', 'general', '760153623723900982', 'Mangya']]
    total = 0

    embed = Embed(title="BOT UPDATE",description="Now you can get links for any day order from 1-5 or a particular date for the current month.\n\nType '<day [day_order]'\nEG: '**<day 5**'\n\nType '<date [yyyy-mm--dd]'\nEG: '**<date 2021-05-01**'\n\nType '**<help**' to know more about particular commands.",colour=0x6900d3)
    embed.set_thumbnail(url="https://raw.githubusercontent.com/adarsh-mamgain/image-server/main/Mangya.jpg")
    for a in values:
      guild = bot.get_guild(int(a[0]))
      if guild is None:
        guild = await bot.fetch_guild(int(a[0]))
      channel = guild.get_channel(int(a[2]))
      await channel.send(embed=embed)
      print(guild)
      total+=1
    print("Total: ", total)
    await ctx.send(f"Total: {total}")

bot.remove_command("help")

@bot.command()
async def help(ctx):
  embed = Embed(title="MangyaBOT",description="Add MangyaBOT to your server -> http://bit.ly/MangyaBOT",url="http://bit.ly/MangyaBOT",colour=0x6900d3)
  embed.set_author(name=f"{ctx.author}")
  embed.set_thumbnail(url=f"{ctx.author.avatar_url_as(format='png')}")
  embed.set_image(url="https://raw.githubusercontent.com/adarsh-mamgain/image-server/main/donate.JPG")
  embed.add_field(name="**HELP COMMAND:**", value="List of all available commands of Mangya Bot.", inline=False)
  embed.add_field(name="**add:**", value="Register your server with Course and Time-Table to start recieving links.", inline=False)
  embed.add_field(name="**link:**", value="Responds with todays Time-Table and links.", inline=False)
  embed.add_field(name="**day:**", value="Responds with requested day order\'s (1-5) Time-Table and links.", inline=False)
  embed.add_field(name="**date:**", value="Responds with requested date\'s Time-Table and links.", inline=False)
  embed.add_field(name="**donate:**", value="Donate to help run this BOT.", inline=False)
  await ctx.send(embed=embed)

@bot.command(name='donate', help='Donate to help run this BOT.')
async def donate(ctx):
  embed = Embed() 
  embed = Embed(title="DONATE", description="You can support me by donating a small amount and keep this BOT running.\n\nScan or Download the below QR Code to pay through any UPI app.\n\nThank You :)",colour=0x6900d3)
  embed.set_image(url="https://raw.githubusercontent.com/adarsh-mamgain/image-server/main/donate.JPG")
  await ctx.send(embed=embed)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = None
if os.path.exists('token.json'):
  creds = Credentials.from_authorized_user_file('token.json', SCOPES)

if not creds or not creds.valid:
  if creds and creds.expired and creds.refresh_token:
    creds.refresh(Request())
  else:
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
  with open('token.json', 'w') as token:
    token.write(creds.to_json())

service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

@bot.command(name='add', help='Register your server with Course and Time-Table to start recieving links.')
async def add(ctx):
  try:
    SERVER_ID = "1RKoIECHtN-DY7ale9a_gWm1ETc3qmkRyL8ZiylfoMrQ"
    result = sheet.values().get(spreadsheetId=SERVER_ID,range="Sheet1!A2:A").execute()
    values = result.get('values', [])

    boolean = False
    for server in values:
      if str(ctx.guild.id) == server[0]:
        boolean = True
        break
    if boolean:
      embed = Embed(title="Already Registered",url="https://docs.google.com/spreadsheets/d/1-GZbJVAOyXaPri64Q-5LEVFHYMOHM4CZAS2pjomugV0/edit#gid=637994594",description="You have already registered your bot :)\nYou can update your time-table by clicking on above link.",colour=0x6900d3)
      await ctx.send(embed=embed)
    else:
      # Call the Sheets API, appends data
      server_data = [
        [f"{ctx.guild.id}"], [f"{ctx.guild.name}"], [f"{ctx.channel.id}"], [f"{ctx.channel.name}"], [f"{ctx.message.author.id}"], [f"{ctx.message.author.name}"]
      ]

          
      password = os.environ['TOKEN']
      print("Server => ",server_data)
      yag = yagmail.SMTP("mangyabot@gmail.com", password)
      yag.send(subject="!! Mangya Alert !!", contents=server_data)

      request_body = {
        'requests': [
          {
            'duplicateSheet': {
              'sourceSheetId': '2117286483',
              'newSheetName': f'{ctx.guild.name} || {ctx.guild.id}'
            }
          }
        ]
      }
      request = sheet.batchUpdate(spreadsheetId="1-GZbJVAOyXaPri64Q-5LEVFHYMOHM4CZAS2pjomugV0", body=request_body).execute()
      print("sheet copied", request)

      sheet.values().append(spreadsheetId=SERVER_ID,range="Sheet1!A:Z",
      body={
        "majorDimension": "COLUMNS",
        "values": server_data
      },
      valueInputOption="USER_ENTERED").execute()

      embed = Embed(title="Add your Time Table",description="After following the link read the 'Instruction' sheet and follow the steps to add your classes data.",colour=0x6900d3,url="https://docs.google.com/spreadsheets/d/1-GZbJVAOyXaPri64Q-5LEVFHYMOHM4CZAS2pjomugV0/edit#gid=637994594")
      await ctx.send(embed=embed)
  except:
    embed = Embed(title="!! REGISTRATION ERROR\n!! Mail Us: mangyabot@gmail.com",description="Couldn't register your bot, let us know by emailing the error at the above mentioned email address.",colour=0x6900d3)
    await ctx.send(embed=embed)

''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
# get the date of 'Asia/Kolkata' timezone
now = datetime.now()
day = pytz.timezone('Asia/Kolkata')
dayIndex = now.astimezone(day).date()

async def print_today(ctx,today):
  if today == "Error":
    today = None
    await ctx.send("-------------------------")
    await ctx.send("ERROR: The BOT faced an error :/")
    await ctx.send("-------------------------")
    print("-------------------------")
    print("ERROR: The BOT faced an error :/")
    print("-------------------------")
  elif today == "Holiday":
    today = None
    await ctx.send("-------------------------")
    await ctx.send("No class today :)")
    await ctx.send("-------------------------")
    print("-------------------------")
    print("No class today :)")
    print("-------------------------")
  else:
    pass

async def print_link(ctx,class_range):
  try:
    # Call the Sheets API
    
    RANGES = [f"{ctx.guild.id}!A3:B15",f"{ctx.guild.id}!{class_range[0]}",f"{ctx.guild.id}!{class_range[1]}"]
    result = sheet.values().batchGet(spreadsheetId="1-GZbJVAOyXaPri64Q-5LEVFHYMOHM4CZAS2pjomugV0",ranges=RANGES).execute()

    values = result.get('valueRanges', [])
    timing = values[0]["values"]
    className = values[1]["values"]
    link = values[2]["values"]
    print(values)
    if not values:
      print('No classes found :(')
      await ctx.send("No classes found :(")
    else:
      print("-------------------------")
      await ctx.send("-------------------------")
      for a,b,c in zip(timing,className,link):
        if b and c:
          embed = Embed(title=f"{b[0]}",colour=0x6900d3,url=f"{c[0]}",description=f"{a[0]} - {a[1]}")
          await ctx.send(embed=embed)
          print('%s - %s, %s' % (a[0],a[1],b[0]))
        elif b and not c:
          embed = Embed(title=f"{b[0]}",colour=0x6900d3,description=f"{a[0]} - {a[1]}")
          await ctx.send(embed=embed)
          print('%s - %s, %s' % (a[0],a[1],b[0]))
        else:
          continue
      await ctx.send("-------------------------")
      print("-------------------------")
  except:
    embed = Embed(title="!!Complete your registration\n!! Mail Us: mangyabot@gmail.com",colour=0x6900d3,description="You haven't linked your time-table, complete it to recieve class links. Type '<add' to register. If problem persist mail us at the above email address.")
    await ctx.send(embed=embed)

''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''
''' ---------------------------------------------------- '''

@bot.command(name='link', help='Responds with today\'s Time-Table and links.')
async def link(ctx):
  try:
    # get the day-order details
    dayFile = open("Day.json")
    dayOrder = json.load(dayFile)
    today = dayOrder[f"{dayIndex}"]
    if today =="Holiday":
      await print_today(ctx,today)
    else:
      class_range = dayOrder[f"{today}"]
      print(dayIndex)
      await print_link(ctx,class_range)
      dayFile.close()
  except:
    today = "Error"
    await print_today(ctx,today)
    
@bot.command(name='day', help='Responds with requested day order\'s (1-5) Time-Table and links.')
async def day(ctx, getDay):
  try:
    global today
    # get the day-order details
    dayFile = open("Day.json")
    dayOrder = json.load(dayFile)
    class_range = dayOrder[f"{getDay}"]
    print(getDay)
    await print_link(ctx,class_range)
    dayFile.close()
  except:
    embed = Embed(title="Uh Oh!",description="Please type the day like: **[<day 5]** and the day order must be **between 1-5**.",colour=0x6900d3)
    await ctx.send(embed=embed)
    
@bot.command(name='date', help='Responds with requested date\'s Time-Table and links.')
async def date(ctx, getDate):
  try:
    # get the day-order details
    dayFile = open("Day.json")
    dayOrder = json.load(dayFile)
    today = dayOrder[f"{getDate}"]
    print(getDate)
    if today =="Holiday":
      await print_today(ctx,today)
    else:
      class_range = dayOrder[f"{today}"]
      print(dayIndex)
      dayFile.close()
      await print_link(ctx,class_range)
  except:
    embed = Embed(title="Uh Oh!",description="Please type the date like: **[<date 2021-05-15]**",colour=0x6900d3)
    await ctx.send(embed=embed)
    
@bot.command(name='exam', help='Responds with upcoming exam dates and timings.')
async def exam(ctx):
  file = open("Exam.json")
  exam = json.load(file)
  if not exam:
    await ctx.send("-------------------------")
    await ctx.send("No exams yet :)")
    await ctx.send("-------------------------")
    print("-------------------------")
    print("No exams yet :)")
    print("-------------------------")
  else:
    for x in exam:
      embed = Embed(title=f"{x}: {exam[x][0]}",description=f"{exam[x][1]} to {exam[x][2]}",colour=0x6900d3)
      await ctx.send(embed=embed)

bot.remove_command("exam")

"""
The below is Flask server to make this repl run 24/7
"""
app = Flask('')

@app.route('/')
def home():
  return "Hello Mangya is alive :)"

def run():
  app.run(host='0.0.0.0',port=8080)
  
def keep_alive():
  t = Thread(target=run)
  t.start()

keep_alive()
bot.run(TOKEN)