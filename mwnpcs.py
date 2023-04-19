import os
import time
import pandas as pd
from openpyxl import load_workbook

from cohost.models.block import AttachmentBlock, MarkdownBlock
from cohost.models.user import User

os.environ['COHOST_COOKIE'] = 's%3AuHC23-G7PjXau5W1eJh9cfV4amhFk7Nx.U9HjU%2B6y67t%2FX%2FmoiYKg03xE1veWg5iBU1slgMHbWGk'

t = time.localtime()
current_time = time.strftime("%H:%M:%S", t)


# Open Morrowind_NPCS excel file
npcs = pd.read_excel('Morrowind_NPCs.xlsx')
npcwb = load_workbook('Morrowind_NPCs.xlsx')
npcws = npcwb.active

# Use panda sample() to find unposted characters from Morrowind_NPCs
unposted = npcs.loc[(npcs['Posted'] == False)]
new_post = unposted.sample()
new_post_index = int(new_post.index.values)
# For manual blasting
# new_post_index = 25


# Because of indexes starting at 0 and the Excel heading row, the row number is 2 higher than the index.
excelRow = new_post_index + 2
name_value = "A{row}".format(row=excelRow)
post_value = "G{row}".format(row=excelRow)

# Changes post status to True, saves file
npcws[post_value].value = "True"
npcwb.save('Morrowind_NPCs.xlsx')


class MorrowindNPC:
    '''A class for Morrowind NPCS'''

    def __init__(self, id):
        self.id = id
        self.name = npcs['Name'][self.id]
        self.url = npcs['URL'][self.id]
        self.bio = npcs['Description'][self.id]
        self.quote = npcs['Quote'][self.id]
        self.pic = npcs['Image'][self.id]
        self.alt = npcs['Alt'][self.id]
        self.status = npcs['Posted'][self.id]

    def splitBio(self):
        # Splits two sentence description into two strings by splitting at double pipes in Excel
        fullBio = self.bio
        p1 = fullBio.split('||')[0]
        p2 = fullBio.split('||')[1]
        return [p1, p2]

    def getImagePath(self):
        imagePath = npcs['Image'][self.id]
        if pd.isnull(npcs.loc[characterIndex, 'Tamriel Rebuilt']) == False:
            image = 'screenshots/TR/{}'.format(imagePath)
        else:
            image = 'screenshots/{}'.format(imagePath)
        return image


newNPC = MorrowindNPC(new_post_index)

characterIndex = newNPC.id
characterName = newNPC.name
url = newNPC.url
characterQuote = newNPC.quote
quoteNull = pd.isnull(npcs.loc[characterIndex, 'Quote'])
image = newNPC.getImagePath()
altText = newNPC.alt
postStatus = newNPC.status
bioOne = newNPC.splitBio()[0]
bioTwo = newNPC.splitBio()[1]
tamrielRebuiltNull = pd.isnull(npcs.loc[characterIndex, 'Tamriel Rebuilt'])


# Function to add CSS stuff:
# Braces (in order) represent header name, UESP URL, link name, and description bios
def startTextBox(name, link, bio1, bio2):
    # TR CSS format variant
    if tamrielRebuiltNull == False:
        return """<div id="outer-border" style="width: 100%; margin: 1% auto; border: 6px double rgb(200, 187, 166); background-color: rgba(0, 0, 0, 0.9);">
        <div id="nameplate-bg" style="background: url(&quot;https://staging.cohostcdn.org/attachment/ef4a6606-bade-4496-8e1b-a0e1cac80ce7/trans-noise.png&quot;), url(&quot;https://staging.cohostcdn.org/attachment/b9b38de1-431f-4104-a9e5-3eb118405bd2/trans-noise02.png&quot;), url(&quot;https://staging.cohostcdn.org/attachment/e794a2d8-db71-404e-ae75-9098565f17cc/gradient-5D4F23.svg&quot;);">
        <div id="nameplate" style="background-color: rgb(86, 85, 74); border: 2px ridge rgba(81, 73, 52, 0.7);">
        <p id="name" style="color: rgb(200, 187, 166); margin: 1px auto; background-color: rgba(0, 0, 0, 0.9); width: 45%; text-align: center; height: 100%; font-size: 105%;">{}<sup><a href="https://www.tamriel-rebuilt.org/" style="color: rgb(110, 129, 232); text-decoration: none;">[TR]</a></sup></p>
        </div>
        </div>
        <div>
        <div id="text-box" style="border: 2px ridge rgb(200, 187, 166); margin: 1%; min-height: 150px;">
        <p id="name-text" style="color: rgb(200, 187, 166); margin: 1.5%;"><a href="{}" style="color: rgb(110, 129, 232); text-decoration: none;">{} </a>{}</p>
        <p style="color: rgb(200, 187, 166); margin: 1.5%;">{}</p>
        """.format(name, link, name, bio1, bio2)
    else:
        return """<div id="outer-border" style="width: 100%; margin: 1% auto; border: 6px double rgb(202, 183, 132); background-color: rgba(0, 0, 0, 0.9);">
        <div id="nameplate-bg" style="background: url(&quot;https://staging.cohostcdn.org/attachment/ef4a6606-bade-4496-8e1b-a0e1cac80ce7/trans-noise.png&quot;), url(&quot;https://staging.cohostcdn.org/attachment/b9b38de1-431f-4104-a9e5-3eb118405bd2/trans-noise02.png&quot;), url(&quot;https://staging.cohostcdn.org/attachment/e794a2d8-db71-404e-ae75-9098565f17cc/gradient-5D4F23.svg&quot;);">
        <div id="nameplate" style="background-color: rgba(202, 183, 132, 0.69); border: 2px ridge rgba(81, 73, 52, 0.7);">
        <p id="name" style="color: rgb(223, 200, 158); margin: 1px auto; background-color: rgba(0, 0, 0, 0.9); width: 45%; text-align: center; height: 100%; font-size: 105%;">{}</p>
        </div>
        </div>
        <div>
        <div id="text-box" style="border: 2px ridge rgb(202, 183, 132); margin: 1%; min-height: 150px;">
        <p id="name-text" style="color: rgb(223, 200, 158); margin: 1.5%;"><a href="{}" style="color: rgb(110, 129, 232); text-decoration: none;">{}</a> {}</p>
          <p style="color: rgb(223, 200, 158); margin: 1.5%;">{}</p>
        """.format(name, link, name, bio1, bio2)


def endTextBox():
    return """</div></div></div>"""


def createLink(url, linkName: str):
    return """<a href="{}" style="color: rgb(110, 129, 232); text-decoration: none;">{}</a>""".format(url, linkName)


def createParagraph(contents):
    return """<p style="color: rgb(223, 200, 158); margin: 1.5%;">{}</p>""".format(contents)


def createQuote(quote):
    # Checks if there is text in the "Quote" field
    characterQuote = quote

    if tamrielRebuiltNull == False:
        # TR Text color
        rbg = """rgb(200, 187, 166)"""
    else:
        # Default
        rbg = """rgb(223, 200, 158)"""

    if quoteNull == False:
        characterQuote = """<p style="color: {}; margin: 1.5%; border: 1px solid {}; background-color: rgba(0, 0, 0, 0.9); font-size: 92%; text-align: center; padding: 2.5%;"><em>{}</em></p>""".format(
            rbg, rbg, quote)
    else:
        characterQuote = ""
    return characterQuote


def main():
    cookie = os.environ.get('COHOST_COOKIE')
    if cookie is None:
        print('COHOST_COOKIE environment variable not set - please paste your cookie below')
        print('To skip this, please set the COHOST_COOKIE environment variable to the cookie you want to use')
        cookie = input('COHOST_COOKIE: ')
    user = User.loginWithCookie(cookie)
    project = user.getProject('morrowind-npcs')
    print('Logged in as: {}'.format(project))
    blocks = [
        MarkdownBlock(
            startTextBox(characterName, url, bioOne, bioTwo)),
        MarkdownBlock(createQuote(characterQuote)),
        MarkdownBlock(endTextBox()),
        AttachmentBlock(image, alt_text=altText)
    ]

    # Empty string where post title normally goes
    # print('Live at: {}'.format(p.url))

    p = project.post('',
                     blocks,
                     adult=False, draft=False, tags=['morrowind', 'The Elder Scrolls', 'pc games', 'The Cohost Global Feed', 'bot'])
    print(characterName)
    print('Index: ', characterIndex)
    print('Excel cell: ', name_value)
    print(url)
    print(bioOne, bioTwo)
    print(characterQuote)
    print('No Quote?', pd.isnull(npcs.loc[characterIndex, 'Quote']))
    print('Posted: ', bool(postStatus))
    print('Posted at: ', current_time)
    # print('Chosted')


if __name__ == '__main__':
    main()
