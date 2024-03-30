# Create URL previews for Apple Pages. Michael Z Freeman 2024.
#
# Apple Pages can "preview" Youtube links. Pages automatically embeds the Youtube video. But that's about it.
# Pasted URL's can be auto linked by hitting return. But then we end up with a massive list of hard to read links, so why not actually have proper previews of links such as X and Facebook have ?
# At first I looked at directly constructing a binary object to be inserted into the macOS clipboard. This was based on examining an image and link copied from an Apple Pages document. However this becomes over complicated, as does using Applescript.
# So I finally settled on keeping the solution in Python. Appending the
# previews and links to a Word document means that can be opened in Apple
# Pages.
from linkpreview import Link, LinkPreview, LinkGrabber
import os
from colorama import Fore, Back, Style
import urllib.request
import requests
import cgi
import re
from html2image import Html2Image
from docx import Document
from htmldocx import HtmlToDocx
from numpy import loadtxt
import subprocess
from playwright.sync_api import sync_playwright

hti = Html2Image(output_path='PREVIEWS')

# This function


def URLFind(string):

    # findall() has been used
    # with valid conditions for urls in string
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    url = re.findall(regex, string)
    return [x[0] for x in url]

ua = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/69.0.3497.100 Safari/537.36"
)
with open('links.txt', 'r') as file:
    links = file.read().splitlines()
length = len(links)
base_path = r"PREVIEWS"
ext = r".html"
absolutefavicon = "Favicon"
document = Document()
new_parser = HtmlToDocx()
for i in range(length):
        URL = links[i]
        file_path = base_path + "/" + str(i) + ext
        pic_path = str(i) + ".png"
        print(file_path)
        grabber = LinkGrabber()
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)
            page = browser.new_page(user_agent=ua)
            page.goto(URL)
            page.wait_for_timeout(2000)
            
            content = page.content()
        #content, URL = grabber.get_content(URL)
        link = Link(URL, content)
        preview = LinkPreview(link, parser="lxml")
        #print(content)
        if preview.title is None:
            preview.title = "None"
        print("title:", Fore.GREEN + preview.title + Fore.WHITE)
        if preview.description is None:
            preview.description = "None"
        print("description:", Fore.GREEN + preview.description + Fore.WHITE)
        if preview.image is None:
            preview.image = "None"        
        print("image:", Fore.MAGENTA + preview.image + Fore.WHITE)
        if preview.force_title is None:
            preview.force_title = "None"          
        print("force_title:", Fore.GREEN + preview.force_title + Fore.WHITE)
        if preview.absolute_image is None:
            preview.absolute_image = "None"         
        print("absolute_image:", Fore.MAGENTA + preview.absolute_image + Fore.WHITE)
        if preview.site_name is None:
            preview.site_name = "None"         
        print("site_name:", Fore.GREEN + preview.site_name + Fore.WHITE)
        if preview.favicon is None:
            preview.favicon = "None"         
        print("favicon:", Fore.MAGENTA + str(preview.favicon) + Fore.WHITE)
        if preview.absolute_favicon is None:
            preview.absolute_favicon = "None"         
        print("absolute_favicon:", Fore.MAGENTA + str(preview.absolute_favicon) + Fore.WHITE)
        # os.system("aria2c " + preview.absolute_image)
        # print("Urls: ", (URLFind(str(preview.absolute_favicon))[0]))
        print(preview.absolute_favicon)
        if not preview.absolute_favicon:
            print("preview.absolute_favicon is empty.")
        else:
            absolutefavicon = URLFind(str(preview.absolute_favicon))[0]
        # to open/create a new html file in the write mode
        f = open(file_path, 'w')
        # the html code which will go in the file GFG.html
        html_template = f"""<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Document</title>
            <link rel="stylesheet" href="styles.css">
        </head>
        <body>
            <div id="LinkPreview"">
                <div class="preview">
                    <img class="preview" src="{preview.absolute_image}">
                    <div class="preview-content">
                        <div class="title">{preview.title}</div>
                        <p class="text">{preview.description}</p>
                        <div class=favicon"><img class=favicon alt="Link Preview Favicon" src="{
            absolutefavicon}">
                            <a href="{URL}" target="_blank" style="line-height: 1.5;">{URL}</a>
                        </div>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """
        # writing the code into the file
        f.write(html_template)
        # close the file
        f.close()
        # screenshot an HTML file
        hti.screenshot(
            html_file=file_path,
            css_file='PREVIEWS/styles.css',
            save_as=pic_path)
        # Image created by html2image needs cropping
        process = subprocess.Popen(f'gm convert /Users/michaelzfreeman/UseTheSourceLuke/linkpreview/PREVIEWS/{pic_path} -trim /Users/michaelzfreeman/UseTheSourceLuke/linkpreview/PREVIEWS/CROPPED/{pic_path}', shell=True)
        process.wait()
        # gm convert page.png -crop 1632x384+0+0 cropped.png
        html = f'<img src="PREVIEWS/CROPPED/{pic_path}" /><p style="text-align: center;"><small><a href="{URL}">{preview.title}</a></small>'
        #print(html)
        new_parser.add_html_to_document(html, document)
document.save('PREVIEWS/previews.docx')
