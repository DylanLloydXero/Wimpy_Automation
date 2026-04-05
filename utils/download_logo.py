import urllib.request
import re

url = 'https://commons.wikimedia.org/w/index.php?curid=30143027'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
html = urllib.request.urlopen(req).read().decode('utf-8')

# Find the original file link containing 'upload.wikimedia.org'
# Usually the full resolution image link on the page
match = re.search(r'href="(https://upload\.wikimedia\.org/wikipedia/commons/[^"]+logo.+?\.(?:png|svg))"i', html, re.IGNORECASE)
if not match:
    match = re.search(r'href="(https://upload\.wikimedia\.org/wikipedia/commons/[^"]+)"', html)

if match:
    img_url = match.group(1)
    if img_url.endswith('.svg') and '/thumb/' in html:
         # let's try to get a png thumb if the original is an SVG
         png_match = re.search(r'src="(https://upload\.wikimedia\.org/wikipedia/commons/thumb/[^"]+\.png)"', html)
         if png_match:
             img_url = png_match.group(1)
             
    print("Found image URL:", img_url)
    img_req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    with open('logo.png', 'wb') as f:
        f.write(urllib.request.urlopen(img_req).read())
    print("Downloaded to logo.png")
else:
    print("Image not found in HTML")
