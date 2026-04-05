import urllib.request
import re

url = 'https://en.wikipedia.org/wiki/Wimpy_(restaurant)'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')

match = re.search(r'src="(//upload\.wikimedia\.org/wikipedia/[^"]+Wimpy_logo[^"]+\.png)"', html, re.IGNORECASE)
if not match:
    # fallback to any png in infobox
    match = re.search(r'class="infobox.+?src="(//upload\.wikimedia\.org/wikipedia/[^"]+\.png)"', html, re.DOTALL | re.IGNORECASE)

if match:
    img_url = 'https:' + match.group(1)
    print("Downloading:", img_url)
    req2 = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
    with open('logo.png', 'wb') as f:
        f.write(urllib.request.urlopen(req2).read())
    print("Done")
else:
    print("Could not find logo")
