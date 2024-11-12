import os
import requests
from datetime import datetime
from glob import glob
import pandas as pd

# set the links to the docs
# first is the url to download a csv of the 'Master List' sheet, where the interesting stuff happens
# second is just the link to the sheet in gneeral, linked in the resultant feed
dl_url = 'https://docs.google.com/spreadsheets/d/1kspw-4paT-eE5-mrCrc4R9tg70lH2ZTFrJOUmOtOytg/export?format=csv&gid=0#gid=0'
link_url = 'https://docs.google.com/spreadsheets/d/1kspw-4paT-eE5-mrCrc4R9tg70lH2ZTFrJOUmOtOytg/edit?gid=0#gid=0'

# set up directory paths
script_dir = os.path.dirname(os.path.abspath(__file__))

# here we store our csvs for comparison
csv_directory = script_dir + '/csv'

# here we store our snippets of comparison html
item_directory = script_dir + '/items'

# create dirs if they don't exist
if not os.path.exists(csv_directory):
    os.makedirs(csv_directory)
    
if not os.path.exists(item_directory):
    os.makedirs(item_directory)

# load list of items
csvs = glob(os.path.join(csv_directory, '*'))
items = glob(os.path.join(item_directory, '*'))



#
# clear up old files
#
#

print('Cleaning up old data')

# sort our csvs so we can remove old checks
old_csvs = sorted(csvs, key=os.path.getmtime, reverse=True)

# Keep the 24 latest csvs, delete the others
old_csvs_to_delete = old_csvs[23:]

# Loop through and delete the files
for file in old_csvs_to_delete:
    os.remove(file)
    print(f'Deleted: {file}')

# do the same with html items
old_items = sorted(items, key=os.path.getmtime, reverse=True)

# Keep the 24 latest files, delete the others
old_items_to_delete = old_items[50:]

# Loop through and delete the files
for file in old_items_to_delete:
    os.remove(file)
    print(f'Deleted: {file}')

#
# Get CSV to diff
#
#

print('Downloading CSV')

# create filename to save
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
filename_with_date = f"gamepass_{current_time}.csv"
filepath = os.path.join(csv_directory, filename_with_date)

# download csv and store with timestamp
response = requests.get(dl_url)

with open(filepath, 'wb') as file:
    file.write(response.content)

print(f"CSV downloaded and saved as: {filepath}")

# reload data, this is necessary as we may have deleted data earlier
csvs = glob(os.path.join(csv_directory, '*'))
items = glob(os.path.join(item_directory, '*'))

#
# Diff the latest 2
#
#

# sort by newest to oldest and get latest 2 csvs
files_sorted_by_mtime = sorted(csvs, key=os.path.getmtime, reverse=True)
newest_files = files_sorted_by_mtime[:2]

# if not enough csvs, don't compare
if len(newest_files) < 2:
    print('Not enough files to diff, exit')
    exit()

print(f'Comparing "{newest_files[0]}" and "{newest_files[1]}"')

# Load the two CSV files
df_new= pd.read_csv(newest_files[0], skiprows=1, usecols=['Game', 'System', 'Status', 'Added', 'Removed'])
df_old = pd.read_csv(newest_files[1], skiprows=1, usecols=['Game', 'System', 'Status', 'Added', 'Removed'])

# Merge the two DataFrames to compare them
# `indicator=True` adds a column to show where each row came from.
diff = pd.merge(df_old, df_new, how='outer', indicator=True)

# Rows that exist only in the old CSV will have '_merge' as 'left_only'
# Rows that exist only in the new CSV will have '_merge' as 'right_only'
# Rows common to both will have '_merge' as 'both'

# Filter the rows that are new or old
new_rows = diff[diff['_merge'] == 'right_only']
old_rows = diff[diff['_merge'] == 'left_only']

changes = len(new_rows) + len(old_rows)

print(f"Number of changes: {changes}")

# if nothing changed, don't create an item
if changes == 0:
    print('No changes, exit')
    exit()

#
# Generate new rss entry
#
#

print('Creating new item')

# remove extra column
new_rows = new_rows.drop('_merge', axis=1)
old_rows = old_rows.drop('_merge', axis=1)

#r remove NaN - pandas likes to try parse every field
new_rows = new_rows.fillna('')
old_rows = old_rows.fillna('')

# Output differences to an HTML file
html_report = '<h2>New rows</h2>\n'
html_report += new_rows.to_html()
html_report += '\n<h2>Old rows</h2>\n'
html_report += old_rows.to_html()
html_report += '\n<br><br><a href="' + link_url + '">View View Google Sheet</a>'

if not os.path.exists(item_directory):
    os.makedirs(item_directory)

entry_filename = f"item_{current_time}.html"

entry_filepath = os.path.join(item_directory, entry_filename)

# Write the HTML report to a file
with open(entry_filepath, 'w') as file:
    file.write(html_report)

#
# generate rss
#
#

print('Generating feed')

pub_time = datetime.utcnow()

# Format the time according to RFC 1123
pub_time = pub_time.strftime('%a, %d %b %Y %H:%M:%S GMT')

# build start
rss = """<?xml version="1.0"?>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom">
   <channel>
      <title>GamePass Updates</title>
      <link>http://www.neilgaryallen.dev</link>
      <description>Feed of changes to the GamePass Master List Google Doc</description>
      <language>en-gb</language>
      <lastBuildDate>{time}</lastBuildDate>
      <docs>https://www.rssboard.org/rss-specification</docs>
      <webMaster>neil@biscuits.digital (Neil Allen)</webMaster>
"""

# replace time
rss = rss.format(time=pub_time)

# Get latest entries
items = glob(os.path.join(item_directory, '*'))

# order all our html comparison items and limit to ten
items_sorted_by_mtime = sorted(items, key=os.path.getmtime, reverse=True)
newest_items = items_sorted_by_mtime[:10]

# created our entries
for item in newest_items:
    lm = os.path.getmtime(item)
    lm = datetime.fromtimestamp(lm)
    lm = lm.strftime('%a, %d %b %Y %H:%M:%S GMT')
    with open(item, 'r') as file:
        i = file.read()
    item_data = """
    <item>
        <title>GamePass Spreadsheet Updated</title>
        <link>{link}</link>
        <description><![CDATA[{item}]]></description>
        <pubDate>{time}</pubDate>
    </item>
    """
    item_data = item_data.format(link=link_url, item=i, time=lm)
    rss += item_data
    
rss += """
    </channel>
</rss>"""

# write rss feel
with open(script_dir + '/feed.xml', 'w') as file:
    file.write(rss)

print('Complete!')