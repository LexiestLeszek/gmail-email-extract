import mailbox
import re
import csv

# Path to your mbox file
mbox_file = './inbox.mbox'

# Regex pattern to match fund links and names
pattern = r'"(http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+)" target="_blank">(.*?)</a>'

# Initialize a list to store the extracted data
extracted_data = []

# Open the mbox file
mbox = mailbox.mbox(mbox_file)

# Iterate over each email in the mbox file
for message in mbox:
    # Extract the subject of the email
    subject = message['subject']
    
    # Check if the email is relevant
    if subject.startswith("Subj 1"):
        # Extract the HTML content of the email
        email_content = message.get_payload()
        
        # Find all matches of the regex pattern in the email content
        matches = re.findall(pattern, email_content)
        
        # Iterate over each match and append to the extracted data list
        for match in matches:
            fund_link, fund_name = match
            extracted_data.append((fund_link, fund_name))

# Write the extracted data to a CSV file
output_file = 'extracted_data.csv'
with open(output_file, 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Fund Link', 'Fund Name'])
    writer.writerows(extracted_data)

print(f"Extraction completed! Extracted data has been saved to {output_file}.")
