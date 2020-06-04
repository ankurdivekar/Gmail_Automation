import ezgmail
from pathlib import Path
import pandas as pd
import random
import codecs
import os
import shutil
import time


def get_clean_attrib(attrib):
    if type(attrib) == str:
        return attrib.strip()
    else:
        return ''


simulate_emails = True
general_promo_mode = True
custom_promo_mode = True
promo_code = 'PromoTest'

email_list_xlsx = Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailList.xlsx')
email_list_xlsx_simulated = Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailList_Simulated.xlsx')
email_config_xlsx = Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailConfig.xlsx')

custom_promo_dir = Path(r'D:\TechWork\DJAV\SongoBingo\Files\GeneratedGames\Games_200508_IN\Game_200508_IN_5\Cards')
general_promo_dir = Path(r'D:\TechWork\DJAV\SongoBingo\Files\GeneratedGames\Games_200508_IN\Game_200508_IN_3\Cards')
sent_promo_dir = Path(r'D:\TechWork\DJAV\SongoBingo\Files\GeneratedGames\Games_200508_IN\Game_200508_IN_5\SentCards')

# Define visual separators
separator_1 = 25 * "~~"
separator_2 = 25 * "*~"

# Get the email list
df = pd.read_excel(email_list_xlsx)

# Get list of text for email subject and body
email_body_list = pd.read_excel(email_config_xlsx, sheet_name='Email body')['Messages'].tolist()
email_subject_list = pd.read_excel(email_config_xlsx, sheet_name='Email subject')['Subjects'].tolist()

# If promo code is not present in XLS raise error
if promo_code not in df.columns:
    raise Exception(f'Column with promocode not found in XLS: {promo_code}')

if custom_promo_mode:
    # Get list of all custom promo assets
    custom_promo_list = os.listdir(custom_promo_dir)

    # Make sent promo directory if it doesnt exist
    sent_promo_dir.mkdir(parents=True, exist_ok=True)

    # Check if enough promo assets are available
    emails_to_send = df[promo_code].str.contains('yes', case=False).sum()
    if emails_to_send > 0:
        if len(custom_promo_list) < emails_to_send:
            raise Exception('Insufficient number of custom promo assets.')
    else:
        raise Exception('No emails to send')

if general_promo_mode:
    # Get list of all general promo assets
    general_promo_list = os.listdir(general_promo_dir)

# Init
if not simulate_emails:
    ezgmail.init()

# Send emails one by one
card_counter = 0
email_counter = 0
for index, row in df.iterrows():
    firstname = get_clean_attrib(row["First Name"])
    lastname = get_clean_attrib(row['Last Name'])
    email_id = row['Email'].strip()
    status = row[promo_code]

    if type(status) == str:
        status = status.strip().lower()

    if status == 'yes':
        # Send email
        email_subject = random.choice(email_subject_list).format(firstname=firstname, lastname=lastname)
        email_body = random.choice(email_body_list).format(firstname=firstname, lastname=lastname)

        # Handle the carriage returns
        email_body = codecs.decode(email_body, 'unicode_escape')

        # Create attachment list
        email_attachment_paths = []
        email_attachment_files = []

        # Get the card to attach
        if custom_promo_mode:
            sent_promo_list = df[promo_code].tolist()
            card_not_found = True
            while card_not_found:
                if card_counter >= len(custom_promo_list):
                    raise Exception('Insufficient number of cards!')
                else:
                    promo_asset = custom_promo_list[card_counter]
                    if promo_asset in sent_promo_list:
                        card_counter += 1
                    else:
                        card_not_found = False

            # Rename card with added first name
            sent_promo = promo_asset.replace('.png', '') + f'_{firstname}{lastname}.png'.replace(' ', '')
            sent_promo_path = sent_promo_dir / sent_promo
            promo_asset_path = custom_promo_dir / promo_asset

            # Copy to sent cards folder
            shutil.copy(promo_asset_path, sent_promo_path)

            # Set as attachment for email
            email_attachment_paths += [str(sent_promo_path)]
            email_attachment_files.append(sent_promo)

            # Update dataframe
            df.iloc[index, df.columns.get_loc(promo_code)] = promo_asset

        if general_promo_mode:
            # Set as attachment for email
            email_attachment_paths += [str(general_promo_dir / p) for p in general_promo_list]
            email_attachment_files += general_promo_list

        if simulate_emails:
            print(f'{separator_1}\n'
                  f'Sending email to: {firstname} {lastname}\n'
                  f'at {email_id}\n'
                  f'with Subject: {email_subject}\n'
                  f'and body:\n{email_body}\n'
                  f'with {len(email_attachment_files)} attachments:\n{email_attachment_files}\n'
                  f'{separator_1}')
        else:
            print(f'{separator_1}\n'
                  f'Sending email to: {email_id}\n'
                  f'{separator_1}')
            ezgmail.send(
                recipient=email_id,
                subject=email_subject,
                body=email_body,
                attachments=email_attachment_paths)

        # Write back to xlsx
        if simulate_emails:
            df.to_excel(email_list_xlsx_simulated, sheet_name='Email List', index=False)
        else:
            df.to_excel(email_list_xlsx, sheet_name='Email List', index=False)

        # Increment count of emails sent
        email_counter += 1

        # Insert random delay
        if not simulate_emails:
            time.sleep(random.randint(2, 5))

print(f'\n\n{separator_2}\n'
      f'Status report:\nFinished sending {email_counter} emails!'
      f'\n{separator_2}')
