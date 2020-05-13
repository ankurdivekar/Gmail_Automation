import ezgmail
from pathlib import Path
import pandas as pd
import random
import codecs
import os
import shutil
import time

simulate_emails = True
promo_mode = True

email_list_xlsx =Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailList.xlsx')
email_list_xlsx_simulated =Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailList_Simulated.xlsx')
email_config_xlsx =Path(r'D:\TechWork\DJAV\Gmail_Automation\EmailConfig.xlsx')
promo_dir = Path(r'D:\TechWork\DJAV\SongoBingo\Files\GeneratedGames\Games_200508_IN\Game_200508_IN_3')


def get_clean_attrib(attrib):
    if type(attrib) == str:
        return attrib.strip()
    else:
        return ''


# Define directories for generated and sent cards
cards_dir = promo_dir / 'Cards'
sent_cards_dir = promo_dir / 'SentCards'
# Make directory if it doesnt exist
sent_cards_dir.mkdir(parents=True, exist_ok=True)

# Get promo code
promo_code = promo_dir.stem

# Get list of cards
promo_cards_list = os.listdir(cards_dir)

# Get the email list
df = pd.read_excel(email_list_xlsx, sheet_name='Email List')

# Get list of text for email subject and body
email_body_list = pd.read_excel(email_config_xlsx, sheet_name='Email body')['Messages'].tolist()
email_subject_list = pd.read_excel(email_config_xlsx, sheet_name='Email subject')['Subjects'].tolist()

# Check if enough promo assets are available
if not promo_mode:
    if len(promo_cards_list) < len(df['Email']):
        raise Exception('Insufficient number of promo assets. Please run with promo mode ON')

# If game code is not present in XLS
if promo_code not in df.columns:
    # Insert game code column
    df.insert(len(df.columns), promo_code, 'None')

# Init
if not simulate_emails:
    ezgmail.init()

# Send emails one by one
card_counter = 0

for index, row in df.iterrows():
    firstname = get_clean_attrib(row["First Name"])
    lastname = get_clean_attrib(row['Last Name'])
    email_id = row['Email'].strip()
    card = row[promo_code]

    if card in promo_cards_list:
        print(f'Email already sent to {firstname} {lastname}')
    else:
        email_subject = random.choice(email_subject_list).format(firstname=firstname, lastname=lastname)
        email_body = random.choice(email_body_list).format(firstname=firstname, lastname=lastname)

        # Handle the carriage returns
        email_body = codecs.decode(email_body, 'unicode_escape')

        # Get the card to attach
        if promo_mode:
            if len(promo_cards_list) > 1:
                gen_card = ', '.join(promo_cards_list)
            else:
                gen_card = promo_cards_list[0]
        else:
            sent_cards_list = df[promo_code].tolist()
            card_not_found = True
            while card_not_found:
                if card_counter >= len(promo_cards_list):
                    raise Exception('Insufficient number of cards!')
                else:
                    gen_card = promo_cards_list[card_counter]
                    if gen_card in sent_cards_list:
                        card_counter += 1
                    else:
                        card_not_found = False

        # Update dataframe
        df.iloc[index, df.columns.get_loc(promo_code)] = gen_card
        gen_card_path = cards_dir / gen_card

        if promo_mode:
            # Set as attachment for email
            email_attachments = [str(promo_dir / 'Cards' / p) for p in promo_cards_list]
        else:
            # Rename card with added first name
            sent_card = gen_card.replace('.png', '') + f'_{firstname}{lastname}.png'.replace(' ', '')
            sent_card_path = sent_cards_dir / sent_card

            # Copy to sent cards folder
            shutil.copy(gen_card_path, sent_card_path)

            # Set as attachment for email
            email_attachments = [str(sent_card_path)]

        if simulate_emails:
            print(f'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n'
                  f'Sending email to: {firstname} {lastname}\n'
                  f'at {email_id}\n'
                  f'with Subject: {email_subject}\n'
                  f'and body:\n{email_body}\n'
                  f'with {len(email_attachments)} attachments:\n{email_attachments}\n'
                  f'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
        else:
            print(f'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n'
                  f'Sending email to: {email_id}\n'
                  f'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            ezgmail.send(
                recipient=email_id,
                subject=email_subject,
                body=email_body,
                attachments=email_attachments)

        # Write back to xlsx
        if simulate_emails:
            df.to_excel(email_list_xlsx_simulated, sheet_name='Email List', index=False)
        else:
            df.to_excel(email_list_xlsx, sheet_name='Email List', index=False)

        # Insert random delay
        if not simulate_emails:
            time.sleep(random.randint(2, 5))

print('Done')
