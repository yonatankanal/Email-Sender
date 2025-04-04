import pandas as pd
import smtplib
from email.mime.text import MIMEText


def main():
    make_data()
    send_email()
    

def make_data():
    file_one_path = 'sample_files/spreadsheet_1.xlsx'
    file_two_path = 'sample_files/spreadsheet_2.xlsx'
    df_1 = pd.read_excel(file_one_path)
    df_cleaned_1 = df_1.dropna(subset=['Dates'])
    df_2 = pd.read_excel(file_two_path)
    df_cleaned_2 = df_2.dropna(subset=['Dates'])

    all_the_data = pd.merge(df_cleaned_1, df_cleaned_2, on=['First Name', 'Last Name'], how='inner')
    print (all_the_data)

    for index, row in all_the_data.iterrows():
        name = row['First Name']
        current = row['Practice Setting_x']
        future = row['Practice Setting_y']
        filler = input(f"Which email format for {name} (1, 2, or 3)? ")
        while filler != '1' and filler != '2' and filler != '3':
            filler  = input("***Input 1, 2, or 3*** Which email format for {name}? ")

        create_email(filler,name,current,future)
        move_on = input("Press enter to move on to the next student. Press any character and then press enter to choose another format. ")
        while move_on != '':
            filler = input(f"Which email format for {name} (1, 2, or 3)? ")
            while filler != '1' and filler != '2' and filler != '3':
                filler  = input("***Input 1, 2, or 3*** Which email format for {name}? ")
            create_email(filler,name,current,future)
            move_on = input("Press enter to move on to the next student. Press any character and then press enter to choose another format. ")



def create_email(filler,name,current,future):
    if filler == '1':
        body = f'''

            Dear {name},

            Congratulations on completing your first Level II Fieldwork at {current}! You made it! Although it may seem like you are starting over in some capacity at the {future}, remember that you have even more knowledge than you had 3 months ago. You’ve got this!

            I hope this week goes well—please let me know how things are going. Remember that you can always reach out to myself or Ann with any questions or concerns.

            Take care,
            Avital

        '''
    
    elif filler == '2':
        body = f'''

            Dear {name},

            Can you believe you are already at midterm for your second Level II Fieldwork! I hope things are going well. What are some of the high points or low points you’ve experienced?

            Feel free to reach out to either Ann or me with any questions or concerns. Enjoy the rest of your time at {current}.

            Take care,
            Avital

        '''

    elif filler == '3':
        body = f'''

            Dear {name},

            Congratulations on finishing your second Level II Fieldwork! How was it at {current}? You should feel very proud of your accomplishment. Have a wonderful summer and I’ll see you back here at BSP in August!

            Take care,
            Avital

        '''

    # print(f'\n\n{body}\n\n')


def send_email():
    subject = 'First Email With Python!!!'
    body = '''
    This is the coolest thing in the world!!!
    '''
    sender = 'yonatankanal@gmail.com'
    recipients = ['yonatankanal@gmail.com']
    password = 'suuy lobl ldhn dleg'


    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    with smtplib.SMTP_SSL('smtp-mail.outlook.com', 587) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())

    send_email(subject, body, sender, recipients, password)
    # print('Message Sent!')

if __name__ == "__main__":
    main()