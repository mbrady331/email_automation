import yagmail
import pandas
import datetime
import time
from news import NewsFeed


def send_email():
    today = datetime.datetime.now().strftime('%Y-%m-%d')
    yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    news_feed = NewsFeed(interest=row['interest'],
                         from_date=yesterday,
                         to_date=today)
    email = yagmail.SMTP(user='your email address here', password="email password here")
    email.send(to=row['email'],
               subject=f"Your{row['interest']} news for the day!",
               contents=f"Hello {row['name']},\n\n See what's new about {row['interest']} today.\n\n{news_feed.get()}")


while True:
    if datetime.datetime.now().hour == 15 and datetime.datetime.now().minute == 30:

        df = pandas.read_excel('people.xlsx')

        for index, row in df.iterrows():
            send_email()

    time.sleep(60)

