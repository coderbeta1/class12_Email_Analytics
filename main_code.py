import time
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd
from collections import Counter

#Introduction
def fig():
    print('''
------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                    Email_Analytics
------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''')
    time.sleep(1)
    print('''
                                     _______                _  _     _______                _                _
                                    (_______)              (_)| |   (_______)              | |          _   (_)
                                     _____    ____   _____  _ | |    _______  ____   _____ | | _   _  _| |_  _   ____   ___
                                    |  ___)  |    \ (____ || || |   |  ___  ||  _ \ (____ || || | | |(_   _)| | / ___) /___)
                                    | |_____ | | | |/ ___ || || |   | |   | || | | |/ ___ || || |_| |  | |_ | |( (___ |___ |
                                    |_______)|_|_|_|\_____||_|\_)   |_|   |_||_| |_|\_____| \_)\__  |  \__)|_|  \____)(___/
                                                                                               (____/
                                                              ''')
    time.sleep(1)
    print('''
------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                    Email_Analytics
------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''')

#Main Code
def main():
    try:
        df2=pd.DataFrame()
        while True:
            if df2.empty==False:
                df=df2
                received = df[df['Sent/Received'] == 'Received']
            else:
                df = pd.read_excel('Email_Analytics.xlsx')
                received = df[df['Sent/Received'] == 'Received']

            time.sleep(1)
            print('''
------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                                                Please select a command from Below
------------------------------------------------------------------------------------------------------------------------------------------------------------------------


                                                   +---------------------------------------------------------------------+
                                                   |           *1.View a Mail Category                                   |
                                                   |           *2.View Mails from  a Specific Sender                     |
                                                   |           *3.Read a Mail with subject containg _________            |
                                                   |           *4.Delete a Mail Category                                 |
                                                   |           *5.Delete Mails from a Specific Sender                    |
                                                   |           *6.Delete a Mail with subject containg _________          |
                                                   |           *7.View Common Words in Subject Headers                   |
                                                   |           *8.View Weekly Email Traffic                              |
                                                   |           *9.View Top Email Senders                                 |
                                                   |           *10.View Subject Word Count                               |
                                                   |           *11.View Monthly Email Traffic                            |
                                                   |           *12.View Hourly Email Traffic                             |
                                                   |           *13.Close Programme                                       |
                                                   +---------------------------------------------------------------------+
                         ''')
            ch=int(input('''
                                                   ....Select option.... '''))

            if ch==3:
                words=input('''
                                                   What Word Sould it Contain : ''')
                alpha=df[df['Subject'].str.contains(words)]
                print(alpha)


            elif ch==1:
                print('''
                                                    1.Spam Mails
                                                    2.Important Mails
                                                    3.Inbox Mails
                                                    4.Sent Mails
                                                    5.Recieved Mails
                                                    6.Social Mails
                                                    7.All Mails      ''')

                sc_1=int(input('''
                                                    Enter the Preferred Category :: '''))

                if sc_1==1:
                    spam=df[df['Category'].str.contains('Spam')]
                    print('\n SPAM MAILS:-')
                    print('\n',spam)

                elif sc_1==2:
                    important=df[df['Category'].str.contains('Important')]
                    print('\n IMPORTANT:-')
                    print('\n',important)


                elif sc_1==3:
                    inbox=df[df['Category'].str.contains('Inbox')]
                    print('\n INBOX:-')
                    print('\n',inbox)

                elif sc_1==4:
                    sent=df[df['Sent/Received'].str.contains('Sent')]
                    print('\n SENT MAILS:-')
                    print('\n',sent)

                elif sc_1==5:
                    recieved=df[df['Sent/Received'].str.contains('Received')]
                    print('\n RECIEVED MAILS:-')
                    print('\n',recieved)

                elif sc_1==6:
                    a=['Facebook','LinkedIn','CodeChef','Instagram','Twitter','Snapchat','TikTok']
                    social=df[df['From (Sender)'].str.contains('|'.join(a))]
                    print('\n SOCIAL MAILS:-')
                    print('\n',social)

                elif sc_1==7:
                    print('\n ALL MAILS:-')
                    print('\n',df)

                else:
                    print('NONE')

            elif ch==2:
                sender=input('''
                                                  Enter the Name/Email ID of sender : ''')
                alpha=df[df['From (Sender)'].str.contains(sender)]
                if alpha.empty==True:
                    alpha=df[df['From (Email ID)'].str.contains(sender)]
                print(alpha)

            elif ch==5:
                sender=input('''
                                                   Whose mails to delete : ''')
                alpha=df[df['From (Sender)'].str.contains(sender)]
                df2=df.drop(alpha.index, axis=0 )
                print('''
                                                   The mails have been deleted ''')
                time.sleep(1.5)

            elif ch==6:
                words=input('''
                                                   What Word Sould it Contain : ''')
                alpha=df[df['Subject'].str.contains(words)]
                if alpha.empty==True:
                    alpha=df[df['From (Email ID)'].str.contains(words)]
                df2=df.drop(alpha.index, axis=0)
                print('''
                                                   The mails have been deleted ''')
                time.sleep(1.5)

            elif ch==4:
                print('''
                                                    1.Spam Mails
                                                    2.Important Mails
                                                    3.Inbox Mails
                                                    4.Sent Mails
                                                    5.Recieved Mails
                                                    6.Social Mails
                                                    7.All Mails      ''')

                sc_1=int(input('''
                                                    Enter the Preferred Category :: '''))

                if sc_1==1:
                    alpha=df[df['Category'].str.contains('Spam')]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                elif sc_1==2:
                    alpha=df[df['Category'].str.contains('Important')]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)


                elif sc_1==3:
                    alpha=df[df['Category'].str.contains('Inbox')]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                elif sc_1==4:
                    alpha=df[df['Sent/Received'].str.contains('Sent')]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                elif sc_1==5:
                    alpha=df[df['Sent/Received'].str.contains('Received')]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                elif sc_1==6:
                    a=['Facebook','LinkedIn','CodeChef','Instagram','Twitter','Snapchat','TikTok']
                    alpha=df[df['From (Sender)'].str.contains('|'.join(a))]
                    df2=df.drop(alpha.index, axis=0 )
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                elif sc_1==7:
                    df2=pd.DataFrame([['-','-','-','-','-','-','-','-','-','-']], columns=['Date', 'Month', 'Year', 'Day', 'Time',
                     'From (Sender)','From (Email ID)', 'Subject', 'Sent/Received', 'Category'])
                    print('''
                                                   The mails have been deleted ''')
                    time.sleep(1.5)

                else:
                    print('NONE')

            elif ch==7:

                word_list_2d = df['Subject'].str.split(' ').fillna('none').tolist()
                word_list_1d = [word for list in word_list_2d for word in list]


                word_list_1d = [word.lower() for word in word_list_1d]


                exclude_list = ['this', 'that', 'your', 'with', 'from']
                word_list_1d = [word for word in word_list_1d if word not in exclude_list and len(word)>3]


                common_words_map = Counter(word_list_1d).most_common(10)
                common_words = [pair[0] for pair in common_words_map]
                frequency = [pair[1] for pair in common_words_map]

                plt.figure()
                plt.barh(common_words, frequency, color = 'lightcoral', ec = 'black', linewidth = 1.25)
                plt.gca().invert_yaxis()
                plt.title('Most Common Words in Subjects', fontsize = 14 ,fontweight = 'bold')
                y = 0.15
                for i in range(len(frequency)):
                    if len(str(frequency[i])) == 3:
                        x = frequency[i] - 14
                    else:
                        x = frequency[i] - 10
                    plt.text(x,y,frequency[i], fontsize = 10,fontweight = 'bold')
                    y = y + 1
                plt.xticks([0,200])
                plt.xlabel('Occurrences', fontweight = 'bold', labelpad=-5)
                plt.show()

            elif ch==8:

                df['Day'] = pd.Categorical(df['Day'], categories= ['Mon','Tue','Wed','Thu','Fri','Sat', 'Sun'],ordered=True)

                count_sorted_by_day = pd.DataFrame(df['Day'].value_counts().sort_index())

                count_sorted_by_day.plot(kind='bar', color = 'blueviolet', linewidth = 2, ylim = [0,350])
                plt.title('Weekly Email Traffic', fontweight = 'bold' ,fontsize = 14)
                plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
                plt.grid()
                plt.show()

            elif ch==9:

                sender_top_20 =  received['From (Sender)'].value_counts().nlargest(20)
                sender_top_20_count = sender_top_20.values
                sender_top_20_names = sender_top_20.index.tolist()

                plt.figure()
                plt.barh(sender_top_20_names, sender_top_20_count, color = 'forestgreen', ec = 'black', linewidth = 1.0)
                plt.gca().invert_yaxis()
                plt.title('Top 20 Senders', fontsize = 14 ,fontweight = 'bold')
                plt.xlabel('Received Email Count', fontweight = 'bold')
                plt.tight_layout()
                plt.show()

            elif ch==10:
                df['Subject Word Count'] = df['Subject'].str.split(' ').str.len()

                plt.figure()
                plt.hist(df['Subject Word Count'], bins=15, color = 'slategray', ec = 'black')
                plt.axis([0, 30, 0, 450])
                plt.xlabel('Word Count', fontweight = 'bold')
                plt.ylabel('No. of Emails', fontweight = 'bold')
                plt.title('Subject Word Count Histogram', fontsize = 14, fontweight = 'bold')
                plt.show()

            elif ch==11:
                month = received['Month']

                count_sorted_by_month = month.value_counts()

                count_sorted_by_month.plot(marker = 'o', color = 'green')
                plt.title('Monthly Email Traffic', fontsize = 14, fontweight = 'bold')
                plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
                plt.xlabel("Month of the Year", fontweight = 'bold', labelpad = 15)
                plt.xticks(range(len(count_sorted_by_month.index)), count_sorted_by_month.index)
                plt.xticks(rotation=90)
                plt.grid()
                plt.show()

            elif ch==12:

                hour = received['Time'].str.split(':').str[0] + ':00'

                count_sorted_by_hour = hour.value_counts().sort_index()

                count_sorted_by_hour.plot(marker = 'o', color = 'green')
                plt.title('Hourly Email Traffic', fontsize = 14, fontweight = 'bold')
                plt.ylabel("Received Email Count", fontweight = 'bold', labelpad = 15)
                plt.xlabel("Hour of the Day", fontweight = 'bold', labelpad = 15)
                plt.xticks(range(len(count_sorted_by_hour.index)), count_sorted_by_hour.index)
                plt.xticks(rotation=90)
                plt.grid()
                plt.show()

            elif ch==13:
                print("\n..............................................................Thank you for using our service...........................................................................")
                break
                quit()

    except:
        main()



#fig()
#import recieve_mail
main()
