!pip install aspose-words

import tweepy
import pandas as pd
import aspose.words as aw
consumer_key = ""
consumer_secret = ""
auth = tweepy.AppAuthHandler(consumer_key, consumer_secret)
api = tweepy.API(auth)
doc = aw.Document()
builder = aw.DocumentBuilder(doc)
builder.font.name = "Calibri"
font = builder.font

##### ONLY FIELD TO CHANGE DAILY #####
# Format must be like this: 01-25
current_date = "03-11"
#######################################

font.bold = True;
builder.write("\n" + "Chinese embassies in South Asia:" + "\n")
font.bold = False;

twitter_accounts_first_half = ["@China_Amb_India", "@CathayPak", "@ChinaEmbKabul", "@ChinaEmbSL", 
           "@PRCAmbNepal", "@zlj517", "@MFA_China", "@MEAIndia", "@DrSJaishankar"]

twitter_account_names_first_half = ["@ChinaEmbIndia", "@ChinaEmbPakistan", "@ChinaEmbAf", "@ChinaEmbSL",
                                    "@ChineseEmbNepal", "Chinese MFA", "Chinese MFA", "Indian MEA", "S. Jaishankar"]

count = 0
for account in twitter_accounts_first_half:
  tweets = api.user_timeline(screen_name=account, count=200, include_rts = True, tweet_mode = 'extended')
  final_tweets = []
  font.bold = True;
  builder.write("\n" + twitter_account_names_first_half[count])
  font.bold = False;
  for tweet in tweets[:15]:
     date_string = str(tweet.created_at)
     date = date_string[5:10]
     if date == current_date:
       tweet_text_raw = "Retweet: " + tweet.retweeted_status.full_text if tweet.full_text.startswith("RT @") else tweet.full_text
       tweet_text = tweet_text_raw.replace("\n", " ")
       final_tweets.append(tweet_text)
       builder.write("\n" + tweet_text + ": " + "https://twitter.com/twitter/statuses/" + str(tweet.id))
  if final_tweets == []:
    builder.write("\n" + "N/A")
  print(final_tweets)
  doc_name = "first half " + current_date + " test.docx"
  doc.save(doc_name)
  count+=1




