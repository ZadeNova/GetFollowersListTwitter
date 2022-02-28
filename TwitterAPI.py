import tweepy
import pandas as pd
from openpyxl import load_workbook


ACCESS_TOKEN = ""  # Put your access token [ Api public key ]
ACCESS_SECRET = "" # Put your Access secret [ Api secret key ]
BearerToken = "" # Put your bearer key here

client = tweepy.Client(bearer_token=BearerToken)
UserID = ""
theUser = ""
# Enter Username properly
try:
    UserName = input("Enter the correct username: ")
    theUser = client.get_user(username=UserName)

except Exception as e:
    print(e)

print(theUser.data.id)
print(theUser.data.name)
print(theUser.data.username)
UserID = theUser.data.id


UserFollowing = client.get_users_following(id=UserID , user_fields=['profile_image_url'])
UserFollowers = client.get_users_followers(id=UserID)

#print(UserFollowing.data)
#print(UserFollowers.data)
SortedUserFollowing = []
SortedUserFollowers = []
if UserFollowing.data != None:

    UserFollowingDict = {i.id:i.username for i in UserFollowing.data}
    SortedUserFollowing = sorted(UserFollowingDict.items(), key=lambda x: x[1]) # Sort by alphabetical order
if UserFollowers.data != None:

    UserFollowersDict = {i.id:i.username for i in UserFollowers.data}
    SortedUserFollowers = sorted(UserFollowersDict.items(), key=lambda x: x[1])  # Sort by alphabetical order



FinalSortedUserFollowers = {"ID[Followers]" : [ str(a[0]) for a in SortedUserFollowers ] , "Username[Followers]" : [ a[1] for a in SortedUserFollowers ]}
FinalSortedUserFollowing = {"ID[Following]" : [ str(a[0]) for a in SortedUserFollowing ] , "Username[Following]": [ a[1] for a in SortedUserFollowing ]}


# Pandas
df = pd.DataFrame(data=FinalSortedUserFollowers)
df2 = pd.DataFrame(data=FinalSortedUserFollowing)
with pd.ExcelWriter("Data.xlsx") as writer:

    df.to_excel(writer,sheet_name="Followers" , startrow=10)
    df2.to_excel(writer,sheet_name="Following" , startrow=10)

workbook = load_workbook(filename="Data.xlsx")
sheet = workbook["Followers"]
sheet2 = workbook["Following"]
sheet["B3"] = "UserTwitterID"
sheet["C3"] = "name"
sheet["D3"] = "Username"
sheet["B4"] = str(theUser.data.id)
sheet["C4"] = theUser.data.name
sheet["D4"] = theUser.data.username
# Sheet 2
sheet2["B3"] = "UserTwitterID"
sheet2["C3"] = "name"
sheet2["D3"] = "Username"
sheet2["B4"] = str(theUser.data.id)
sheet2["C4"] = theUser.data.name
sheet2["D4"] = theUser.data.username
workbook.save(filename="Data.xlsx")

# Testing python script for small side project.

