from pymongo import MongoClient 

from utils.aws_utils import get_aws_secret
MONGODB_URL = get_aws_secret("MONGODB").get("URL")

def connectMongoDB():
    try: 
        connSecure = MongoClient(MONGODB_URL)
        print("Connected successfully to MongoDB!") 
    except Exception as e:
        print(str(e))
        print("Could not connect to MongoDB!")

    return connSecure