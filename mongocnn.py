import os
import pymongo 
from bson.json_util import dumps
import pandas as pd


def main():
    uri = os.getenv('uri')
    collection_name = os.getenv('collection_name')
    db_name = os.getenv('db_name')

    try:
        client = pymongo.MongoClient(uri)
        if db_name in client.list_database_names():
            db = client[db_name]
        else:
            print("Error: databse not found")
            raise ValueError
        if collection_name in db.list_collection_names():
            collection = db[collection_name]
        else:
            print("Error:collection not found")
            raise ValueError

    except pymongo.errors.ServerSelectionTimeoutError as e:
        print ("Error: could not connect to database:{}".format(e))
        raise e

    results = collection.find(
          {"vote_average":{"$gt":"9.0"}},
          {"title":1,"vote_average":1}
          )

    with open('movies.json', 'w') as file:
        file.write('[')
        for document in results:
            file.write(dumps(document))
            file.write(',')
        file.write(']')

    mv_json = pd.read_json('movies.json')
    mv_frame = pd.DataFrame.from_dict(mv_json)
    mv_filter = mv_frame.filter(["title","vote_average"])

    writer = pd.ExcelWriter('movies.xlsx', engine='xlsxwriter')

    mv_filter.to_excel(writer, sheet_name='Movies')

    writer.close()

if __name__ == '__main__':
    main()

