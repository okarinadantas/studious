import os
import pymongo
import pandas as pd

def main():
    uri = os.getenv('uri')
    collection_name = os.getenv('collection_name')
    db_name = os.getenv('db_name')

    try:
        client = pymongo.MongoClient(uri, serverSelectionTimeoutMS=5000)
        db = client[db_name]
        collection = db[collection_name]
    except pymongo.errors.ServerSelectionTimeoutError as e:
        print("Error: Could not connect to the database: {}".format(e))
        raise e

    results = collection.find(
        {"vote_average": {"$gt": 9.0}},
        {"title": 1, "vote_average": 1}
    )

    movies = list(results)

    if movies:
        df = pd.DataFrame(movies)
        df_filtered = df[["title", "vote_average"]]

        writer = pd.ExcelWriter('movies.xlsx', engine='xlsxwriter')
        df_filtered.to_excel(writer, sheet_name='Movies', index=False)
        writer.save()

        print("Movies exported to movies.xlsx successfully.")
    else:
        print("No movies found with vote average greater than 9.0.")

if __name__ == '__main__':
    main()
