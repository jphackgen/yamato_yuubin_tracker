import pandas as pd
import requests
import json

# If csv is detected, read & extract tracking id
try:
    df = open("~/Desktop/郵便追跡.csv").read().replace('"', '').replace("'", "")
    tracking_ids = [t for t in df.split('\n')]
    print("郵便追跡.csv を使います！")
except:
    print("郵便追跡.csv 見つかりませんでした、、")
    # If xlsx is detected, read & extract tracking id
    try:
        df = pd.read_excel(
            "~/Desktop/郵便追跡.xlsx",
            sheet_name=None,
            engine='openpyxl',
            header=None,
            index_col=None
        )
        print("郵便追跡.xlsx を使います！")
        df = df[list(df.keys())[0]][13]
        df = df.apply(lambda x: x.replace('"', ''))
        df = df.apply(lambda x: x.replace("'", ""))
        tracking_ids = df.values
    except:
        print("郵便追跡.xlsx 見つかりませんでした、、")

# Get data from API
status = []
print(f"{len(tracking_ids)}件読み込み完了！")
print("追跡開始！")
print("-----------------------------------------")
for i in tracking_ids:
    try:
        print(f"Tracking id: {i}", end="\t")
        r = requests.get('http://nanoappli.com/tracking/api/{}.json?_=1615855517639'.format(i))
        status_returned = json.loads(r.text)["status"]
        print(f"status: {status_returned}")
    except:
        print(" Failed!!")
        status_returned = ""
    finally:
        status.append(status_returned)

# Output to 郵便追跡_結果.csv
print("-----------------------------------------")
output_df = pd.DataFrame({"追跡ID": tracking_ids, "追跡結果": status})
output_df.to_csv("~/Desktop/郵便追跡_結果.csv", index=False)
print("追跡終了!「郵便追跡_結果.csv」をご確認ください！")
