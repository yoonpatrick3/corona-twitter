import twitter
import datetime
import schedule
import time
import xlsxwriter
import xlrd
import ast

#https://rapidapi.com/KishCom/api/covid-19-coronavirus-statistics?endpoint=apiendpoint_53587227-476d-4279-8f1d-4884e60d1db7

api_key = 'Mdzv5TYgJhVhCMIDoFNjG1q7d'
api_secret_key = '3cwTkKahKLK2wKkFuPYeaoUQBWG3Hmc4k7vLI06C629twr8bNH'
access_token = '801190254506082310-DRUdsOz16zOVKV9WTzU2NUkwLu6lCuT'
access_token_secret = 's8XsZ2T87vY3Uw6sH0762pk1dUPDyAW2jxjSq2633g7ev'

api = twitter.Api(consumer_key=api_key,
                  consumer_secret=api_secret_key,
                  access_token_key=access_token,
                  access_token_secret=access_token_secret)

d = datetime.datetime.now()

current_date = str(d.year) + "-" + str(d.month) + "-" + str(d.day)

def get_corona_trends(trends):
    keywords = ['corona', 'covid', 'quarantine', 'stayathomeandstaysafe', 'lockdown']

    arr = [current_date]
    total_volume = 0

    for word in keywords:
        for trend in trends:
            trend1 =  ast.literal_eval(str(trend).replace("\"", "'"))
            name = trend1['name']
            if word in name.lower() and 'tweet_volume' in trend1:
                arr.append(name + ": " + str(trend1['tweet_volume']))
                total_volume += trend1['tweet_volume']
            elif word in name.lower():
                arr.append(name + ": NED") #Not enough data from twitter for tweet volume
    arr.insert(1, total_volume)
    return arr

def job():

    data = get_corona_trends(api.GetTrendsCurrent())

    try:
        sheetname = 'C:/Users/12244/yoonp/independentCS/corona/' + d.strftime("%B") + '-Twitter-Corona.xlsx'
        book = xlrd.open_workbook(sheetname)
        sheet = book.sheet_by_index(0)
        next_row_num = len(sheet.col(0))

        workbook = xlsxwriter.Workbook(sheetname)
        worksheet = workbook.add_worksheet()
        workbook.close()

        for j in range(next_row_num):
            for i in range(len(sheet.row(i))):
                coord = xlrd.formula.cellname(j,i)
                worksheet.write(coord, sheet.row(j)[i])
        for i in range(len(data)):
            coord = xlrd.formula.cellname(next_row_num, i)
            worksheet.write(coord, data[i])

        print("Added entries")
    except:
        # creates new workbook for new month
        workbook = xlsxwriter.Workbook(sheetname)
        worksheet = workbook.add_worksheet()
        for i in range(len(data)):
            coord = xlrd.formula.cellname(0,i)
            worksheet.write(coord, data[i])
        workbook.close()

        print("Added entries")

schedule.every().day.at("01:00").do(job)
# calls job method everyday at 1
while True: # CHANGE TO TRUE WHEN U WANT TO RUn
    schedule.run_pending()
    time.sleep(60) # wait one minute
#nohup python cc.py &