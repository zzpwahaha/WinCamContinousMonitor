import datetime
import schedule
import time
print(datetime.datetime.now())

def job(dd):
    dd['count']+=1
    print("I'm working... at {}".format(datetime.datetime.now()))
    print(dd)
    print(time.time())
    
store = {'count': 0}
schedule.every(3).seconds.do(job,store)
while 1:
    schedule.run_pending()
    time.sleep(0.1)
    
    