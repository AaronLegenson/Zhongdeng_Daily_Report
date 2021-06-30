# -*- coding: utf-8 -*-

import schedule
from daily_job import daily_job
from tools import now_time_string
from notice import notice_hours_dev, notice_minutes_dev, notice_hours_prd, notice_minutes_prd


def start():
    """
    启动定时自动化部署，部署的时间放在notice.py中
    :return: None
    """
    print(now_time_string(), "[ ok ] Starting ...")
    for hour in notice_hours_dev:
        for minute in notice_minutes_dev:
            schedule.every().day.at("%02d:%02d" % (hour, minute)).do(daily_job, "dev")
    for hour in notice_hours_prd:
        for minute in notice_minutes_prd:
            schedule.every().day.at("%02d:%02d" % (hour, minute)).do(daily_job, "prd")
    print(now_time_string(), "[ ok ] Triggers dev mod at {0} : {1} every day".format(
        str(["%02d" % item for item in notice_hours_dev]).replace("'", ""),
        str(["%02d" % item for item in notice_minutes_dev]).replace("'", "")))
    print(now_time_string(), "[ ok ] Triggers prd mod at {0} : {1} every day".format(
        str(["%02d" % item for item in notice_hours_prd]).replace("'", ""),
        str(["%02d" % item for item in notice_minutes_prd]).replace("'", "")))
    while True:
        schedule.run_pending()


if __name__ == "__main__":
    start()
