# -*- coding: utf-8 -*-

import smtplib
import os
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart


class Mail:
    def __init__(self):
        self.mail_host = "smtp.exmail.qq.com"
        self.mail_pass = "fgXqaq66pBu8VNKp"  # "9uNdzPUsfn3bGdMC"
        self.sender = "zengzhifuwu@fusionfintrade.com"
        self.to_receivers = []
        self.cc_receivers = []
        self.bcc_receivers = []
        self.subject = "测试"
        self.book = {
            "xuenze@fusionfintrade.com": "徐恩泽",
            "tanshuya@fusionfintrade.com": "谭舒雅",
            "pengcheng@fusionfintrade.com": "彭程",
            "wangyiwei@fusionfintrade.com": "王怿炜",
            "liminke@fusionfintrade.com": "李旻珂",
            "fangyupeng@fusionfintrade.com": "方雨鹏",
            "wenghuiying@fusionfintrade.com": "翁慧莹",
            "zhuyouzhi@fusionfintrade.com": "朱有志",
            "jiaoyanjia@fusionfintrade.com": "焦彦嘉",
            "shenzhihua@fusionfintrade.com": "申志华",
            "xiexiaokun@fusionfintrade.com": "谢晓坤",
            "gusongtao@fusionfintrade.com": "谷松涛",
            "zengzhifuwu@fusionfintrade.com": "增值服务部",
            "renchen@fusionfintrade.com": "任琛",
            "jinyanxu@fusionfintrade.com": "金延旭",
            "enzexu@qq.com": "徐恩泽的备份"
        }

    def format_address(self, addresses):
        formatted_address = []
        for address in addresses:
            if not self.book.get(address):
                name = address.split("@")[0]
            else:
                name = self.book.get(address)
            formatted_address.append(formataddr((Header(name, "utf-8").encode(), address)))
        return "; ".join(formatted_address)
        
    def set_receivers(self, to_receivers, cc_receivers, bcc_receivers):
        self.to_receivers = to_receivers
        self.cc_receivers = cc_receivers
        self.bcc_receivers = bcc_receivers

    def send(self, content, file_list, subject=None, content_type="plain"):
        if "[" in self.mail_pass:
            print("No Permission Warning: Skipped sending mail. Please contact Enze Xu for more information.")
            return False
        message = MIMEMultipart()
        message["From"] = formataddr((Header("增值服务部 [自动]", "utf-8").encode(), self.sender)) # self.format_address([self.sender])
        message["To"] = self.format_address(self.to_receivers)  # ",".join(self.to_receivers)
        message["Cc"] = self.format_address(self.cc_receivers)  # ",".join(self.cc_receivers)
        message["Bcc"] = self.format_address(self.bcc_receivers)  # ",".join(self.bcc_receivers)
        if subject is not None and subject != "":
            self.subject = subject
        message["Subject"] = Header(self.subject, "utf-8")
        message.attach(MIMEText(content, content_type, "utf-8"))
        try:
            for full_filename in file_list:
                file = MIMEApplication(open(full_filename, "rb").read())
                filename = full_filename.split(os.sep)[-1]
                file.add_header("Content-Disposition", "attachment", filename=filename)
                message.attach(file)
            smtp_obj = smtplib.SMTP_SSL(self.mail_host, 465)
            smtp_obj.login(self.sender, self.mail_pass)
            smtp_obj.sendmail(
                self.sender,
                self.to_receivers + self.cc_receivers + self.bcc_receivers,
                message.as_string()
            )
            smtp_obj.quit()
            return True
        except smtplib.SMTPException as err:
            print(err)
            return False


if __name__ == "__main__":
    mail = Mail()
    mail.set_receivers(["tanshuya@fusionfintrade.com", "enzexu@qq.com"], [], [])
    mail.send("test", [r"xxxxxx\1.txt"], "subject")
