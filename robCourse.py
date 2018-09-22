# -*- coding:utf8 -*-
# author: chengxi_zhang@pku.edu.cn
# 功能：实现选课网址的登录与选课人数监控
# 如果选课人数发生变化，在电脑端进行提醒，并发送通知邮件到指定邮箱
from selenium import webdriver
from time import sleep
import subprocess
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import ctypes


if __name__ == "__main__":
    # default-setting
    url = "http://elective.pku.edu.cn/"
    username = ""
    password = ""
    course_name = ["人工智能实践"]            # 如果要同时监控多个课程，在列表中以逗号分隔
    # -------------------- 如果使用windows系统，请取消下行的注释 --------------------
    # p = ctypes.windll.kernel32

    # 第三方SMTP服务，选课人数变化时进行邮件通知
    mail_host = "mail.pku.edu.cn"           # 邮件服务器
    mail_user = ""                          # 邮箱用户名
    mail_pass = ""                          # 密码
    sender = ""                             # 发送邮件的地址
    receivers = ['']                        # 接收邮件的邮件地址，如果要同时发送到多个邮箱，在列表中以逗号分隔
    message = MIMEText('选课人数发生变化', 'plain', 'utf-8')
    message['From'] = Header("Your name", 'utf-8')
    message['To'] = Header("Your name", 'utf-8')
    subject = '所选课程的选课人数发生变化，请注意抢课'
    message['Subject'] = Header(subject, 'utf-8')

    print "Program Starting..."
    # 启动模拟器，打开登录网址
    op = webdriver.ChromeOptions()
    # -------------------- 如果想在浏览器端查看页面情况，请取消下行的注释 --------------------
    # op.add_argument("headless")
    chrome = webdriver.Chrome(chrome_options=op)
    chrome.get(url)
    # 输入用户名与密码，并点击登录键
    print "Entering username and password..."
    chrome.find_element_by_id("user_name").send_keys(username)
    chrome.find_element_by_id("password").send_keys(password)
    chrome.find_element_by_id("logon_button").click()
    chrome.implicitly_wait(10)
    # 点击进入补退选界面
    print "Redirect to main page..."
    chrome.find_element_by_xpath(
        "//a[@href='/elective2008/edu/pku/stu/elective/controller/supplement/SupplyCancel.do']"
    ).click()
    # 循环查询所选课程的选课人数
    execute_times = 1
    while True:
        for each_course in course_name:
            content = chrome.find_element_by_xpath(
                "//span[contains(text(), '%s')]/following::td[11]/child::span[1]" % each_course
            ).text
            content = str(content).split(' ')
            print "当前课程", each_course, "刷新次数", execute_times, "选课情况", str(content[0])+" / "+str(content[2])

            if content[0] != content[2]:
                # mac版本弹窗提醒（无声音）
                # -------------------- 如果使用windows系统，请注释以下4行 --------------------
                title = course_name
                content = "选课人数发生变化"
                cmd = 'display notification "%s" with title "%s"' % (content, title)
                subprocess.call(["osascript", "-e", cmd])
                # window版本响铃提醒
                # -------------------- 如果使用windows系统，请取消下行的注释 --------------------
                # p.Beep(1000, 200)

                # 发送消息到指定邮箱
                try:
                    smtpObj = smtplib.SMTP()
                    smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
                    smtpObj.login(mail_user, mail_pass)
                    smtpObj.sendmail(sender, receivers, message.as_string())
                    print "通知邮件发送成功"
                except smtplib.SMTPException:
                    print "Error：通知邮件发送失败，请注意邮箱设置"
                # 结束本次抢课
                chrome.close()
                exit()

        # 如果选课人数没有变化，等待一段时间后刷新页面，重新获取人数
        execute_times += 1
        chrome.refresh()
        sleep(10)
