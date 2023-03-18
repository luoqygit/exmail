#encoding: utf-8

import argparse
import pandas as pd
import smtplib
import openpyxl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#from email.mime.application import MIMEApplication
from jinja2 import Template
import configparser


# 设置命令行参数
parser = argparse.ArgumentParser(description='发送邮件给Excel文件中的收件人。')
parser.add_argument('excel_file', help='包含收件人列表的Excel文件')
parser.add_argument('--temp_file', required=False, default='email_template.html', 
                    help='邮件模板文件，HTML格式，缺省为email_template.html')
parser.add_argument('--all', required=False, action="store_true", 
                    help='如果包含这个参数，发送邮件给excel列表中的所有人。否则只发送给列表中的第一个人，可用于测试配置文件和邮件模板。')
args = parser.parse_args()

# 是否发送邮件给所有人。
send_all = False
if args.all is True:
    print("Send to all receipients.")
else:
    print("send to the first person.")

exit()

# 读取 Excel 文件
try:
    df = pd.read_excel(args.excel_file)

    # 检查是否有发送状态列，如果没有，则增加状态列并把发送状态设置为"N"
    if "邮件已发送" not in df.columns:
        df["邮件已发送"]="N"

except (FileNotFoundError, openpyxl.utils.exceptions.InvalidFileException) as e:
    print('读取Excel文件出错: {e}')
    exit()

# 从配置文件中读取配置信息
try:
    with open('config.ini', encoding='UTF-8') as f:
        config = configparser.ConfigParser()
        config.read_file(f)

        # 读取 SMTP 服务器和认证信息
        smtp_server = config['SMTP']['smtp_server']
        smtp_port = int(config['SMTP']['smtp_port'])
        smtp_username = config['SMTP']['smtp_username']
        smtp_password = config['SMTP']['smtp_password']

        # 读取邮件信息
        subject = config['MAIL']['subject']
        email_col = config['MAIL']['email_col']
except IOError:
    print("读取config.ini出错。")
    exit()

# 打开邮件模板
try:
    with open(args.temp_file, 'r', encoding='UTF-8') as f:
        template = Template(f.read())
except IOError:
    print("读取{}出错。".format(args.temp_file))
    exit()

# 登录邮件服务器
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)

    # 循环遍历每个收件人
    for index, row in df.iterrows():

        # 如果邮件已经发过了，则跳过这条记录。
        if row["邮件已发送"] != "N":
            continue

        email_addr = row[email_col]

        # 配置邮件头信息
        message = MIMEMultipart()
        message['From'] = smtp_username
        message['To'] = email_addr
        message['Subject'] = subject

        # 生成邮件正文内容
        # 订单编号，姓名，手机号，邮箱，俱乐部名称，所属中区
        # html = template.render(order_no=row['订单编号'], name=row['姓名'], phone=row['手机号'], email=row['邮箱'], 
        #                         club=row['俱乐部名称'], zone=row['所属中区'])
        data = dict(row)
        html = template.render(data)

        # 将 HTML 正文作为邮件内容
        body = MIMEText(html, 'html')
        message.attach(body)

        # 发送邮件
        try:
            server.sendmail(smtp_username, email_addr, message.as_string())

            # 设置发送状态
            df.loc[index, "邮件已发送"] = "Y"
        except Exception as e:
            print("Error sending mail to {}".format(email_addr))

        if send_all is False:
            break

        if (index+1) % 50 == 0:
            print("{} records have been processed.".format(index+1))


# 保存发送结果
df.to_excel("results.xlsx", index=False)

print("总共 {} 个收件人，发送了 {} 封邮件。".format(len(df), len(df[df["邮件已发送"]=="Y"])))