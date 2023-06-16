from tempfile import template

import openpyxl as op
import pandas as pd
from IPython.display import HTML
import smtplib  # 发邮件
from email.mime.text import MIMEText  # 用于构建内容文本
from email.header import Header  # 用于构建邮件头

f1 = pd.read_excel("./Payroll.xlsx", engine='openpyxl', sheet_name='Sheet1')

f2 = './Payroll.xlsx'
wb = op.load_workbook(f2)
ws = wb.worksheets[0]
row_count = ws.max_row  # 表中的行数
print(f'共计有{row_count-1}条数据')

#  遍历表中数据
mail_list = None
salary = None
for n in range(0, row_count-1):
    try:
        forms1 = f1.loc[n]
        salary = forms1.to_dict()                       # 個人詳情
        mail_addr = salary['邮箱']                       # 提取邮箱地址
        employee_email = str(mail_addr)                 # 邮箱地址转换成字符串
        employee_name = salary['姓名']                   # 提取员工姓名
        salary_month = salary['计薪月']                  # 提取记薪月
        del salary['邮箱']                               # 删除邮箱地址
        salary_description = str(salary)

        df = pd.DataFrame(forms1)
        html = df.to_html(header=False)
        # 将h5生成到文件
        text_file = open("index.html", "w")
        text_file.write(html)
        text_file.close()
        # 服务器，端口
        host = 'smtp.qq.com'
        port = 465
        # 我方账户，授权码
        username = '283773655@qq.com'
        password = 'exgfxskqmqvtbjdi'
        # 对方账户
        to_addr = [employee_email]  # 添加多个账户采用列表形式
        # 要发送的内容
        moment = (f"Hi {employee_name} 您好！ 以下为 {salary_month} 月工资条" + html)
        # 构建纯文本的邮件内容
        msg = MIMEText(moment, 'html', 'utf-8')
        # 构建邮件头
        msg['From'] = Header('283773655@qq.com')  # 发件人的名称或地址
        msg['To'] = Header(employee_email)  # to收件人邮箱地址
        msg['Subject'] = Header(f'{employee_name}<{salary_month}>月工资表')  # 主题

        server = smtplib.SMTP_SSL(host)  # 开启发信服务
        server.connect(host, port)  # 连接发信服务
        server.login(username, password)  # 登录发信邮箱
        server.sendmail(username, to_addr, msg.as_string())  # 发送邮件
        server.quit()  # 关闭服务器
        print(f'[{employee_name}]{salary_month}月的薪资信息已发送至{mail_addr}')

    except Exception as e:
        print(e)