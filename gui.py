import PySimpleGUI as psg
import openpyxl as op
import pandas as pd
# from IPython.display import HTML
import smtplib  # 发邮件
from email.mime.text import MIMEText  # 用于构建内容文本
from email.header import Header  # 用于构建邮件头
from datetime import datetime # 获取时间

# 备注信息
text = psg.popup_get_text("请输入邮件内容：")
if not text:
    psg.popup("操作中止！", "没有输入文字")
    raise SystemExit("Cancelling: No Text Entered")
else:
    psg.popup("您输入的内容为：", text)

    f1 = pd.read_excel("./Payroll.xlsx", engine='openpyxl', sheet_name='Sheet1')
    f2 = './Payroll.xlsx'
    wb = op.load_workbook(f2)
    ws = wb.worksheets[0]
    row_count = ws.max_row  # 表中的行数
    print(f'共计有{row_count - 1}条数据')

    #  遍历表中数据
    mail_list = None
    salary = None
    for n in range(0, row_count - 1):
        try:
            forms1 = f1.loc[n]
            salary = forms1.to_dict()  # 個人詳情
            mail_addr = salary['邮箱']  # 提取邮箱地址
            employee_email = str(mail_addr)  # 邮箱地址转换成字符串
            employee_name = salary['姓名']  # 提取员工姓名
            salary_month = salary['计薪月']  # 提取记薪月
            del salary['邮箱']  # 删除邮箱地址
            salary_description = str(salary)

            df = pd.DataFrame(forms1)
            # 获取时间
            dt01 = datetime.today()
            # print(dt01.date()) # 获取日期
            # print(dt01.time()) # 获取时间
            month = dt01.month-1 # 获取发薪月
            # 邮件尾部署名
            footer = f'<br>芜湖海立新能源-经营管理部 <br>{dt01.date()}'
            # 工资明细主体
            html = df.to_html(header=False)
            # 将h5生成到文件
            text_file = open("index.html", "w")
            text_file.write(html)
            text_file.close()
            # 服务器，端口
            host = 'smtp.qiye.163.com'
            port = 465
            # 我方账户，授权码
            username = 'payroll@highly-whnet.com'
            password = 'KzZmzsd6ayPb6UeT'
            # 对方账户
            to_addr = [employee_email]  # 添加多个账户采用列表形式
            # 要发送的内容
            moment = (f"Hi {employee_name} 您好！<br>&emsp;&emsp;以下为您 {month} 月工资明细，请查收~ 如有疑问请联系经营管理部-[董敏]。<br>" + f'备注: {text}' + html + footer)
            # 构建纯文本的邮件内容
            msg = MIMEText(moment, 'html', 'utf-8')
            # 构建邮件头
            msg['From'] = Header('payroll@highly-whnet.com')  # 发件人的名称或地址
            msg['To'] = Header(employee_email)  # to收件人邮箱地址
            msg['Subject'] = Header(f'{employee_name} {month} 月工资明细')  # 主题

            server = smtplib.SMTP_SSL(host)  # 开启发信服务
            server.connect(host, port)  # 连接发信服务
            server.login(username, password)  # 登录发信邮箱
            server.sendmail(username, to_addr, msg.as_string())  # 发送邮件
            server.quit()  # 关闭服务器
            print(f'[{employee_name}]{salary_month}月的薪资信息已发送至{mail_addr}')
            # psg.popup_notify(f'[{employee_name}]{salary_month}月的薪资信息已发送至{mail_addr}')
            with open('./log.txt', 'a', encoding='UTF-8') as f:
                f.write(f'{dt01.date()} + {dt01.time()}: [{employee_name}]{month}月的薪资信息已发送至{mail_addr}\n')
        except Exception as e:
            print(e)

psg.popup_notify(f'{row_count - 1}条数据均已发送，发送日志请查看log.txt文件')