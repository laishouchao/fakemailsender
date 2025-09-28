from flask import Flask, render_template, request, jsonify
from flask_wtf import FlaskForm
from wtforms import StringField, TextAreaField, MultipleFileField
from wtforms.validators import DataRequired, Email
from flask_ckeditor import CKEditor, CKEditorField
import smtplib
import base64
import re
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
app.config['CKEDITOR_HEIGHT'] = 400
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB

ckeditor = CKEditor(app)

class EmailForm(FlaskForm):
    sender_name = StringField('发件人昵称', validators=[DataRequired()])
    sender_email = StringField('发件人邮箱', validators=[DataRequired(), Email()])
    recipients = StringField('收件人', validators=[DataRequired()])
    cc = StringField('抄送人')
    subject = StringField('邮件标题', validators=[DataRequired()])
    content = CKEditorField('邮件内容', validators=[DataRequired()])
    attachments = MultipleFileField('附件')

# ------------------- 配置区域 -------------------
POSTFIX_LOG_PATH = '/var/log/mail.log'  # Postfix日志路径（根据服务器调整）
CHECK_INTERVAL = 10  # 检查日志间隔（秒）
TIMEOUT = 300        # 超时时间（秒）

def extract_queue_id_from_postfix_log():
    """从 Postfix 日志中提取最新的队列ID"""
    try:
        with open(POSTFIX_LOG_PATH, "r") as log_file:
            lines = log_file.readlines()
            for line in reversed(lines):  # 从后往前读取，找到最新的队列ID
                match = re.search(r"postfix/[^:]+: (\w+):", line)
                if match:
                    return match.group(1)
    except Exception as e:
        print(f"解析 Postfix 日志时发生错误：{e}")
    return None

def check_delivery_status(queue_id):
    """检查Postfix日志中的投递状态，并提取详细信息"""
    status = None
    details = []
    try:
        with open(POSTFIX_LOG_PATH, 'r') as log_file:
            for line in log_file:
                if queue_id in line:
                    if "status=sent" in line:
                        status = "sent"
                        details.append("邮件已成功投递。")
                    elif "status=bounced" in line:
                        status = "bounced"
                        details.append("邮件投递失败（退回）：")
                    elif "status=deferred" in line:
                        status = "deferred"
                        details.append("邮件投递延迟：")
                    # 提取详细信息
                    if "status=" in line:
                        details.append(line.strip())
                    # 特别解析 DMARC 错误
                    if "DMARC check failed" in line:
                        details.append("\nDMARC 检查失败，可能的原因：")
                        details.append("- 发件域名的 DMARC 记录未正确配置。")
                        details.append("- SPF 或 DKIM 配置问题。")
                        details.append("- 发件服务器的 IP 地址被目标服务器限制。")
                        details.append("建议参考以下链接解决问题：")
                        details.append("https://open.work.weixin.qq.com/help2/pc/20049")
            if not status:
                status = "pending"
                details = ["邮件正在发送中..."]
        return status, "\n".join(details)
    except FileNotFoundError:
        print(f"错误：日志文件 {POSTFIX_LOG_PATH} 不存在")
        return "error", "无法访问邮件日志文件"
    except Exception as e:
        return "error", f"检查状态时发生错误：{str(e)}"

def send_email(form_data, files):
    try:
        msg = MIMEMultipart()
        
        # 设置发件人信息
        sender_name = form_data.get('sender_name')
        sender_email = form_data.get('sender_email')
        sendername = base64.b64encode(sender_name.encode('utf-8')).decode('utf-8')
        msg['From'] = f'=?utf-8?B?{sendername}?= <{sender_email}>'
        
        # 设置收件人和抄送人
        to_emails = [email.strip() for email in form_data.get('recipients').split(',')]
        cc_emails = [email.strip() for email in form_data.get('cc').split(',')] if form_data.get('cc') else []
        
        msg['To'] = ", ".join(to_emails)
        if cc_emails:
            msg['Cc'] = ", ".join(cc_emails)
        
        # 设置邮件主题和内容
        msg['Subject'] = Header(form_data.get('subject'), 'utf-8')
        msg.attach(MIMEText(form_data.get('content'), 'html', 'utf-8'))
        
        # 添加附件
        for file in files:
            if file.filename:
                part = MIMEApplication(file.read(), Name=file.filename)
                part['Content-Disposition'] = f'attachment; filename="{file.filename}"'
                msg.attach(part)
        
        # 发送邮件
        server = smtplib.SMTP('localhost')
        server.sendmail(sender_email, to_emails + cc_emails, msg.as_string())
        server.quit()
        
        # 获取队列ID和状态
        queue_id = extract_queue_id_from_postfix_log()
        if queue_id:
            status, details = check_delivery_status(queue_id)
            return {"status": status, "message": details, "queue_id": queue_id}
        
        return {"status": "success", "message": "邮件已发送成功"}
        
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.route('/')
def index():
    form = EmailForm()
    return render_template('index.html', form=form)

@app.route('/send', methods=['POST'])
def send():
    form_data = request.form.to_dict()
    files = request.files.getlist('attachments')
    result = send_email(form_data, files)
    return jsonify(result)

@app.route('/check_status')
def check_status():
    queue_id = request.args.get('queue_id')
    if not queue_id:
        return jsonify({"status": "error", "message": "未提供队列ID"})
    
    status, message = check_delivery_status(queue_id)
    return jsonify({"status": status, "message": message})

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)