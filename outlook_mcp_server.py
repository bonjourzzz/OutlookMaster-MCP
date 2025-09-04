import datetime
import os
import win32com.client
import json
import tempfile
import shutil
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP, Context

# Initialize FastMCP server
mcp = FastMCP("OutlookMaster-MCP")

# Constants
MAX_DAYS = 30
email_cache = {}
CACHE_FILE = os.path.join(tempfile.gettempdir(), "outlook_email_cache.json")

def save_email_cache(cache_data):
    """将邮件缓存保存到文件"""
    try:
        serializable_cache = {}
        for key, email in cache_data.items():
            email_copy = email.copy()
            email_copy['id'] = str(email_copy['id'])
            serializable_cache[str(key)] = email_copy
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(serializable_cache, f, ensure_ascii=False, default=str)
        return True
    except Exception as e:
        print(f"保存缓存出错: {str(e)}")
        return False

def load_email_cache():
    """从文件加载邮件缓存"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache = json.load(f)
                return {int(k): v for k, v in cache.items()}
        except Exception as e:
            print(f"加载缓存出错: {str(e)}")
            return {}
    return {}

def clear_email_cache():
    """清空邮件缓存"""
    global email_cache
    email_cache = {}
    if os.path.exists(CACHE_FILE):
        try:
            os.remove(CACHE_FILE)
        except Exception as e:
            print(f"清除缓存文件失败: {str(e)}")

def connect_to_outlook():
    """连接到Outlook应用程序"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return outlook, namespace
    except Exception as e:
        raise Exception(f"连接Outlook失败：{str(e)}")

def get_folder_by_name(namespace, folder_name: str):
    """根据名称获取特定的Outlook文件夹，如果不存在则创建"""
    try:
        # 检查默认文件夹
        default_folders = {
            "收件箱": 6, "已发送邮件": 5, "草稿": 16, 
            "已删除邮件": 3, "垃圾邮件": 18
        }
        
        if folder_name in default_folders:
            return namespace.GetDefaultFolder(default_folders[folder_name])
        
        # 搜索现有文件夹
        inbox = namespace.GetDefaultFolder(6)
        for folder in inbox.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
        for folder in namespace.Folders:
            if folder.Name.lower() == folder_name.lower():
                return folder
            for subfolder in folder.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
        
        # 如果找不到，在收件箱下创建新文件夹
        try:
            new_folder = inbox.Folders.Add(folder_name)
            return new_folder
        except Exception:
            return None
            
    except Exception as e:
        raise Exception(f"访问文件夹 {folder_name} 失败：{str(e)}")

def format_email(mail_item) -> Dict[str, Any]:
    """将Outlook邮件项格式化为结构化字典"""
    recipients = []
    try:
        if hasattr(mail_item, 'Recipients') and mail_item.Recipients:
            for i in range(1, mail_item.Recipients.Count + 1):
                recipient = mail_item.Recipients(i)
                try:
                    recipients.append(f"{recipient.Name} <{recipient.Address}>")
                except Exception:
                    recipients.append(f"{recipient.Name}")
    except Exception:
        pass
    
    return {
        "id": getattr(mail_item, "EntryID", ""),
        "conversation_id": getattr(mail_item, "ConversationID", None),
        "subject": getattr(mail_item, "Subject", "无主题"),
        "sender": getattr(mail_item, "SenderName", "未知发件人"),
        "sender_email": getattr(mail_item, "SenderEmailAddress", ""),
        "received_time": mail_item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S") if hasattr(mail_item, 'ReceivedTime') and mail_item.ReceivedTime else None,
        "recipients": recipients,
        "body": getattr(mail_item, "Body", ""),
        "has_attachments": hasattr(mail_item, 'Attachments') and mail_item.Attachments.Count > 0,
        "attachment_count": mail_item.Attachments.Count if hasattr(mail_item, 'Attachments') else 0,
        "unread": getattr(mail_item, "UnRead", False),
        "importance": getattr(mail_item, "Importance", 1),
        "categories": getattr(mail_item, "Categories", "")
    }

def get_emails_from_folder(folder, days: int, search_term: Optional[str] = None):
    """从文件夹获取邮件，支持可选的搜索过滤器"""
    emails_list = []
    now = datetime.datetime.now()
    threshold_date = now - datetime.timedelta(days=days)
    folder_items = folder.Items
    folder_items.Sort("[ReceivedTime]", True)

    for item in folder_items:
        try:
            if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                if item.ReceivedTime.replace(tzinfo=None) < threshold_date:
                    continue
                email_data = format_email(item)
                emails_list.append(email_data)
        except Exception as e:
            continue
    return emails_list

# ===== 基础邮件操作 =====
@mcp.tool()
def list_folders() -> str:
    """列出Outlook中所有可用的邮件文件夹"""
    try:
        _, namespace = connect_to_outlook()
        folders_info = []
        
        default_folders = {
            3: "已删除邮件", 4: "发件箱", 5: "已发送邮件", 6: "收件箱",
            9: "日历", 10: "联系人", 11: "日记", 12: "便笺", 13: "任务", 16: "草稿", 18: "垃圾邮件"
        }
        
        for folder_id, folder_name in default_folders.items():
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                if folder:
                    folders_info.append(f"- {folder_name} ({folder.Items.Count} 封邮件)")
            except Exception:
                continue
        
        for folder in namespace.Folders:
            try:
                folders_info.append(f"- {folder.Name} ({folder.Items.Count} 封邮件)")
                for subfolder in folder.Folders:
                    folders_info.append(f"  - {subfolder.Name} ({subfolder.Items.Count} 封邮件)")
            except Exception:
                continue
                
        return "可用的Outlook文件夹：\n" + "\n".join(folders_info)
    except Exception as e:
        return f"列出文件夹时出错：{str(e)}"

@mcp.tool()
def list_recent_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """列出最近几天的邮件"""
    if not isinstance(days, int) or days < 1 or days > MAX_DAYS:
        return f"错误：'days'必须是1到{MAX_DAYS}之间的整数"
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        if not folder:
            return f"错误：找不到文件夹'{folder_name}'"
        clear_email_cache()
        emails = get_emails_from_folder(folder, days)
        if not emails:
            return f"在{folder_name or '收件箱'}中没有找到最近{days}天的邮件。"
        
        global email_cache
        result = f"找到{len(emails)}封邮件：\n\n"
        for i, email in enumerate(emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']} <{email['sender_email']}>\n接收时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"获取邮件时出错：{str(e)}"

@mcp.tool()
def get_email_by_number(email_number: int) -> str:
    """获取指定邮件的完整内容"""
    try:
        global email_cache
        loaded_cache = load_email_cache()
        
        if not email_cache and loaded_cache:
            email_cache = loaded_cache
            
        if not email_cache:
            return "错误：还没有列出任何邮件。请先列出邮件。"
            
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}。"
            
        email_data = email_cache[email_number]
        _, namespace = connect_to_outlook()
        
        try:
            email = namespace.GetItemFromID(email_data["id"])
        except Exception as e:
            return f"错误：无法获取邮件。该邮件可能已被移动或删除。错误：{str(e)}"
            
        if not email:
            return f"错误：无法获取邮件 #{email_number}。"

        result = f"邮件 #{email_number} 详情：\n"
        result += f"主题：{email.Subject}\n"
        result += f"发件人：{email.SenderName} <{email.SenderEmailAddress}>\n"
        result += f"接收时间：{email.ReceivedTime}\n"
        
        recipients = email_data.get('recipients', [])
        result += f"收件人：{', '.join(recipients)}\n"
        
        if hasattr(email, 'Attachments') and email.Attachments.Count > 0:
            result += "附件：\n"
            for i in range(1, email.Attachments.Count + 1):
                try:
                    result += f" - {email.Attachments(i).FileName}\n"
                except Exception:
                    result += f" - [附件 {i}]\n"
                    
        result += "\n正文：\n"
        try:
            body = email.Body or "[未找到纯文本正文]"
            result += body
        except Exception as e:
            result += f"[获取邮件正文失败：{str(e)}]"
            
        return result
    except Exception as e:
        return f"获取邮件详情时出错：{str(e)}"

@mcp.tool()
def compose_email(to: str, subject: str, body: str, cc: Optional[str] = None, bcc: Optional[str] = None) -> str:
    """创建并发送新邮件"""
    if not to.strip():
        return "错误：'收件人'字段不能为空"
    if not subject.strip():
        return "错误：'主题'字段不能为空"
    if not body.strip():
        return "错误：'正文'字段不能为空"
        
    try:
        outlook, _ = connect_to_outlook()
        mail = outlook.CreateItem(0)
        
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        
        if cc and cc.strip():
            mail.CC = cc
        if bcc and bcc.strip():
            mail.BCC = bcc
            
        mail.Send()
        return f"邮件已成功发送给 {to}，主题为 '{subject}'"
    except Exception as e:
        return f"发送邮件时出错：{str(e)}"

@mcp.tool()
def reply_to_email_by_number(email_number: int, reply_body: str, reply_all: bool = False) -> str:
    """回复指定的邮件"""
    if not reply_body.strip():
        return "错误：回复内容不能为空"
        
    try:
        global email_cache
        loaded_cache = load_email_cache()
        
        if not email_cache and loaded_cache:
            email_cache = loaded_cache
            
        if not email_cache:
            return "错误：还没有列出任何邮件。请先列出邮件。"
            
        if email_number not in email_cache:
            return f"错误：在缓存中找不到邮件 #{email_number}。"
            
        email_data = email_cache[email_number]
        _, namespace = connect_to_outlook()
        
        try:
            original_email = namespace.GetItemFromID(email_data["id"])
        except Exception as e:
            return f"错误：无法获取原始邮件。该邮件可能已被移动或删除。错误：{str(e)}"
            
        if not original_email:
            return f"错误：无法获取邮件 #{email_number}。"
        
        if reply_all:
            reply = original_email.ReplyAll()
        else:
            reply = original_email.Reply()
            
        reply.Body = reply_body + "\n\n" + reply.Body
        reply.Send()
        
        action = "全部回复" if reply_all else "回复"
        return f"{action}已成功发送到邮件 #{email_number}（主题：{original_email.Subject}）"
    except Exception as e:
        return f"回复邮件时出错：{str(e)}"

# ===== 搜索功能 =====
@mcp.tool()
def search_emails(search_term: str, days: int = 7, folder_name: Optional[str] = None) -> str:
    """通过联系人姓名、关键词或短语搜索邮件，支持OR操作符"""
    if not search_term.strip():
        return "错误：搜索词不能为空"
    
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        if not folder:
            return f"错误：找不到文件夹'{folder_name}'"
            
        clear_email_cache()
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        search_terms = [term.strip().lower() for term in search_term.split(" OR ")]
        
        matching_emails = []
        folder_items = folder.Items
        folder_items.Sort("[ReceivedTime]", True)
        
        for item in folder_items:
            try:
                if not hasattr(item, "ReceivedTime") or not item.ReceivedTime:
                    continue
                if item.ReceivedTime.replace(tzinfo=None) < threshold_date:
                    continue
                    
                email_text = f"{item.Subject} {item.SenderName} {item.Body}".lower()
                if any(term in email_text for term in search_terms):
                    email_data = format_email(item)
                    matching_emails.append(email_data)
            except Exception:
                continue
        
        if not matching_emails:
            return f"在{folder_name or '收件箱'}中没有找到匹配'{search_term}'的邮件（最近{days}天）。"
        
        global email_cache
        result = f"找到{len(matching_emails)}封匹配'{search_term}'的邮件：\n\n"
        for i, email in enumerate(matching_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']} <{email['sender_email']}>\n接收时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"搜索邮件时出错：{str(e)}"

@mcp.tool()
def list_and_get_email(days: int = 7, folder_name: Optional[str] = None, email_number: Optional[int] = None) -> str:
    """列出邮件并可选获取特定邮件的内容"""
    # 先列出邮件
    result = list_recent_emails(days, folder_name)
    # 如果指定了邮件编号，直接返回邮件内容
    if email_number is not None:
        return get_email_by_number(email_number)
    return result

@mcp.tool()
def search_by_date_range(start_date: str, end_date: str, folder_name: Optional[str] = None) -> str:
    """按日期范围搜索邮件 (格式: YYYY-MM-DD)"""
    try:
        start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d")
        
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        clear_email_cache()
        matching_emails = []
        
        for item in folder.Items:
            try:
                if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                    received_dt = item.ReceivedTime.replace(tzinfo=None)
                    if start_dt <= received_dt <= end_dt:
                        matching_emails.append(format_email(item))
            except Exception:
                continue
        
        if not matching_emails:
            return f"在{start_date}到{end_date}期间没有找到邮件"
        
        global email_cache
        result = f"找到{len(matching_emails)}封邮件（{start_date} 到 {end_date}）：\n\n"
        for i, email in enumerate(matching_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"按日期搜索时出错：{str(e)}"

@mcp.tool()
def search_unread_emails(days: int = 7, folder_name: Optional[str] = None) -> str:
    """只搜索未读邮件"""
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        clear_email_cache()
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        unread_emails = []
        
        for item in folder.Items:
            try:
                if (hasattr(item, "UnRead") and item.UnRead and 
                    hasattr(item, "ReceivedTime") and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    unread_emails.append(format_email(item))
            except Exception:
                continue
        
        if not unread_emails:
            return f"最近{days}天没有未读邮件"
        
        global email_cache
        result = f"找到{len(unread_emails)}封未读邮件：\n\n"
        for i, email in enumerate(unread_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"搜索未读邮件时出错：{str(e)}"

@mcp.tool()
def search_with_attachments(days: int = 7, folder_name: Optional[str] = None) -> str:
    """只搜索有附件的邮件"""
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        clear_email_cache()
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        attachment_emails = []
        
        for item in folder.Items:
            try:
                if (hasattr(item, "Attachments") and item.Attachments.Count > 0 and
                    hasattr(item, "ReceivedTime") and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    attachment_emails.append(format_email(item))
            except Exception:
                continue
        
        if not attachment_emails:
            return f"最近{days}天没有带附件的邮件"
        
        global email_cache
        result = f"找到{len(attachment_emails)}封带附件的邮件：\n\n"
        for i, email in enumerate(attachment_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n附件数：{email['attachment_count']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"搜索带附件邮件时出错：{str(e)}"

@mcp.tool()
def search_by_importance(importance_level: str = "高", days: int = 7, folder_name: Optional[str] = None) -> str:
    """按重要性搜索邮件 (高/中/低)"""
    try:
        importance_map = {"高": 2, "中": 1, "低": 0}
        if importance_level not in importance_map:
            return "错误：重要性级别必须是'高'、'中'或'低'"
        
        target_importance = importance_map[importance_level]
        
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        clear_email_cache()
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        important_emails = []
        
        for item in folder.Items:
            try:
                if (hasattr(item, "Importance") and item.Importance == target_importance and
                    hasattr(item, "ReceivedTime") and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    important_emails.append(format_email(item))
            except Exception:
                continue
        
        if not important_emails:
            return f"最近{days}天没有{importance_level}重要性的邮件"
        
        global email_cache
        result = f"找到{len(important_emails)}封{importance_level}重要性邮件：\n\n"
        for i, email in enumerate(important_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"按重要性搜索时出错：{str(e)}"

# ===== 邮件管理功能 =====
@mcp.tool()
def mark_email_as_read(email_number: int, mark_read: bool = True) -> str:
    """标记邮件为已读或未读"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        email.UnRead = not mark_read
        email.Save()
        
        status = "已读" if mark_read else "未读"
        return f"邮件 #{email_number} 已标记为{status}"
    except Exception as e:
        return f"标记邮件状态时出错：{str(e)}"

@mcp.tool()
def delete_email_by_number(email_number: int) -> str:
    """删除指定邮件"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        subject = email.Subject
        email.Delete()
        
        return f"邮件 #{email_number} '{subject}' 已删除"
    except Exception as e:
        return f"删除邮件时出错：{str(e)}"

@mcp.tool()
def move_email_to_folder(email_number: int, target_folder: str) -> str:
    """移动邮件到指定文件夹"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        folder = get_folder_by_name(namespace, target_folder)
        if not folder:
            return f"错误：找不到文件夹 '{target_folder}'"
        
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        subject = email.Subject
        email.Move(folder)
        
        return f"邮件 #{email_number} '{subject}' 已移动到 '{target_folder}'"
    except Exception as e:
        return f"移动邮件时出错：{str(e)}"

@mcp.tool()
def flag_email(email_number: int, flag_status: str = "重要") -> str:
    """标记邮件为重要或跟进"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        if flag_status == "重要":
            email.Importance = 2  # High importance
        elif flag_status == "跟进":
            email.FlagStatus = 2  # Flagged
        
        email.Save()
        return f"邮件 #{email_number} 已标记为{flag_status}"
    except Exception as e:
        return f"标记邮件时出错：{str(e)}"

@mcp.tool()
def get_folder_summary() -> str:
    """获取所有文件夹摘要信息"""
    try:
        _, namespace = connect_to_outlook()
        
        default_folders = {
            6: "收件箱",
            5: "已发送邮件", 
            16: "草稿",
            3: "已删除邮件",
            18: "垃圾邮件"
        }
        
        result = "📁 文件夹摘要：\n\n"
        
        for folder_id, folder_name in default_folders.items():
            try:
                folder = namespace.GetDefaultFolder(folder_id)
                total = folder.Items.Count
                unread = sum(1 for item in folder.Items if hasattr(item, "UnRead") and item.UnRead)
                result += f"{folder_name}：{total} 封邮件（{unread} 封未读）\n"
            except Exception:
                continue
        
        # 自定义文件夹
        for folder in namespace.Folders:
            try:
                if folder.Name not in ["收件箱", "已发送邮件", "草稿", "已删除邮件", "垃圾邮件"]:
                    total = folder.Items.Count
                    unread = sum(1 for item in folder.Items if hasattr(item, "UnRead") and item.UnRead)
                    result += f"{folder.Name}：{total} 封邮件（{unread} 封未读）\n"
            except Exception:
                continue
        
        return result
    except Exception as e:
        return f"获取文件夹摘要时出错：{str(e)}"

@mcp.tool()
def get_sender_statistics(days: int = 30, top_count: int = 10) -> str:
    """获取发件人统计"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        sender_count = {}
        total_emails = 0
        
        for item in inbox.Items:
            try:
                if (hasattr(item, 'ReceivedTime') and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    sender = getattr(item, 'SenderName', '未知发件人')
                    sender_count[sender] = sender_count.get(sender, 0) + 1
                    total_emails += 1
            except Exception:
                continue
        
        if not sender_count:
            return f"最近{days}天没有邮件"
        
        # 排序并取前N个
        top_senders = sorted(sender_count.items(), key=lambda x: x[1], reverse=True)[:top_count]
        
        result = f"最近{days}天发件人统计（总邮件{total_emails}封）：\n\n"
        for i, (sender, count) in enumerate(top_senders, 1):
            percentage = (count / total_emails) * 100
            result += f"#{i} {sender}：{count}封 ({percentage:.1f}%)\n"
        
        return result
    except Exception as e:
        return f"获取发件人统计时出错：{str(e)}"

# ===== 附件管理功能 =====
@mcp.tool()
def download_attachment(email_number: int, attachment_name: Optional[str] = None, save_path: Optional[str] = None) -> str:
    """下载邮件附件"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        if email.Attachments.Count == 0:
            return f"邮件 #{email_number} 没有附件"
        
        if not save_path:
            save_path = os.path.join(os.getcwd(), "attachments")
            os.makedirs(save_path, exist_ok=True)
        
        downloaded = []
        for i in range(1, email.Attachments.Count + 1):
            attachment = email.Attachments(i)
            if not attachment_name or attachment_name in attachment.FileName:
                file_path = os.path.join(save_path, attachment.FileName)
                attachment.SaveAsFile(file_path)
                downloaded.append(attachment.FileName)
        
        if downloaded:
            return f"已下载附件：{', '.join(downloaded)} 到 {save_path}"
        else:
            return f"未找到匹配的附件：{attachment_name}"
    except Exception as e:
        return f"下载附件时出错：{str(e)}"

@mcp.tool()
def get_attachment_info(email_number: int) -> str:
    """获取邮件附件详细信息"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        if email.Attachments.Count == 0:
            return f"邮件 #{email_number} 没有附件"
        
        result = f"邮件 #{email_number} 附件信息：\n\n"
        total_size = 0
        
        for i in range(1, email.Attachments.Count + 1):
            attachment = email.Attachments(i)
            size_kb = attachment.Size / 1024
            total_size += attachment.Size
            
            result += f"附件 #{i}\n"
            result += f"文件名：{attachment.FileName}\n"
            result += f"大小：{size_kb:.2f} KB\n"
            result += f"类型：{attachment.Type}\n\n"
        
        result += f"总大小：{total_size/1024:.2f} KB"
        return result
    except Exception as e:
        return f"获取附件信息时出错：{str(e)}"

@mcp.tool()
def list_attachments_only(days: int = 7, folder_name: Optional[str] = None) -> str:
    """只列出有附件的邮件"""
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        clear_email_cache()
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        attachment_emails = []
        
        for item in folder.Items:
            try:
                if (hasattr(item, "Attachments") and item.Attachments.Count > 0 and
                    hasattr(item, "ReceivedTime") and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    attachment_emails.append(format_email(item))
            except Exception:
                continue
        
        if not attachment_emails:
            return f"最近{days}天没有带附件的邮件"
        
        global email_cache
        result = f"找到{len(attachment_emails)}封带附件的邮件：\n\n"
        for i, email in enumerate(attachment_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n附件数：{email['attachment_count']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"列出带附件邮件时出错：{str(e)}"

# ===== 批量操作功能 =====
@mcp.tool()
def mark_multiple_emails(email_numbers: str, mark_read: bool = True) -> str:
    """批量标记多封邮件为已读或未读"""
    try:
        numbers = [int(x.strip()) for x in email_numbers.split(",")]
        results = []
        
        for num in numbers:
            result = mark_email_as_read(num, mark_read)
            results.append(f"邮件 #{num}: {result}")
        
        return "批量操作结果：\n" + "\n".join(results)
    except Exception as e:
        return f"批量标记邮件时出错：{str(e)}"

@mcp.tool()
def delete_multiple_emails(email_numbers: str) -> str:
    """批量删除多封邮件"""
    try:
        numbers = [int(x.strip()) for x in email_numbers.split(",")]
        results = []
        
        for num in numbers:
            result = delete_email_by_number(num)
            results.append(f"邮件 #{num}: {result}")
        
        return "批量删除结果：\n" + "\n".join(results)
    except Exception as e:
        return f"批量删除邮件时出错：{str(e)}"

@mcp.tool()
def export_emails_to_file(days: int = 7, folder_name: Optional[str] = None, file_path: Optional[str] = None) -> str:
    """导出邮件到文件"""
    try:
        if not file_path:
            file_path = f"emails_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        emails = get_emails_from_folder(folder, days)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"邮件导出报告 - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"文件夹：{folder_name or '收件箱'}\n")
            f.write(f"时间范围：最近{days}天\n")
            f.write(f"邮件数量：{len(emails)}\n\n")
            
            for i, email in enumerate(emails, 1):
                f.write(f"=== 邮件 #{i} ===\n")
                f.write(f"主题：{email['subject']}\n")
                f.write(f"发件人：{email['sender']}\n")
                f.write(f"时间：{email['received_time']}\n")
                f.write(f"正文：{email['body'][:200]}...\n\n")
        
        return f"已导出{len(emails)}封邮件到文件：{file_path}"
    except Exception as e:
        return f"导出邮件时出错：{str(e)}"

@mcp.tool()
def check_folder_exists(folder_name: str) -> str:
    """检查文件夹是否存在"""
    try:
        _, namespace = connect_to_outlook()
        folder = get_folder_by_name(namespace, folder_name)
        
        if folder:
            return f"文件夹 '{folder_name}' 存在，包含 {folder.Items.Count} 封邮件"
        else:
            return f"文件夹 '{folder_name}' 不存在"
    except Exception as e:
        return f"检查文件夹时出错：{str(e)}"

@mcp.tool()
def create_simple_rule(rule_name: str, condition_type: str, condition_value: str, 
                      action_type: str, action_value: Optional[str] = None) -> str:
    """创建简单邮箱规则 (条件类型: 发件人/主题, 操作类型: 移动/标记/转发)"""
    try:
        outlook, namespace = connect_to_outlook()
        rules = outlook.Session.DefaultStore.GetRules()
        
        # 创建规则
        rule = rules.Create(rule_name, 0)
        
        # 设置条件
        if condition_type == "发件人":
            rule.Conditions.From.Enabled = True
            rule.Conditions.From.Recipients.Add(condition_value)
        elif condition_type == "主题":
            rule.Conditions.Subject.Enabled = True
            rule.Conditions.Subject.Text = [condition_value]
        else:
            return "错误：条件类型必须是'发件人'或'主题'"
        
        # 设置操作
        if action_type == "移动":
            if not action_value:
                return "错误：移动操作需要指定目标文件夹"
            folder = get_folder_by_name(namespace, action_value)
            if folder:
                rule.Actions.MoveToFolder.Enabled = True
                rule.Actions.MoveToFolder.Folder = folder
            else:
                return f"错误：无法访问文件夹 '{action_value}'"
        elif action_type == "标记":
            rule.Actions.MarkAsRead.Enabled = True
        elif action_type == "转发":
            if not action_value:
                return "错误：转发操作需要指定邮箱地址"
            rule.Actions.Forward.Enabled = True
            rule.Actions.Forward.Recipients.Add(action_value)
        else:
            return "错误：操作类型必须是'移动'、'标记'或'转发'"
        
        rule.Enabled = True
        rules.Save()
        
        return f"简单规则 '{rule_name}' 创建成功！"
        
    except Exception as e:
        return f"创建简单规则时出错：{str(e)}。建议使用Outlook手动创建复杂规则。"

# ===== 邮箱规则功能 =====
@mcp.tool()
def list_email_rules() -> str:
    """列出所有现有的邮箱规则"""
    try:
        outlook, namespace = connect_to_outlook()
        rules = outlook.Session.DefaultStore.GetRules()
        
        if rules.Count == 0:
            return "当前没有设置任何邮箱规则。"
        
        result = f"找到 {rules.Count} 条邮箱规则：\n\n"
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            status = "启用" if rule.Enabled else "禁用"
            result += f"规则 #{i}\n"
            result += f"名称：{rule.Name}\n"
            result += f"状态：{status}\n"
            result += f"执行顺序：{rule.ExecutionOrder}\n\n"
        
        return result
    except Exception as e:
        return f"获取邮箱规则时出错：{str(e)}"

@mcp.tool()
def create_email_rule(rule_name: str, sender_contains: Optional[str] = None, 
                     subject_contains: Optional[str] = None, move_to_folder: Optional[str] = None,
                     mark_as_read: bool = False, forward_to: Optional[str] = None) -> str:
    """创建新的邮箱规则"""
    if not rule_name.strip():
        return "错误：规则名称不能为空"
    
    if not any([sender_contains, subject_contains]):
        return "错误：必须指定至少一个条件（发件人包含 或 主题包含）"
    
    if not any([move_to_folder, mark_as_read, forward_to]):
        return "错误：必须指定至少一个操作（移动到文件夹、标记为已读 或 转发）"
    
    try:
        outlook, namespace = connect_to_outlook()
        rules = outlook.Session.DefaultStore.GetRules()
        
        # 创建新规则
        rule = rules.Create(rule_name, 0)  # 0 = olRuleReceive
        
        # 设置条件 - 简化条件设置
        conditions = rule.Conditions
        
        if sender_contains:
            try:
                conditions.From.Enabled = True
                conditions.From.Recipients.Add(sender_contains)
            except Exception:
                # 如果From不工作，尝试SenderAddress
                try:
                    conditions.SenderAddress.Enabled = True
                    conditions.SenderAddress.Address = [sender_contains]
                except Exception:
                    pass
        
        if subject_contains:
            try:
                conditions.Subject.Enabled = True
                conditions.Subject.Text = [subject_contains]
            except Exception:
                pass
        
        # 设置操作 - 简化操作设置
        actions = rule.Actions
        
        # 只设置一个主要操作以避免冲突
        if move_to_folder:
            target_folder = get_folder_by_name(namespace, move_to_folder)
            if target_folder:
                try:
                    actions.MoveToFolder.Enabled = True
                    actions.MoveToFolder.Folder = target_folder
                except Exception:
                    return f"错误：无法设置移动到文件夹 '{move_to_folder}'"
            else:
                return f"错误：无法创建或访问文件夹 '{move_to_folder}'"
        
        elif mark_as_read:
            try:
                # 修复MarkAsRead属性访问
                actions.MarkAsRead.Enabled = True
            except Exception:
                try:
                    # 尝试其他可能的属性名
                    actions.MarkRead.Enabled = True
                except Exception:
                    return "错误：无法设置标记为已读操作"
        
        elif forward_to:
            try:
                actions.Forward.Enabled = True
                actions.Forward.Recipients.Add(forward_to)
            except Exception:
                return f"错误：无法设置转发到 '{forward_to}'"
        
        # 启用并保存规则
        rule.Enabled = True
        rules.Save()
        
        return f"邮箱规则 '{rule_name}' 创建成功！"
        
    except Exception as e:
        return f"创建邮箱规则时出错：{str(e)}。建议手动在Outlook中创建规则。"

@mcp.tool()
def delete_email_rule(rule_name: str) -> str:
    """删除指定的邮箱规则"""
    if not rule_name.strip():
        return "错误：规则名称不能为空"
    
    try:
        outlook, namespace = connect_to_outlook()
        rules = outlook.Session.DefaultStore.GetRules()
        
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            if rule.Name.lower() == rule_name.lower():
                rules.Remove(i)
                rules.Save()
                return f"邮箱规则 '{rule_name}' 已成功删除！"
        
        return f"错误：找不到名为 '{rule_name}' 的规则"
        
    except Exception as e:
        return f"删除邮箱规则时出错：{str(e)}"

@mcp.tool()
def toggle_email_rule(rule_name: str, enable: bool = True) -> str:
    """启用或禁用指定的邮箱规则"""
    if not rule_name.strip():
        return "错误：规则名称不能为空"
    
    try:
        outlook, namespace = connect_to_outlook()
        rules = outlook.Session.DefaultStore.GetRules()
        
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            if rule.Name.lower() == rule_name.lower():
                rule.Enabled = enable
                rules.Save()
                status = "启用" if enable else "禁用"
                return f"邮箱规则 '{rule_name}' 已{status}！"
        
        return f"错误：找不到名为 '{rule_name}' 的规则"
        
    except Exception as e:
        return f"修改邮箱规则状态时出错：{str(e)}"

# ===== AI辅助功能 =====
@mcp.tool()
def summarize_email_thread(email_number: int) -> str:
    """总结邮件对话"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        # 简单的文本摘要（基于关键词和长度）
        body = email.Body
        sentences = body.split('。')
        
        # 提取关键信息
        keywords = ['会议', '项目', '截止', '完成', '需要', '请', '谢谢', '重要', '紧急']
        important_sentences = []
        
        for sentence in sentences[:10]:  # 只处理前10句
            if any(keyword in sentence for keyword in keywords) and len(sentence) > 10:
                important_sentences.append(sentence.strip())
        
        summary = f"邮件摘要：\n\n"
        summary += f"主题：{email.Subject}\n"
        summary += f"发件人：{email.SenderName}\n"
        summary += f"时间：{email.ReceivedTime.strftime('%Y-%m-%d %H:%M')}\n\n"
        
        if important_sentences:
            summary += "关键内容：\n"
            for i, sentence in enumerate(important_sentences[:3], 1):
                summary += f"{i}. {sentence}\n"
        else:
            summary += f"内容概要：{body[:200]}...\n"
        
        return summary
    except Exception as e:
        return f"总结邮件时出错：{str(e)}"

@mcp.tool()
def suggest_reply(email_number: int) -> str:
    """建议回复内容"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        body = email.Body.lower()
        subject = email.Subject.lower()
        
        # 基于关键词的回复建议
        suggestions = []
        
        if any(word in body for word in ['谢谢', '感谢']):
            suggestions.append("不客气，很高兴能帮助您。")
        
        if any(word in body for word in ['会议', '开会']):
            suggestions.append("我会准时参加会议。如有任何变更请及时通知。")
        
        if any(word in body for word in ['文件', '附件', '资料']):
            suggestions.append("我已收到文件，会仔细查看并尽快回复。")
        
        if any(word in body for word in ['截止', '期限', '时间']):
            suggestions.append("我了解时间要求，会按时完成并及时汇报进度。")
        
        if any(word in body for word in ['问题', '疑问', '咨询']):
            suggestions.append("关于您提到的问题，我需要进一步了解详情才能给出准确回复。")
        
        if not suggestions:
            suggestions = [
                "收到，我会尽快处理。",
                "谢谢您的邮件，我已了解相关情况。",
                "好的，如有问题我会及时联系您。"
            ]
        
        result = f"针对邮件 #{email_number} 的回复建议：\n\n"
        for i, suggestion in enumerate(suggestions[:3], 1):
            result += f"建议 {i}：{suggestion}\n\n"
        
        return result
    except Exception as e:
        return f"生成回复建议时出错：{str(e)}"

@mcp.tool()
def detect_email_sentiment(email_number: int) -> str:
    """检测邮件情感"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        text = f"{email.Subject} {email.Body}".lower()
        
        # 情感词典
        positive_words = ['谢谢', '感谢', '很好', '优秀', '满意', '高兴', '成功', '完成', '赞', '棒']
        negative_words = ['问题', '错误', '失败', '不满', '抱怨', '延迟', '困难', '紧急', '担心', '不行']
        neutral_words = ['通知', '会议', '文件', '资料', '时间', '地点', '联系', '确认', '安排']
        
        positive_count = sum(1 for word in positive_words if word in text)
        negative_count = sum(1 for word in negative_words if word in text)
        neutral_count = sum(1 for word in neutral_words if word in text)
        
        # 判断情感倾向
        if positive_count > negative_count and positive_count > 0:
            sentiment = "积极"
            confidence = min(90, 60 + positive_count * 10)
        elif negative_count > positive_count and negative_count > 0:
            sentiment = "消极"
            confidence = min(90, 60 + negative_count * 10)
        else:
            sentiment = "中性"
            confidence = 70
        
        # 紧急程度检测
        urgent_words = ['紧急', '立即', '马上', '尽快', '急']
        urgency = "高" if any(word in text for word in urgent_words) else "普通"
        
        result = f"邮件 #{email_number} 情感分析：\n\n"
        result += f"📧 主题：{email.Subject}\n"
        result += f"😊 情感倾向：{sentiment} (置信度: {confidence}%)\n"
        result += f"⚡ 紧急程度：{urgency}\n"
        result += f"📊 情感词统计：积极({positive_count}) 消极({negative_count}) 中性({neutral_count})\n"
        
        # 处理建议
        if sentiment == "消极":
            result += f"\n💡 建议：此邮件可能需要优先处理和谨慎回复"
        elif urgency == "高":
            result += f"\n💡 建议：此邮件标记为紧急，建议尽快回复"
        
        return result
    except Exception as e:
        return f"检测邮件情感时出错：{str(e)}"

@mcp.tool()
def auto_categorize_email(email_number: int) -> str:
    """自动分类邮件"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        text = f"{email.Subject} {email.Body}".lower()
        sender = email.SenderName.lower()
        
        # 分类规则
        categories = []
        
        # 工作相关
        if any(word in text for word in ['项目', '会议', '工作', '任务', '报告', '计划']):
            categories.append("工作")
        
        # 会议相关
        if any(word in text for word in ['会议', '开会', '讨论', '会面', '议程']):
            categories.append("会议")
        
        # 通知类
        if any(word in text for word in ['通知', '公告', '提醒', '更新', '变更']):
            categories.append("通知")
        
        # 个人相关
        if any(word in text for word in ['个人', '私人', '家庭', '朋友']):
            categories.append("个人")
        
        # 系统邮件
        if any(word in sender for word in ['noreply', 'system', 'admin', '系统']):
            categories.append("系统")
        
        # 营销邮件
        if any(word in text for word in ['优惠', '促销', '广告', '推广', '订阅']):
            categories.append("营销")
        
        # 紧急邮件
        if any(word in text for word in ['紧急', '立即', '马上', '重要']):
            categories.append("紧急")
        
        if not categories:
            categories = ["其他"]
        
        # 应用分类
        suggested_category = categories[0]
        current_categories = getattr(email, 'Categories', '')
        
        if current_categories:
            email.Categories = f"{current_categories}, {suggested_category}"
        else:
            email.Categories = suggested_category
        
        email.Save()
        
        result = f"邮件 #{email_number} 自动分类结果：\n\n"
        result += f"📧 主题：{email.Subject}\n"
        result += f"🏷️ 建议分类：{', '.join(categories)}\n"
        result += f"✅ 已应用分类：{suggested_category}\n"
        
        return result
    except Exception as e:
        return f"自动分类邮件时出错：{str(e)}"

# ===== 高级分析功能 =====
@mcp.tool()
def analyze_email_trends(days: int = 30) -> str:
    """分析邮件趋势"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        daily_count = {}
        hourly_count = {}
        total_emails = 0
        unread_count = 0
        
        for item in inbox.Items:
            try:
                if (hasattr(item, 'ReceivedTime') and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    
                    date_key = item.ReceivedTime.strftime("%Y-%m-%d")
                    hour_key = item.ReceivedTime.hour
                    
                    daily_count[date_key] = daily_count.get(date_key, 0) + 1
                    hourly_count[hour_key] = hourly_count.get(hour_key, 0) + 1
                    total_emails += 1
                    
                    if getattr(item, 'UnRead', False):
                        unread_count += 1
            except Exception:
                continue
        
        if not daily_count:
            return f"最近{days}天没有邮件数据"
        
        # 计算统计数据
        avg_daily = total_emails / len(daily_count)
        peak_hour = max(hourly_count.items(), key=lambda x: x[1]) if hourly_count else (0, 0)
        
        result = f"最近{days}天邮件趋势分析：\n\n"
        result += f"📊 总邮件数：{total_emails}封\n"
        result += f"📈 日均邮件：{avg_daily:.1f}封\n"
        result += f"🔵 未读邮件：{unread_count}封 ({(unread_count/total_emails*100):.1f}%)\n"
        result += f"⏰ 邮件高峰时段：{peak_hour[0]}:00-{peak_hour[0]+1}:00 ({peak_hour[1]}封)\n\n"
        
        result += "📅 最近7天邮件数量：\n"
        for date in sorted(daily_count.keys())[-7:]:
            result += f"{date}：{daily_count[date]}封\n"
        
        return result
    except Exception as e:
        return f"分析邮件趋势时出错：{str(e)}"

@mcp.tool()
def get_response_time_stats(days: int = 30) -> str:
    """获取回复时间统计"""
    try:
        _, namespace = connect_to_outlook()
        sent_folder = namespace.GetDefaultFolder(5)  # Sent Items
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        # 收集发送的邮件
        sent_emails = {}
        for item in sent_folder.Items:
            try:
                if (hasattr(item, 'SentOn') and item.SentOn and
                    item.SentOn.replace(tzinfo=None) >= threshold_date):
                    conversation_id = getattr(item, 'ConversationID', None)
                    if conversation_id:
                        sent_emails[conversation_id] = item.SentOn
            except Exception:
                continue
        
        # 计算回复时间
        response_times = []
        for item in inbox.Items:
            try:
                if (hasattr(item, 'ReceivedTime') and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    conversation_id = getattr(item, 'ConversationID', None)
                    if conversation_id in sent_emails:
                        time_diff = (sent_emails[conversation_id] - item.ReceivedTime).total_seconds() / 3600
                        if 0 < time_diff < 168:  # 1周内的回复
                            response_times.append(time_diff)
            except Exception:
                continue
        
        if not response_times:
            return f"最近{days}天没有回复时间数据"
        
        avg_response = sum(response_times) / len(response_times)
        min_response = min(response_times)
        max_response = max(response_times)
        
        # 分类统计
        quick_replies = len([t for t in response_times if t <= 1])  # 1小时内
        same_day = len([t for t in response_times if t <= 24])  # 24小时内
        
        result = f"最近{days}天回复时间统计：\n\n"
        result += f"📧 分析邮件数：{len(response_times)}封\n"
        result += f"⏱️ 平均回复时间：{avg_response:.1f}小时\n"
        result += f"🚀 最快回复：{min_response:.1f}小时\n"
        result += f"🐌 最慢回复：{max_response:.1f}小时\n"
        result += f"⚡ 1小时内回复：{quick_replies}封 ({quick_replies/len(response_times)*100:.1f}%)\n"
        result += f"📅 24小时内回复：{same_day}封 ({same_day/len(response_times)*100:.1f}%)\n"
        
        return result
    except Exception as e:
        return f"获取回复时间统计时出错：{str(e)}"

@mcp.tool()
def get_sender_statistics_advanced(days: int = 30, analysis_type: str = "详细") -> str:
    """高级发件人统计 (详细/简要)"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        sender_data = {}
        total_emails = 0
        
        for item in inbox.Items:
            try:
                if (hasattr(item, 'ReceivedTime') and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    
                    sender = getattr(item, 'SenderName', '未知发件人')
                    sender_email = getattr(item, 'SenderEmailAddress', '')
                    
                    if sender not in sender_data:
                        sender_data[sender] = {
                            'count': 0,
                            'email': sender_email,
                            'unread': 0,
                            'with_attachments': 0,
                            'high_importance': 0
                        }
                    
                    sender_data[sender]['count'] += 1
                    total_emails += 1
                    
                    if getattr(item, 'UnRead', False):
                        sender_data[sender]['unread'] += 1
                    
                    if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                        sender_data[sender]['with_attachments'] += 1
                    
                    if getattr(item, 'Importance', 1) == 2:  # High importance
                        sender_data[sender]['high_importance'] += 1
            except Exception:
                continue
        
        if not sender_data:
            return f"最近{days}天没有邮件数据"
        
        # 排序
        top_senders = sorted(sender_data.items(), key=lambda x: x[1]['count'], reverse=True)[:10]
        
        result = f"最近{days}天高级发件人统计（总邮件{total_emails}封）：\n\n"
        
        for i, (sender, data) in enumerate(top_senders, 1):
            percentage = (data['count'] / total_emails) * 100
            result += f"#{i} {sender}\n"
            result += f"   邮箱：{data['email']}\n"
            result += f"   邮件数：{data['count']}封 ({percentage:.1f}%)\n"
            
            if analysis_type == "详细":
                result += f"   未读：{data['unread']}封\n"
                result += f"   带附件：{data['with_attachments']}封\n"
                result += f"   高重要性：{data['high_importance']}封\n"
            
            result += "\n"
        
        return result
    except Exception as e:
        return f"获取高级发件人统计时出错：{str(e)}"

# ===== 邮件模板功能 =====
@mcp.tool()
def save_email_as_template(email_number: int, template_name: str) -> str:
    """保存邮件为模板"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        # 创建模板文件夹
        template_dir = os.path.join(os.getcwd(), "email_templates")
        os.makedirs(template_dir, exist_ok=True)
        
        template_data = {
            'name': template_name,
            'subject': email.Subject,
            'body': email.Body,
            'created_date': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        template_file = os.path.join(template_dir, f"{template_name}.json")
        with open(template_file, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=2)
        
        return f"邮件模板 '{template_name}' 保存成功"
    except Exception as e:
        return f"保存邮件模板时出错：{str(e)}"

@mcp.tool()
def list_email_templates() -> str:
    """列出邮件模板"""
    try:
        template_dir = os.path.join(os.getcwd(), "email_templates")
        if not os.path.exists(template_dir):
            return "没有找到邮件模板"
        
        templates = []
        for file in os.listdir(template_dir):
            if file.endswith('.json'):
                try:
                    with open(os.path.join(template_dir, file), 'r', encoding='utf-8') as f:
                        template_data = json.load(f)
                        templates.append(template_data)
                except Exception:
                    continue
        
        if not templates:
            return "没有可用的邮件模板"
        
        result = f"邮件模板列表（共{len(templates)}个）：\n\n"
        for i, template in enumerate(templates, 1):
            result += f"模板 #{i}\n"
            result += f"名称：{template['name']}\n"
            result += f"主题：{template['subject']}\n"
            result += f"创建时间：{template['created_date']}\n"
            result += f"内容预览：{template['body'][:100]}...\n\n"
        
        return result
    except Exception as e:
        return f"获取邮件模板时出错：{str(e)}"

@mcp.tool()
def compose_from_template(template_name: str, to: str, 
                         subject_override: Optional[str] = None,
                         body_additions: Optional[str] = None) -> str:
    """使用模板撰写邮件"""
    try:
        template_dir = os.path.join(os.getcwd(), "email_templates")
        template_file = os.path.join(template_dir, f"{template_name}.json")
        
        if not os.path.exists(template_file):
            return f"未找到模板：{template_name}"
        
        with open(template_file, 'r', encoding='utf-8') as f:
            template_data = json.load(f)
        
        outlook, _ = connect_to_outlook()
        mail = outlook.CreateItem(0)
        
        mail.To = to
        mail.Subject = subject_override or template_data['subject']
        
        body = template_data['body']
        if body_additions:
            body = f"{body_additions}\n\n{body}"
        
        mail.Body = body
        mail.Send()
        
        return f"使用模板 '{template_name}' 发送邮件成功到 {to}"
    except Exception as e:
        return f"使用模板撰写邮件时出错：{str(e)}"

# ===== 任务管理功能 =====
@mcp.tool()
def list_tasks(status: str = "全部") -> str:
    """列出任务 (全部/未完成/已完成)"""
    try:
        _, namespace = connect_to_outlook()
        tasks = namespace.GetDefaultFolder(13)  # 13 is Tasks
        
        task_list = []
        for item in tasks.Items:
            try:
                task_status = "已完成" if getattr(item, 'Complete', False) else "未完成"
                
                if status == "全部" or status == task_status:
                    task_list.append({
                        'subject': getattr(item, 'Subject', '无主题'),
                        'status': task_status,
                        'due_date': item.DueDate.strftime("%Y-%m-%d") if hasattr(item, 'DueDate') and item.DueDate else '无截止日期',
                        'priority': getattr(item, 'Importance', 1),
                        'percent_complete': getattr(item, 'PercentComplete', 0)
                    })
            except Exception:
                continue
        
        if not task_list:
            return f"没有{status}的任务"
        
        result = f"{status}任务列表（共{len(task_list)}个）：\n\n"
        for i, task in enumerate(task_list, 1):
            priority_text = {0: "低", 1: "普通", 2: "高"}.get(task['priority'], "普通")
            result += f"任务 #{i}\n"
            result += f"主题：{task['subject']}\n"
            result += f"状态：{task['status']}\n"
            result += f"截止日期：{task['due_date']}\n"
            result += f"优先级：{priority_text}\n"
            result += f"完成度：{task['percent_complete']}%\n\n"
        
        return result
    except Exception as e:
        return f"获取任务列表时出错：{str(e)}"

@mcp.tool()
def create_task_from_email(email_number: int, due_date: Optional[str] = None) -> str:
    """从邮件创建任务 (日期格式: YYYY-MM-DD)"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        outlook, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        task = outlook.CreateItem(3)  # 3 is olTaskItem
        task.Subject = f"处理邮件：{email.Subject}"
        task.Body = f"来自邮件：{email.SenderName}\n\n{email.Body[:500]}..."
        
        if due_date:
            task.DueDate = datetime.datetime.strptime(due_date, "%Y-%m-%d")
        
        task.Save()
        return f"已从邮件 #{email_number} 创建任务：{task.Subject}"
    except Exception as e:
        return f"从邮件创建任务时出错：{str(e)}"

@mcp.tool()
def mark_task_complete(task_subject: str) -> str:
    """标记任务完成"""
    try:
        _, namespace = connect_to_outlook()
        tasks = namespace.GetDefaultFolder(13)
        
        for item in tasks.Items:
            try:
                if task_subject.lower() in getattr(item, 'Subject', '').lower():
                    item.Complete = True
                    item.PercentComplete = 100
                    item.Save()
                    return f"任务 '{item.Subject}' 已标记为完成"
            except Exception:
                continue
        
        return f"未找到主题包含'{task_subject}'的任务"
    except Exception as e:
        return f"标记任务完成时出错：{str(e)}"

# ===== 邮件分类和标签功能 =====
@mcp.tool()
def add_category_to_email(email_number: int, category: str) -> str:
    """为邮件添加分类"""
    try:
        global email_cache
        if not email_cache:
            email_cache = load_email_cache()
        if email_number not in email_cache:
            return f"错误：找不到邮件 #{email_number}"
        
        _, namespace = connect_to_outlook()
        email = namespace.GetItemFromID(email_cache[email_number]["id"])
        
        current_categories = getattr(email, 'Categories', '')
        if current_categories:
            email.Categories = f"{current_categories}, {category}"
        else:
            email.Categories = category
        
        email.Save()
        return f"已为邮件 #{email_number} 添加分类：{category}"
    except Exception as e:
        return f"添加邮件分类时出错：{str(e)}"

@mcp.tool()
def list_email_categories() -> str:
    """列出所有邮件分类"""
    try:
        outlook, namespace = connect_to_outlook()
        categories = outlook.Session.Categories
        
        if categories.Count == 0:
            return "没有设置任何邮件分类"
        
        result = f"邮件分类列表（共{categories.Count}个）：\n\n"
        for i in range(1, categories.Count + 1):
            category = categories.Item(i)
            result += f"分类 #{i}\n"
            result += f"名称：{category.Name}\n"
            result += f"颜色：{category.Color}\n\n"
        
        return result
    except Exception as e:
        return f"获取邮件分类时出错：{str(e)}"

@mcp.tool()
def search_by_category(category: str, days: int = 30) -> str:
    """按分类搜索邮件"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        clear_email_cache()
        categorized_emails = []
        
        for item in inbox.Items:
            try:
                if (hasattr(item, 'Categories') and item.Categories and
                    category.lower() in item.Categories.lower() and
                    hasattr(item, 'ReceivedTime') and item.ReceivedTime and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    categorized_emails.append(format_email(item))
            except Exception:
                continue
        
        if not categorized_emails:
            return f"最近{days}天没有找到分类为'{category}'的邮件"
        
        global email_cache
        result = f"找到{len(categorized_emails)}封分类为'{category}'的邮件：\n\n"
        for i, email in enumerate(categorized_emails, 1):
            email_cache[i] = email
            result += f"邮件 #{i}\n主题：{email['subject']}\n发件人：{email['sender']}\n分类：{email['categories']}\n时间：{email['received_time']}\n\n"
        
        save_email_cache(email_cache)
        return result
    except Exception as e:
        return f"按分类搜索邮件时出错：{str(e)}"

# ===== 联系人管理功能 =====
@mcp.tool()
def list_contacts(limit: int = 50) -> str:
    """列出联系人"""
    try:
        _, namespace = connect_to_outlook()
        contacts = namespace.GetDefaultFolder(10)  # 10 is Contacts
        
        contact_list = []
        count = 0
        for item in contacts.Items:
            if count >= limit:
                break
            try:
                contact_list.append({
                    'name': getattr(item, 'FullName', '无姓名'),
                    'email': getattr(item, 'Email1Address', ''),
                    'company': getattr(item, 'CompanyName', ''),
                    'phone': getattr(item, 'BusinessTelephoneNumber', '')
                })
                count += 1
            except Exception:
                continue
        
        if not contact_list:
            return "联系人列表为空"
        
        result = f"联系人列表（前{len(contact_list)}个）：\n\n"
        for i, contact in enumerate(contact_list, 1):
            result += f"联系人 #{i}\n"
            result += f"姓名：{contact['name']}\n"
            result += f"邮箱：{contact['email']}\n"
            result += f"公司：{contact['company']}\n"
            result += f"电话：{contact['phone']}\n\n"
        
        return result
    except Exception as e:
        return f"获取联系人时出错：{str(e)}"

@mcp.tool()
def search_contacts(search_term: str) -> str:
    """搜索联系人"""
    try:
        _, namespace = connect_to_outlook()
        contacts = namespace.GetDefaultFolder(10)
        
        matching_contacts = []
        for item in contacts.Items:
            try:
                contact_text = f"{getattr(item, 'FullName', '')} {getattr(item, 'Email1Address', '')} {getattr(item, 'CompanyName', '')}".lower()
                if search_term.lower() in contact_text:
                    matching_contacts.append({
                        'name': getattr(item, 'FullName', '无姓名'),
                        'email': getattr(item, 'Email1Address', ''),
                        'company': getattr(item, 'CompanyName', ''),
                        'phone': getattr(item, 'BusinessTelephoneNumber', '')
                    })
            except Exception:
                continue
        
        if not matching_contacts:
            return f"未找到匹配'{search_term}'的联系人"
        
        result = f"找到{len(matching_contacts)}个匹配的联系人：\n\n"
        for i, contact in enumerate(matching_contacts, 1):
            result += f"联系人 #{i}\n"
            result += f"姓名：{contact['name']}\n"
            result += f"邮箱：{contact['email']}\n"
            result += f"公司：{contact['company']}\n"
            result += f"电话：{contact['phone']}\n\n"
        
        return result
    except Exception as e:
        return f"搜索联系人时出错：{str(e)}"

@mcp.tool()
def add_contact(name: str, email: str, company: Optional[str] = None, phone: Optional[str] = None) -> str:
    """添加新联系人"""
    try:
        outlook, _ = connect_to_outlook()
        contact = outlook.CreateItem(2)  # 2 is olContactItem
        
        contact.FullName = name
        contact.Email1Address = email
        
        if company:
            contact.CompanyName = company
        if phone:
            contact.BusinessTelephoneNumber = phone
        
        contact.Save()
        return f"联系人 '{name}' 添加成功"
    except Exception as e:
        return f"添加联系人时出错：{str(e)}"

@mcp.tool()
def get_contact_info(contact_name: str) -> str:
    """获取联系人详细信息"""
    try:
        _, namespace = connect_to_outlook()
        contacts = namespace.GetDefaultFolder(10)
        
        for item in contacts.Items:
            try:
                if contact_name.lower() in getattr(item, 'FullName', '').lower():
                    result = f"联系人详细信息：\n\n"
                    result += f"姓名：{getattr(item, 'FullName', '')}\n"
                    result += f"邮箱1：{getattr(item, 'Email1Address', '')}\n"
                    result += f"邮箱2：{getattr(item, 'Email2Address', '')}\n"
                    result += f"公司：{getattr(item, 'CompanyName', '')}\n"
                    result += f"职位：{getattr(item, 'JobTitle', '')}\n"
                    result += f"商务电话：{getattr(item, 'BusinessTelephoneNumber', '')}\n"
                    result += f"手机：{getattr(item, 'MobileTelephoneNumber', '')}\n"
                    result += f"地址：{getattr(item, 'BusinessAddress', '')}\n"
                    result += f"备注：{getattr(item, 'Body', '')}\n"
                    return result
            except Exception:
                continue
        
        return f"未找到联系人：{contact_name}"
    except Exception as e:
        return f"获取联系人信息时出错：{str(e)}"

# ===== 日历集成功能 =====
@mcp.tool()
def list_calendar_events(days: int = 7) -> str:
    """列出日历事件"""
    try:
        outlook, namespace = connect_to_outlook()
        calendar = namespace.GetDefaultFolder(9)  # 9 is Calendar
        
        now = datetime.datetime.now()
        end_date = now + datetime.timedelta(days=days)
        
        events = []
        for item in calendar.Items:
            try:
                if hasattr(item, 'Start') and item.Start:
                    if now <= item.Start.replace(tzinfo=None) <= end_date:
                        events.append({
                            'subject': getattr(item, 'Subject', '无主题'),
                            'start': item.Start.strftime("%Y-%m-%d %H:%M"),
                            'end': item.End.strftime("%Y-%m-%d %H:%M") if hasattr(item, 'End') else '',
                            'location': getattr(item, 'Location', ''),
                            'organizer': getattr(item, 'Organizer', '')
                        })
            except Exception:
                continue
        
        if not events:
            return f"未来{days}天没有日历事件"
        
        result = f"未来{days}天的日历事件：\n\n"
        for i, event in enumerate(events, 1):
            result += f"事件 #{i}\n"
            result += f"主题：{event['subject']}\n"
            result += f"开始：{event['start']}\n"
            result += f"结束：{event['end']}\n"
            result += f"地点：{event['location']}\n"
            result += f"组织者：{event['organizer']}\n\n"
        
        return result
    except Exception as e:
        return f"获取日历事件时出错：{str(e)}"

@mcp.tool()
def create_calendar_event(subject: str, start_time: str, end_time: str, 
                         location: Optional[str] = None, attendees: Optional[str] = None) -> str:
    """创建日历事件 (时间格式: YYYY-MM-DD HH:MM)"""
    try:
        outlook, _ = connect_to_outlook()
        appointment = outlook.CreateItem(1)  # 1 is olAppointmentItem
        
        appointment.Subject = subject
        appointment.Start = datetime.datetime.strptime(start_time, "%Y-%m-%d %H:%M")
        appointment.End = datetime.datetime.strptime(end_time, "%Y-%m-%d %H:%M")
        
        if location:
            appointment.Location = location
        
        if attendees:
            for email in attendees.split(','):
                appointment.Recipients.Add(email.strip())
        
        appointment.Save()
        return f"日历事件 '{subject}' 创建成功"
    except Exception as e:
        return f"创建日历事件时出错：{str(e)}"

@mcp.tool()
def get_meeting_invitations(days: int = 7) -> str:
    """获取会议邀请"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        now = datetime.datetime.now()
        threshold_date = now - datetime.timedelta(days=days)
        
        invitations = []
        for item in inbox.Items:
            try:
                if (hasattr(item, 'MessageClass') and 
                    'IPM.Schedule.Meeting' in item.MessageClass and
                    item.ReceivedTime.replace(tzinfo=None) >= threshold_date):
                    invitations.append({
                        'subject': item.Subject,
                        'sender': item.SenderName,
                        'received': item.ReceivedTime.strftime("%Y-%m-%d %H:%M")
                    })
            except Exception:
                continue
        
        if not invitations:
            return f"最近{days}天没有会议邀请"
        
        result = f"最近{days}天的会议邀请：\n\n"
        for i, inv in enumerate(invitations, 1):
            result += f"邀请 #{i}\n主题：{inv['subject']}\n发起人：{inv['sender']}\n时间：{inv['received']}\n\n"
        
        return result
    except Exception as e:
        return f"获取会议邀请时出错：{str(e)}"

@mcp.tool()
def respond_to_meeting(meeting_subject: str, response: str = "接受") -> str:
    """回复会议邀请 (接受/拒绝/暂定)"""
    try:
        _, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        
        response_map = {"接受": 3, "拒绝": 4, "暂定": 2}
        if response not in response_map:
            return "错误：回复必须是'接受'、'拒绝'或'暂定'"
        
        for item in inbox.Items:
            try:
                if (hasattr(item, 'MessageClass') and 
                    'IPM.Schedule.Meeting' in item.MessageClass and
                    meeting_subject.lower() in item.Subject.lower()):
                    
                    meeting_item = item.GetAssociatedAppointment(True)
                    meeting_item.Respond(response_map[response], True)
                    return f"已{response}会议邀请：{item.Subject}"
            except Exception:
                continue
        
        return f"未找到主题包含'{meeting_subject}'的会议邀请"
    except Exception as e:
        return f"回复会议邀请时出错：{str(e)}"

# ===== 统计功能 =====
@mcp.tool()
def get_email_statistics(folder_name: Optional[str] = None) -> str:
    """获取邮件统计信息"""
    try:
        _, namespace = connect_to_outlook()
        folder = namespace.GetDefaultFolder(6) if not folder_name else get_folder_by_name(namespace, folder_name)
        
        total_count = folder.Items.Count
        unread_count = 0
        today_count = 0
        attachment_count = 0
        
        today = datetime.datetime.now().date()
        
        for item in folder.Items:
            try:
                if hasattr(item, "UnRead") and item.UnRead:
                    unread_count += 1
                
                if hasattr(item, "ReceivedTime") and item.ReceivedTime:
                    if item.ReceivedTime.date() == today:
                        today_count += 1
                
                if hasattr(item, "Attachments") and item.Attachments.Count > 0:
                    attachment_count += 1
            except Exception:
                continue
        
        result = f"📊 {folder_name or '收件箱'} 统计信息：\n\n"
        result += f"📧 总邮件数：{total_count}\n"
        result += f"🔵 未读邮件：{unread_count}\n"
        result += f"📅 今日邮件：{today_count}\n"
        result += f"📎 带附件邮件：{attachment_count}\n"
        result += f"📖 已读邮件：{total_count - unread_count}\n"
        
        return result
    except Exception as e:
        return f"获取统计信息时出错：{str(e)}"

# 运行服务器
if __name__ == "__main__":
    print("正在启动Outlook MCP服务器...")
    try:
        outlook, namespace = connect_to_outlook()
        inbox = namespace.GetDefaultFolder(6)
        print(f"已连接。收件箱有 {inbox.Items.Count} 封邮件。")
        
        # 启动时加载缓存
        loaded_cache = load_email_cache()
        if loaded_cache:
            email_cache = loaded_cache
            print(f"启动时已从缓存文件加载了 {len(email_cache)} 封邮件")
        
        mcp.run()
    except Exception as e:
        print(f"启动服务器时出错：{str(e)}")
