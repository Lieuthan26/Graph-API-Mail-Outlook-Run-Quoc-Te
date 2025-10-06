import requests
from msal import ConfidentialClientApplication
import os
from datetime import datetime, timedelta

client_id = os.environ["CLIENT_ID"]
client_secret = os.environ["CLIENT_SECRET"]
tenant_id = os.environ["TENANT_ID"]
user_email = os.environ["USER_EMAIL"]

authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
token = app.acquire_token_for_client(scopes=scope)

if "access_token" in token:
    headers = {
        "Authorization": f"Bearer {token['access_token']}",
        "Content-Type": "application/json"
    }

    # 1. Create a new event in Outlook calendar
    start_time = (datetime.utcnow() + timedelta(minutes=5)).strftime("%Y-%m-%dT%H:%M:%S")
    end_time = (datetime.utcnow() + timedelta(minutes=30)).strftime("%Y-%m-%dT%H:%M:%S")
    event_data = {
        "subject": "Thông báo: Workflow đã chạy thành công",
        "start": {"dateTime": start_time, "timeZone": "Asia/Ho_Chi_Minh"},
        "end": {"dateTime": end_time, "timeZone": "Asia/Ho_Chi_Minh"},
        "body": {"contentType": "HTML", "content": "Workflow kiểm tra lịch đã được thực hiện."}
    }
    event_response = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/calendar/events",
        headers=headers,
        json=event_data
    )
    if event_response.status_code == 201:
        print("✅ Đã tạo sự kiện mới trong lịch Outlook.")
    else:
        print("❌ Lỗi khi tạo sự kiện:", event_response.text)

    # 2. Send an email notification
    email_data = {
        "message": {
            "subject": "Thông báo: Workflow đã chạy",
            "body": {
                "contentType": "Text",
                "content": "Workflow kiểm tra lịch đã chạy thành công và đã tạo sự kiện mới."
            },
            "toRecipients": [{"emailAddress": {"address": user_email}}]
        },
        "saveToSentItems": "true"
    }
    email_response = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail",
        headers=headers,
        json=email_data
    )
    if email_response.status_code == 202:
        print("📧 Đã gửi email thông báo.")
    else:
        print("❌ Lỗi khi gửi email:", email_response.text)

    # 3. Add a note to an existing event (first upcoming event)
    events_response = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{user_email}/calendar/events?$top=1",
        headers=headers
    )
    if events_response.status_code == 200:
        events = events_response.json().get("value", [])
        if events:
            event_id = events[0]["id"]
            update_data = {
                "body": {
                    "contentType": "HTML",
                    "content": events[0]["body"]["content"] + "<br><b>Ghi chú:</b> Workflow đã kiểm tra lịch."
                }
            }
            patch_response = requests.patch(
                f"https://graph.microsoft.com/v1.0/users/{user_email}/events/{event_id}",
                headers=headers,
                json=update_data
            )
            if patch_response.status_code == 200:
                print("📝 Đã thêm ghi chú vào sự kiện hiện có.")
            else:
                print("❌ Lỗi khi cập nhật sự kiện:", patch_response.text)
        else:
            print("⚠️ Không tìm thấy sự kiện nào để cập nhật.")
    else:
        print("❌ Lỗi khi lấy danh sách sự kiện:", events_response.text)
else:
    print("❌ Lỗi khi lấy access token:", token.get("error_description"))
