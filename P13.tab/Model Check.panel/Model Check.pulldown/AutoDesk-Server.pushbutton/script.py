# -*- coding: utf-8 -*-
__title__ = "AutoDesk\nCheck Server"
__author__ = "Tee_เพิ่มพงษ์"
__doc__ = "pyRevit Script สำหรับตรวจสอบสถานะ Autodesk"

import requests
import re
import time
from pyrevit import forms
from pyrevit import script


target_service = "Revit Cloud Worksharing / Cloud Models"


def get_service_status(html_content, service_name):
    """ตรวจสอบสถานะบริการจาก HTML"""
    try:
        pattern = (
            r'(?:<span[^>]*class="name"|<div[^>]*class="component-name")[^>]*>\s*'
            + re.escape(service_name)
            + r'\s*</(?:span|div)>[\s\S]*?(?:<span|<div)[^>]*class="component-status[^"]*"[^>]*>([\s\S]*?)</(?:span|div)>'
        )
        match = re.search(pattern, html_content, re.IGNORECASE)
        if match:
            status = re.sub(r'\s+', ' ', match.group(1)).strip()
            return status
        return "ไม่พบข้อมูล"
    except Exception as e:
        return "ข้อผิดพลาด: " + str(e)


def get_thai_time():
    """คืนค่าเวลาปัจจุบันในรูปแบบภาษาไทย"""
    from datetime import datetime
    now = datetime.now()
    thai_months = [
        "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
        "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
        "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
    ]
    return "วันที่ {} {} พ.ศ. {} เวลา {}".format(
        now.day, thai_months[now.month - 1], now.year + 543, now.strftime("%H:%M:%S")
    )


def main():
    output = script.get_output()
    output.print_html("<meta charset='utf-8'>")

    try:
        output.print_html("<p><strong>🔍 กำลังตรวจสอบสถานะเซิร์ฟเวอร์ Autodesk...</strong></p>")

        # ตรวจสอบสถานะครั้งแรก
        response1 = requests.get("https://health.autodesk.com/", timeout=15)
        response1.raise_for_status()
        status1 = get_service_status(response1.text, target_service)
        output.print_html("<p>📊 สถานะเริ่มต้น: <strong>{}</strong></p>".format(status1))

        # รอสักครู่
        output.print_html("<p>⏳ รอตรวจสอบอีกครั้ง...</p>")
        time.sleep(10)

        # ตรวจสอบสถานะครั้งที่สอง
        response2 = requests.get("https://health.autodesk.com/", timeout=15)
        response2.raise_for_status()
        status2 = get_service_status(response2.text, target_service)
        output.print_html("<p>📊 สถานะล่าสุด: <strong>{}</strong></p>".format(status2))

        status_pair = "{}-{}".format(status1, status2)
        thai_time = get_thai_time()
        message_template = "{}\n\nบริการ: {}\nสถานะ: {}\n\n{}"

        if status_pair == "Operational-Operational":
            forms.alert(
                msg=message_template.format("บริการทำงานปกติ", target_service, status2, thai_time),
                title="✅ สถานะปกติ"
            )
            output.print_html("<p style='color:green;'>✅ บริการทำงานปกติ</p>")

        elif status_pair == "Operational-Outage":
            forms.alert(
                msg=message_template.format(
                    "แจ้งเตือน: บริการหยุดทำงาน\n\nบริการกำลังประสบปัญหา ทำให้ไม่สามารถใช้งานได้ชั่วคราว",
                    target_service, status2, thai_time
                ),
                title="🚨 การหยุดทำงานของบริการ"
            )
            output.print_html("<p style='color:red;'>🚨 บริการหยุดทำงาน</p>")

        elif status_pair == "Outage-Operational":
            forms.alert(
                msg=message_template.format(
                    "แจ้งเตือน: บริการกลับมาใช้งานได้แล้ว", target_service, status2, thai_time
                ),
                title="✅ การกู้คืนบริการ"
            )
            output.print_html("<p style='color:green;'>✅ บริการกลับมาใช้งานได้แล้ว</p>")

        elif status_pair == "Operational-Degraded Performance":
            forms.alert(
                msg=message_template.format(
                    "แจ้งเตือน: ประสิทธิภาพบริการลดลง\n\nบริการยังใช้งานได้แต่ทำงานช้ากว่าปกติ",
                    target_service, status2, thai_time
                ),
                title="⚠️ ประสิทธิภาพลดลง"
            )
            output.print_html("<p style='color:orange;'>⚠️ ประสิทธิภาพบริการลดลง</p>")

        elif status_pair == "Degraded Performance-Operational":
            forms.alert(
                msg=message_template.format(
                    "แจ้งเตือน: บริการกลับมาทำงานด้วยประสิทธิภาพปกติ",
                    target_service, status2, thai_time
                ),
                title="✅ การกู้คืนประสิทธิภาพ"
            )
            output.print_html("<p style='color:green;'>✅ ประสิทธิภาพบริการกลับสู่ปกติ</p>")

        else:
            forms.alert(
                msg=message_template.format(
                    "แจ้งเตือน: สถานะบริการเปลี่ยนจาก {} เป็น {}".format(status1, status2),
                    target_service,
                    "{} → {}".format(status1, status2),
                    thai_time
                ),
                title="📝 การเปลี่ยนแปลงสถานะ"
            )
            output.print_html(
                "<p>📝 สถานะเปลี่ยนแปลงจาก <strong>{}</strong> เป็น <strong>{}</strong></p>".format(status1, status2)
            )

        # ✅ เพิ่มปุ่มเปิดลิงก์ภายนอกแทน HTML ลิงก์
        output.print_html("<br><b>🔗 เปิดหน้าเว็บสถานะ Autodesk ด้านล่าง:</b><br>")
        if forms.alert(
            msg="ต้องการเปิดหน้า Autodesk Health Dashboard ในเบราว์เซอร์หรือไม่?",
            title="เปิดหน้าเว็บ Autodesk",
            options=["เปิดในเบราว์เซอร์", "ยกเลิก"]
        ) == "เปิดในเบราว์เซอร์":
            script.open_url("https://health.autodesk.com/")

    except requests.RequestException as e:
        msg = "ไม่สามารถเชื่อมต่อเพื่อตรวจสอบสถานะได้\n\nข้อผิดพลาด: {}".format(str(e))
        forms.alert(msg=msg, title="❌ ข้อผิดพลาดในการเชื่อมต่อ")
        output.print_html("<p style='color:red;'>❌ ข้อผิดพลาด: {}</p>".format(str(e)))

    except Exception as e:
        msg = "เกิดข้อผิดพลาดที่ไม่คาดคิด\n\nรายละเอียด: {}".format(str(e))
        forms.alert(msg=msg, title="❌ ข้อผิดพลาด")
        output.print_html("<p style='color:red;'>❌ ข้อผิดพลาด: {}</p>".format(str(e)))


if __name__ == "__main__":
    main()
