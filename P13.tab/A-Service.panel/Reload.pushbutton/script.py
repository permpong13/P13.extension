# -*- coding: utf-8 -*-
import os
import subprocess
from pyrevit import forms, script
from pyrevit.loader import sessionmgr

# ==========================================
# SETTINGS
# ==========================================
SOURCE_LAN = r"\\tee\P13.extension" 
ADMIN_USER = "Permpong13"

def get_repo_root(path):
    parent = os.path.dirname(path)
    if os.path.exists(os.path.join(parent, ".git")):
        return parent
    elif len(parent) > 3:
        return get_repo_root(parent)
    return None

def sync_tools():
    current_user = os.environ.get('USERNAME')
    current_script_path = os.path.abspath(__file__)
    dest_path = get_repo_root(current_script_path)

    # 1. Admin Protection (คงเดิม)
    if current_user == ADMIN_USER:
        forms.alert("โหมด Admin: ระบบจะไม่รันการ Sync ทับไฟล์งานของคุณ", title="Admin Notice")
        return

    if not dest_path:
        forms.alert("ไม่พบโฟลเดอร์ติดตั้งในเครื่องนี้", title="Error")
        return

    # 2. เริ่มการตรวจสอบช่องทาง
    # ลองเช็คว่าเชื่อมต่อ LAN ได้ไหม
    lan_available = os.path.exists(SOURCE_LAN)

    if lan_available:
        mode = "LAN (Fast Sync)"
        cmd = 'robocopy "{}" "{}" /MIR /Z /MT:8 /R:1 /W:1 /XF *.pyc /NFL /NDL'.format(SOURCE_LAN, dest_path)
    else:
        mode = "GitHub (Cloud Sync)"
        cmd = 'git -C "{}" pull'.format(dest_path)

    # 3. ยืนยันการอัปเดต
    msg = "ตรวจพบการเชื่อมต่อแบบ: {}\nคุณต้องการอัปเดตเครื่องมือ BIM หรือไม่?".format(mode)
    res = forms.alert(msg, title="Update P13 Tools", options=["Update Now", "Cancel"])
                      
    if res == "Update Now":
        try:
            with forms.ProgressBar(title="กำลังอัปเดตจาก {}...".format(mode)) as pb:
                subprocess.call(cmd, shell=True)
            
            forms.alert("อัปเดตสำเร็จ! ระบบจะรีโหลดเมนู Revit", title="Success")
            sessionmgr.reload_pyrevit()
        except Exception as e:
            forms.alert("เกิดข้อผิดพลาด: {}".format(str(e)))

if __name__ == '__main__':
    sync_tools()