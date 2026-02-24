import csv
import json
import os
import time

import requests
import win32com.client
import pywintypes


class Contact:
    Name = ''
    Department = ''
    Mail = ''
    Phone = ''


def get_contacts(url, sid, cookies) -> {}:
    cookie_dic = get_cookies(cookies)
    res = requests.post(url + '?func=oab%3AgetDirectories&sid=' + sid,
                        json={"attrIds": ["email"], "needChildCount": True, "depth": 3}, cookies=cookie_dic,
                        verify=False)

    body = res.json()
    if body['code'] != 'S_OK':
        return None

    contacts = {}
    for var in body['var']:
        for department in var['ou']:
            if "ou" not in department.keys():
                get_department(url, sid, cookie_dic, department['id'], "", department['name'], contacts)
            else:
                for sub_depart in department['ou']:
                    get_department(url, sid, cookie_dic, sub_depart['id'], department['name'] + "/", sub_depart['name'],
                                   contacts)

    return contacts


def get_cookies(cookies) -> {}:
    cookie_dic = {}
    for cookie in cookies.split(';'):
        key = cookie.split('=')[0]
        value = cookie.split('=')[1]
        cookie_dic[key] = value
    return cookie_dic


def get_department(url, sid, cookie_dic, department_id, department_prefix, department_name, contacts):
    if department_id == 'ggyx':
        return

    data = {
        "start": 0,
        "limit": 200,
        "defaultReturnMeetingRoom": False,
        "dn": "cfjk/" + department_id,
        "returnAttrs": [
            "@id",
            "true_name",
            "email",
            "mobile_number"
        ]
    }

    res = requests.post(url + '?func=oab%3AlistEx&sid=' + sid,
                        cookies=cookie_dic,
                        verify=False,
                        headers={'Content-Type': 'text/x-json'},
                        json=data)
    body = res.json()
    if body['code'] != 'S_OK':
        return

    for var in body['var']:
        contact = Contact()
        contact.Name = var['true_name']
        contact.Department = department_prefix + department_name
        contact.Mail = var['email']
        contact.Phone = var['mobile_number']
        if contact.Mail in contacts.keys():
            continue

        contacts[contact.Mail] = contact

    time.sleep(1)


def add_outlook_contact(contacts):
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")

    existed_mails = {}
    for item in ns.GetDefaultFolder(10).Items:
        if hasattr(item, "Email1Address"):
            existed_mails[item.Email1Address] = "1"
            if item.Email1Address in contacts:
                old_info = get_contact_critical_info(item)
                set_contact_item(item, item.Email1Address, contacts)
                new_info = get_contact_critical_info(item)
                if old_info != new_info:
                    print("update " + old_info + " to " + new_info)
                    item.save()

    for contact_mail in contacts:
        contact = contacts[contact_mail]
        if contact.Name == '' or contact.Mail == '':
            continue
        if contact.Mail in existed_mails.keys():
            continue
        item = ns.GetDefaultFolder(10).Items.Add("IPM.Contact")
        set_contact_item(item, contact_mail, contacts)
        print("add new contact mail " + contact_mail)
        item.Save()


def get_contact_critical_info(item):
    return item.Email1DisplayName


def set_contact_item(item, mail_address, contacts):
    contact = contacts[mail_address]
    item.LastName = contact.Name[0]
    item.FirstName = contact.Name[1:]
    item.Department = contact.Department
    item.Email1Address = contact.Mail
    display_tip = contact.Department
    if "/" in contact.Department:
        prefix = contact.Department[0:2]
        postfix = contact.Department[contact.Department.index("/") + 1:]
        display_tip = postfix if prefix in postfix else postfix + "（" + prefix + "）"
    item.Email1DisplayName = contact.Name + "-" + display_tip
    if contact.Phone is not None:
        item.BusinessTelephoneNumber = contact.Phone


def export_foxmail_contact(contacts, csv_file):
    with open(csv_file, 'w', encoding="ANSI") as file:
        file.write("Name,E-mail Address\n")
        for contact_mail in contacts:
            contact = contacts[contact_mail]
            if contact.Name == '' or contact.Mail == '':
                continue
            file.write(contact.Name + "," + contact.Mail + "\n")


def read_from_file(text_path):
    contacts = {}
    contact = Contact()
    for root, dirs, files in os.walk(text_path):
        for file in files:
            if contact.Name != "" and contact.Mail != "":
                contact.Department = department
                contacts[contact.Mail] = contact
                contact = Contact()

            department = os.path.splitext(file)[0]
            with open(text_path + "/" + file, 'r', encoding="UTF-8") as f:
                for line in f:
                    if line is None:
                        continue
                    if line.strip() == "" and contact.Name != "" and contact.Mail != "":
                        contact.Department = department
                        contacts[contact.Mail] = contact
                        contact = Contact()
                    if line.strip() != "" and "@" in line and contact.Mail == "":
                        contact.Mail = line.strip()
                    if line.strip() != "" and line.strip().isnumeric():
                        contact.Phone = line.strip()
                    if line.strip() != "" and contact.Name == "":
                        contact.Name = line.strip()
    return contacts


def update_department():
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")

    existed_mails = {}
    for item in ns.GetDefaultFolder(10).Items:
        if hasattr(item, "Email1Address"):
            existed_mails[item.Email1Address] = "1"
            if item.Department != "":
                if "（成都）" in item.Department:
                    item.Department = "成都分公司/" + item.Department.replace("（成都）", "")
                    item.Save()
                if "（天津）" in item.Department:
                    item.Department = "天津分公司/" + item.Department.replace("（天津）", "")
                    item.Save()
                if "（上海）" in item.Department:
                    item.Department = "上海分公司/" + item.Department.replace("（上海）", "")
                    item.Save()


if __name__ == '__main__':
    print("abc")
    # f = open("mail_config.json")
    # config = json.load(f)
    # contacts = get_contacts(config['url'],
    #                         # replace the star characters with correct domain name
    #                         config['sid'],  # find the sid via F12 Chrome Developer tools
    #                         config['cookie'])
    contacts = read_from_file("./mail-contacts")
    is_outlook = True
    if is_outlook:
        add_outlook_contact(contacts)
    else:
        export_foxmail_contact(contacts, "D:/Work/contacts.csv")
