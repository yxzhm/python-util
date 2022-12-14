import csv
import json
import time

import requests
import win32com.client
import pywintypes


class Contact:
    Name = ''
    Department = ''
    Mail = ''


def get_contacts(url, sid, cookies) -> {}:
    cookie_dic = get_cookies(cookies)
    res = requests.post(url + '?func=oab%3AgetDirectories&sid=' + sid, cookies=cookie_dic, verify=False)

    body = res.json()
    if body['code'] != 'S_OK':
        return None

    contacts = {}
    for var in body['var']:
        for department in var['ou']:
            if "ou" not in department.keys():
                get_department(url, sid, cookie_dic, department['id'], department['name'], contacts)
            else:
                for sub_depart in department['ou']:
                    get_department(url, sid, cookie_dic, sub_depart['id'], sub_depart['name'], contacts)

    return contacts


def get_cookies(cookies) -> {}:
    cookie_dic = {}
    for cookie in cookies.split(';'):
        key = cookie.split('=')[0]
        value = cookie.split('=')[1]
        cookie_dic[key] = value
    return cookie_dic


def get_department(url, sid, cookie_dic, department_id, department_name, contacts):
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
        contact.Department = department_name
        contact.Mail = var['email']
        if contact.Mail in contacts.keys():
            continue

        contacts[contact.Mail] = contact

    time.sleep(1)


def add_outlook_contact(contacts):
    o = win32com.client.Dispatch("Outlook.Application")
    ns = o.GetNamespace("MAPI")

    existed_mails = {}
    contactsFolder = ns.GetDefaultFolder(10)
    for item in contactsFolder.Items:
        if hasattr(item, "Email1Address"):
            existed_mails[item.Email1Address] = "1"

    for contact_mail in contacts:
        contact = contacts[contact_mail]
        if contact.Name == '' or contact.Mail == '':
            continue
        if contact.Mail in existed_mails.keys():
            continue
        ContactItem = contactsFolder.Items.Add("IPM.Contact")
        ContactItem.LastName = contact.Name[0]
        ContactItem.FirstName = contact.Name[1:]
        ContactItem.Department = contact.Department
        ContactItem.Email1Address = contact.Mail
        ContactItem.Save()


def export_foxmail_contact(contacts, csv_file):
    with open(csv_file, 'w', encoding="ANSI") as file:
        file.write("Name,E-mail Address\n")
        for contact_mail in contacts:
            contact = contacts[contact_mail]
            if contact.Name == '' or contact.Mail == '':
                continue
            file.write(contact.Name + "," + contact.Mail + "\n")


if __name__ == '__main__':
    print("abc")
    contacts = get_contacts('https://mail.****.cn/coremail/s/json', # replace the star characters with correct domain name
                            '', # find the sid via F12 Chrome Developer tools
                            ''  # find the cookie strings via F12 Chrome Developer tools
                            )
    is_outlook = False
    if is_outlook:
        add_outlook_contact(contacts)
    else:
        export_foxmail_contact(contacts, "c:/Foxmail 7.2/contacts.csv")
