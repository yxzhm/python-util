import os
# import xml.etree.ElementTree as ET
from lxml import etree


def find_xml_files(path, file_list):
    for root, dirs, files in os.walk(path):
        for file in files:
            file_list.append(root + "/" + file)


def filter_files_by_keyword(file_list, keyword):
    result = []
    for file in file_list:
        with open(file, encoding="utf-8") as f:
            for line in f.readlines():
                if keyword in line:
                    result.append(file)
    return result


def replace_target_by_source(src_list, tar_list, keywords):
    for src in src_list:
        tar = find_target_file(src, tar_list)
        if tar is None:
            continue

        replace_xml_content(src, tar, keywords)


def find_target_file(src, tar_list):
    src_file_name = get_file_name(src)
    for tar in tar_list:
        if src_file_name == get_file_name(tar):
            return tar
    return None


def get_file_name(file):
    if "/" in file:
        return file[file.rindex('/') + 1:]
    return file


def replace_xml_content(src, tar, keywords):
    src_xml = etree.parse(src)
    tar_xml = etree.parse(tar)

    src_element = None
    modified_target = False
    for keyword in keywords:
        for child in src_xml.xpath('//insert'):
            if keyword == child.attrib['id']:
                src_element = child
                break
        if src_element is not None:
            for child in tar_xml.xpath('//insert'):
                if keyword == child.attrib['id']:
                    parent = child.getparent()
                    parent.insert(parent.index(child) + 1, src_element)
                    parent.remove(child)
                    src_element = None
                    modified_target = True
                    break
    if modified_target:
        tar_xml.write(tar, encoding="UTF-8", xml_declaration=True)

    print("handling " + tar)


if __name__ == '__main__':
    print("start")
    source = "D:/OneDrive/Desktop/source"
    target = "D:/OneDrive/Desktop/target"

    source_files = []
    target_files = []
    find_xml_files(source, source_files)
    find_xml_files(target, target_files)

    source_files = filter_files_by_keyword(source_files, "property=\"iParamseqno\"")

    replace_target_by_source(source_files, target_files, ["insert", "insertBatch", "insertOrUpdateBatch"])

    print("abc")
