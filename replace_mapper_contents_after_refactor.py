import os
import re
from lxml import etree


def find_xml_files(path):
    file_list = []
    for root, dirs, files in os.walk(path):
        for file in files:
            file_list.append(root + "/" + file)
    return file_list


def find_table_name(target_files):
    table_file_map = {}
    for file in target_files:
        with open(file, encoding="utf-8") as f:
            for line in f.readlines():
                content = line.lower()
                if "<" in content:
                    continue
                group = None
                if "update" in content:
                    group = re.search('update\\s+(\\w+)', content)

                if "from" in content:
                    group = re.search('from\\s+(\\w+)', content)

                if group is not None and len(group.groups()) >= 1:
                    table_name = group.group(1)
                    table_file_map[table_name] = file
                    break

    return table_file_map


def replace_content(source_map, target_map):
    for table_name in target_map:
        if table_name not in source_map:
            print(table_name + " not in source")
            continue

        replace_xml_content(source_map[table_name], target_map[table_name])


def replace_xml_content(source_path, target_path):
    src_xml = etree.parse(source_path)
    tar_xml = etree.parse(target_path)
    should_save = False

    src_result_map = src_xml.xpath('/mapper/*')
    if src_result_map is not None:
        for src_child in src_result_map:
            tag = src_child.tag
            if len(src_child.attrib) == 0 or 'id' not in src_child.attrib:
                continue
            src_id = src_child.attrib['id']
            tar_sub = tar_xml.xpath('/mapper/' + tag)
            for tar_child in tar_sub:
                if len(tar_child.attrib) == 0 or 'id' not in tar_child.attrib:
                    continue
                tar_id = tar_child.attrib['id']
                if tag == 'resultMap' or tar_id == src_id:
                    replacy_child(src_child, tar_child)
                    should_save = True
                    break

    if should_save:
        tar_xml.write(target_path, encoding="UTF-8", xml_declaration=True)
        print("handling " + target_path)


def replacy_child(src_child, tar_child):
    for element in tar_child:
        element.getparent().remove(element)
    n = 0
    tar_child.text = src_child.text
    for element in src_child:
        tar_child.insert(n, element)
        n = n + 1


if __name__ == '__main__':
    print("start")
    source_folder = "D:/OneDrive/Desktop/mapper/source"
    target_folder = "D:/OneDrive/Desktop/mapper/target"

    source_files = find_xml_files(source_folder)
    target_files = find_xml_files(target_folder)

    source_map = find_table_name(source_files)
    target_map = find_table_name(target_files)

    replace_content(source_map, target_map)
    print("end")
