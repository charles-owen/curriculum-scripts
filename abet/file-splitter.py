import os
import xml.dom.minidom

items="faculty"

def process() :
    doc = xml.dom.minidom.parse(f"{items}.xml")

    for node in doc.firstChild.childNodes:
        if node.nodeType == node.ELEMENT_NODE and node.nodeName == items:
             course(node)
# <course number="CSE 231"
# name="Introduction to Programming I"
# credits="4"
# contact="4"
# type="Fundamental">

def course(node) :
    number = node.getAttribute("number")
    name = node.getAttribute("name")
    name = format_name(name)
    print(name)

    course_doc = xml.dom.minidom.Document()
    course_doc.appendChild(node)
    xml_string = course_doc.toprettyxml(indent="  ")
    str = "\n".join(line for line in xml_string.splitlines() if line.strip())
    with open(f"{items}/{name}.xml", "w") as f:
        f.write(str)


def format_name(name: str) -> str:
    parts = name.strip().split()
    first, last = parts[0], parts[-1]
    middle = parts[1:-1]  # Capture middle names if present

    formatted_name = f"{last.lower()}-{first.lower()}"
    if middle:
        formatted_name += f"-{'-'.join(m.lower() for m in middle)}"

    return formatted_name


if __name__ == '__main__':
    process()