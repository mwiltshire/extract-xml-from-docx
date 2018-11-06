import zipfile
import os
from os import walk
import io
import fnmatch
import argparse
import xml.dom.minidom
from os.path import join, dirname, basename


def get_docx_in_filelist(filelist, dirname):
    return [file for file in filelist if fnmatch.fnmatch(file, "*.docx")]


def unpack_and_read(docx):
    return zipfile.ZipFile(docx).read("word/document.xml")


def walk_directory_tree_and_write_files(base_directory, should_pretty_print):
    for dirname, _, filelist in walk(base_directory):
        if not basename(dirname).startswith("."):
            print("Checking in directory: {}".format(dirname))
            for file in get_docx_in_filelist(filelist, dirname):
                print("Found file {}. Unpacking...".format(file))
                document_xml = unpack_and_read(join(dirname, file))
                write_xml(document_xml, file, base_directory,
                          should_pretty_print)


def write_files_from_base_directory(base_directory, should_pretty_print):
    files_in_base_directory = os.listdir(base_directory)
    for file in get_docx_in_filelist(files_in_base_directory, base_directory):
        print("Found file: {}".format(file))
        document_xml = unpack_and_read(join(base_directory, file))
        write_xml(document_xml, file, base_directory, should_pretty_print)


def write_xml(document_xml, file, base_directory, should_pretty_print):
    target_filename = file + ".document.xml"
    if should_pretty_print:
        pretty_print_xml(document_xml, file, base_directory)
    else:
        with io.open(join(base_directory, target_filename), "wb") as target:
            target.write(document_xml)
            print("Finished writing file: {}".format(target_filename))


def pretty_print_xml(document_xml, file, base_directory):
    target_filename = file + ".document.xml"
    try:
        dom = xml.dom.minidom.parseString(document_xml)
        pretty_xml = dom.toprettyxml()
        with io.open(join(base_directory, target_filename), "w") as target:
            target.write(pretty_xml)
            print("Finished writing file: {}".format(target_filename))
    except Exception as e:
        print("Error: Could not parse and pretty print xml. {}. Writing instead as default".format(str(e)))
        write_xml(document_xml, file, base_directory, False)


def main():
    base_directory = os.getcwd()
    argparser = argparse.ArgumentParser(
        description="Extract document.xml from docx files")
    argparser.add_argument("-r", dest="search_method", action="store_const",
                           const=walk_directory_tree_and_write_files, default=write_files_from_base_directory,
                           help="Walk directory structure and extract xml to top level dir (default: search only in current directory)")
    argparser.add_argument("-p", dest="pretty_print", action="store_true",
                           help="Pretty print extracted xml file", default=False)
    args = argparser.parse_args()
    args.search_method(base_directory, args.pretty_print)
    print("Finished")


if __name__ == "__main__":
    main()
