import os
import datetime
import optparse
import docwriter
from htmldocparser import HtmlDocParser as HtmlDocParser
from datatype import DataType as DataType
import util


if __name__ == '__main__':
    usage = """
    软件：doccrawler v1.0.0
    作者：二部十五室 喻鹤
    说明：根据doxygen生成的html文档提取C++类型信息，将其转换未word文档
          [工作方式1]：通过-d或--html_dir指定doxygen生成的html目录(确保生成annotated.html)自动生成工程的类型描述信息，输出到word文档
          [工作方式2]：通过-f或--file_dir指定存放描述类型信息的html文档的目录，目录中的html文档将被解析，输出到word文档
    注意：[工作方式1]和[工作方式2]可同时工作，但所有信息将被输出到同一个文件中
    """
    opt_parser = optparse.OptionParser(usage=usage)
    opt_parser.add_option("-f", "--file_dir", dest="file_dir",
                          help="FILE_DIR:doxygen生成的描述数据类型的html文件所在目录")
    opt_parser.add_option("-d", "--html_dir", dest="html_dir",
                          help="HTML_DIR:doxygen生成的html目录，其中需生成annotated.html文件")
    opt_parser.add_option("-o", "--output", dest="outdoc_name",
                          help="OUTDOC_NAME:生成的office word 文件名称(如：path/filename.doc)",
                          default="{0}\类型文档说明_{1}.doc".format(os.getcwd(), datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")))
    opt_parser.add_option("-t", "--title", dest="start_title",
                          help="START_TITLE:生成的office word文件中类型描述信息的标题级别",
                          default=2)

    opts, args = opt_parser.parse_args()
    # 获取输出的word文件名称
    doc_name = opts.outdoc_name
    print("获取输出文件名称" + doc_name)

    # DocWriter对象
    print("标题级别：{0}".format(opts.start_title))
    doc_writer = docwriter.DocWriter(doc_name, start_title=int(opts.start_title))

    # 获取文件列表
    file_list =[]
    if opts.file_dir:
        file_list += util.get_file_list(opts.file_dir)
    if opts.html_dir:
        file_list += HtmlDocParser.get_data_files(opts.html_dir)

    # 开始处理文件
    total_count = len(file_list)
    for n, html_file in enumerate(file_list):
        print('处理文件:[{0}/{1}]'.format(n + 1, total_count) + html_file)
        doc_parser = HtmlDocParser(html_file)
        name = doc_parser.get_class_name()
        desc = doc_parser.get_class_desc()
        public_var_list = doc_parser.get_var_info('Public')
        protected_var_list = doc_parser.get_var_info('Protected')
        private_var_list = doc_parser.get_var_info('Private')
        public_fun_list = doc_parser.get_fun_info('Public')
        protected_fun_list = doc_parser.get_fun_info('Protected')
        private_fun_list = doc_parser.get_fun_info('Private')
        typedef_list = doc_parser.get_typedefs()
        enum_list = doc_parser.get_enums()
        dt = DataType(name, desc,
                      [public_var_list, protected_var_list, private_var_list],
                      [public_fun_list, protected_fun_list, private_fun_list],
                      typedef_list, enum_list)
        doc_writer.write(dt)
    doc_writer.save()
    print('处理完毕，输出文件路径：{0}'.format(doc_name))
    print('等待程序退出...')
