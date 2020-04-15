import pathlib
from pyquery import PyQuery as pq


class HtmlDocParser:
    """
    用于解析docxygen生成的html文档，生成类信息
    """
    def __init__(self, file_path):
        """
        构造函数
        :param file_path: 需要解析的html文件的路径
        :param src_encoding: 解析的html文件的原始编码
        """
        with open(file_path, encoding='utf-8') as file:
            self.utf8_html = file.read()
        # 创建pyQuery对象，parser采用html否则无法识别原生的html标签
        self.doc = pq(self.utf8_html, parser='html')

    @staticmethod
    def get_data_files(html_dir):
        """
        从dir制定的目录中查找annotated.html文件，并从中提取各数据类型的html文件名
        :param html_dir: 搜索根目录
        :return: 数据类型定义html文件名定义列表
        """
        # 从annotated.html中提取所有class属性为el且href属性为class或struct开头的标签a的href属性即为类型描述所在的html文件
        if not pathlib.Path(html_dir).exists():
            print(html_dir + ' 目录不存在 !')
            exit(-1)
        filename = html_dir + '/annotated.html'
        if not pathlib.Path(filename).exists():
            print('在目录：{0}中未找到 annotated.html '.format(html_dir))
            exit(-1)
        print('找到annotated.html 在：{0}'.format(html_dir))
        with open(filename, encoding='utf-8') as file:
            utf8_html = file.read()
        doc = pq(utf8_html)
        a = doc('.directory .el').not_('[href^=namespace]')
        file_list = ['{0}/{1}'.format(html_dir, item.attr.href) for item in a.items()]
        # print(file_list)
        return file_list

    def get_class_name(self):
        """
        获取类名称
        :return:类的名称
        """
        # 利用正则表达式提取类名
        class_name = self.doc('.title').text().split(' ')[0]
        # result = re.search('\w+\s', class_name)
        # if result:
        #     class_name = result.group()
        # print('类名：', '\n', '  ', class_name)
        return class_name

    def get_class_desc(self):
        """
        获取类的描述信息
        :return: 类的描述信息
        """
        h2 = self.doc('h2:contains(详细描述)')
        class_desc = h2.next().text().strip()
        # print('描述：', '\n', '  ', h2.next().text().replace('\n', '。'))
        return class_desc

    def get_var_info(self, visiable='Public'):
        """
        根据visiable获取成员变量列表
        :param visiable: 成员变量可见性，取值:Public Protected Private
        :return: 对应可见性的成员变量列表
        """
        var_list = []
        attr_table = self.doc("table:contains('{0} 属性')".format(visiable))
        mem_items = attr_table('tr[class^="memitem"]')
        for item in mem_items.items():
            # 排除继承自基类的成员
            if item.attr('class').find('inherit') > 0:
                continue
            mem_visiable = visiable.lower()
            mem_id = item('a[id]').attr('id')
            mem_type = item('.memItemLeft').text()
            mem_name = item('.memItemRight').text()
            mem_desc = ''
            desc = item.next()
            if desc.attr('class') == 'memdesc:{0}'.format(mem_id):
                mem_desc = desc('.mdescRight').text()
            var_list.append((mem_type, mem_name, mem_desc))
            # print('  ', mem_type, mem_name, mem_desc)
        return var_list

    def get_fun_info(self, visiable='Public'):
        """
        根据visiable获取成员函数列表
        :param visiable: 成员函数可见性，取值:Public Protected Private
        :return: 对应可见性的函数描述列表
        """
        # previous_is_template = False  # 表示上一个memitem节点是否为模板参数节点
        fun_list = []
        fun_table = self.doc("table:contains('{0} 成员函数')".format(visiable))
        fun_items = fun_table('tr[class^="memitem"]')
        previous_is_template = False
        for fun_item in fun_items.items():
            # 排除基类对象
            if fun_item.attr('class').find('inherit') > 0:
                continue
            # 排除模板参数声明的memitem节点，memitem的第一个td节点的class属性如果为memTemplParams则排除
            if len(fun_item('td:first-child[class=memTemplParams]').text()):
                previous_is_template = True
                continue
            fun_id = fun_item.attr('class').split(':')[1]
            if previous_is_template:
                template_desc = fun_item.prev().text()
                fun_type = fun_item('.memTemplItemLeft').text()
                fun_name = fun_item('.memTemplItemRight').text()
            else:
                template_desc = ''
                fun_type = fun_item('.memItemLeft').text()
                fun_name = fun_item('.memItemRight').text()
            previous_is_template = False
            fun_desc = ''
            fun_param_list = []
            fun_ret = ''
            desc = fun_item.next()
            if desc.attr('class') == 'memdesc:{0}'.format(fun_id):
                if desc('.mdescRight a').text() == '更多...':
                    fun_detail_table = self.doc('a[id={0}]'.format(fun_id)).next().next()
                    plist = []
                    for p in fun_detail_table('.memdoc p').items():
                        plist.append(p.text())
                    fun_desc = '。'.join(plist)
                    trs = fun_detail_table('.memdoc .params .params tr')  # 参数说明tr
                    for tr in trs.items():
                        fun_param_list.append(
                            ' '.join([tr('.paramdir').text(), tr('.paramname').text(), tr('td:last-child').text()]))
                    # html中的返回
                    fun_ret = fun_detail_table('.memdoc .section.return dd').text()

                    #html中的返回值
                    fun_ret = [fun_ret]
                    for tr in fun_detail_table('.memdoc .retval tr').items():
                        fun_ret.append(tr('td').text())
                    fun_ret = '\n'.join(fun_ret)
                else:
                    fun_desc = desc('.mdescRight').text()
                fun_list.append((fun_type, fun_name, fun_desc, fun_param_list, fun_ret, template_desc))
        return fun_list

    def get_typedefs(self):
        """
        获取数据类型中的数据类型重定义列表
        :return: 重定义变量的二元组列表，每个元组内容为重定义后的类型名称和类型描述
        """
        typedef_list = []
        mem_items = self.doc('tr[class^="memitem"]:contains(typedef)')
        for item in mem_items.items():
            typedef_list.append((item.text().replace('\n', ' '), item.next().text()))
        return typedef_list

    def get_enums(self):
        """
        获取类型中定义的枚举类型列表
        :return:枚举类型的二元组列表，每个二元组内容分别为枚举类型名称和枚举类型描述
        """
        ths = self.doc('tr th:contains(枚举值)')
        enum_list = []
        for th in ths.items():
            for tr in th.parent().siblings().items():
                enum_list.append((tr('.fieldname').text(), tr('.fielddoc').text()))
        return enum_list
