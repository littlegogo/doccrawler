import win32com
from win32com.client import Dispatch, constants


class DocWriter:
    """
    office word 写入工具类
    """
    def __init__(self, doc_name, start_title=1):
        """
        构造函数,创建word文档对象
        :param doc_name: 包含文件名称的完整路径
        :param start_title: 类型名称对应的起始标题级别
        """
        self.doc_name = doc_name
        self.start_title = start_title
        self.word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
        self.word_app.Visible = True
        self.word_app.DisplayAlerts = 0
        self.doc = self.word_app.Documents.Add()
        self.word_app.CaptionLabels.Add('表')  # 增加一个标签

    def _fix_table(self):
        """
        将表格中的"表xx"替换为引用
        :return: 无
        """
        for i in range(1, self.doc.Tables.Count + 1):
            self.doc.Tables(i).Select()
            desc_pha = self.word_app.Selection.Previous(constants.wdParagraph, 2)  # 找到"表XX所在的段落，表格上方第2个"
            if desc_pha:
                desc_pha.Select()
                if self.word_app.Selection.Find.Execute(FindText='表XX'):
                    self.word_app.Selection.Range.InsertCrossReference(
                        ReferenceType='表',
                        ReferenceKind=constants.wdOnlyLabelAndNumber,
                        ReferenceItem='{0}'.format(i),
                        InsertAsHyperlink=True,
                        IncludePosition=False,
                        SeparateNumbers=False,
                        SeparatorString=' ')

    def _get_title_(self, title_level):
        """
        获取标题样式
        :param title_level: 标题级别
        :return: 标题级别变量
        """
        heading_style = constants.wdStyleHeading1
        if title_level == 1:
            heading_style = constants.wdStyleHeading1
        elif title_level == 2:
            heading_style = constants.wdStyleHeading2
        elif title_level == 3:
            heading_style = constants.wdStyleHeading3
        elif title_level == 4:
            heading_style = constants.wdStyleHeading4
        elif title_level == 5:
            heading_style = constants.wdStyleHeading5
        elif title_level == 6:
            heading_style = constants.wdStyleHeading6
        elif title_level == 7:
            heading_style = constants.wdStyleHeading7
        return heading_style

    def _set_table_border(self, table, left, top, right, bottom):
        """
        设置表格边框线宽
        :param left: 左边框宽度
        :param top: 上边框宽度
        :param right: 右边框宽度
        :param bottom: 下边框宽度
        :return: 无
        """
        table.Borders(constants.wdBorderLeft).LineWidth = left
        table.Borders(constants.wdBorderTop).LineWidth = top
        table.Borders(constants.wdBorderRight).LineWidth = right
        table.Borders(constants.wdBorderBottom).LineWidth = bottom

    def _write_type_title(self, type_name):
        """
        写入类型名称标题
        :param type_name: 类型名称
        :return: 无
        """
        # 创建类型名称标题对应的段落
        type_title = self.doc.Paragraphs.Add()  # 将类型名称作为标题
        type_title.Range.InsertBefore(type_name)
        # type_title.Range.Select()
        type_title.Style = self._get_title_(self.start_title)

    def _write_type_desc(self, desc):
        """
        写入类型描述
        :param desc: 类型描述
        :return: 无
        """
        # 创建类型描述对应的段落
        desc_pha = self.doc.Paragraphs.Add()  # 将类型名称作为标题
        # desc_pha.LineSpacingRule = constants.wdLineSpaceExactly
        desc_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
        desc_pha.Range.Font.Name = 'Times New Roman'
        desc_pha.Range.Font.NameFarEast = '宋体'
        desc_pha.Range.Font.Size = 12  # 小四

        desc_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
        desc_pha.Range.InsertBefore(desc)
        # desc_ph.Range.Select()
        # desc_ph.Style = self._get_title_(self.start_title)

    def _write_var_list(self, type_name, var_list, visiable):
        """
        写入类型的成员变量
        :param type_name: 类型名称
        :param var_list: 成员变量列表
        :param visiable: 成员变量列表中成员的可见性
        :return: 无
        """
        # 输出属性标题
        var_heading_pha = self.doc.Paragraphs.Add()
        var_heading_pha.Range.InsertBefore(visiable + '属性')
        # var_heading_pha.Range.Select()
        var_heading_pha.Style = self._get_title_(self.start_title + 1)
        # 输出描述
        var_contents_pha = self.doc.Paragraphs.Add()
        var_contents_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
        var_contents_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
        var_contents_pha.Range.Font.Size = 12  # 小四
        var_contents_pha.Range.Font.Name = 'Times New Roman'
        var_contents_pha.Range.Font.NameFarEast = '宋体'
        if len(var_list):
            var_contents_pha.Range.InsertBefore(visiable + '属性如表XX所示。')
            # 输出表题
            table_heading_pha = self.doc.Paragraphs.Add()
            table_heading_pha.LineSpacing = 1.5*12
            table_heading_pha.Alignment = constants.wdAlignParagraphCenter
            table_heading_pha.Range.InsertCaption("表", '', '', constants.wdCaptionPositionAbove)
            table_heading_pha.Range.InsertBefore('{0}属性列表'.format(visiable))
            table_pha = self.doc.Paragraphs.Add()
            # 输出属性表格 共3列，分别为类型名称，数据类型，描述，外边框1.5磅
            var_table = table_pha.Range.Tables.Add(table_pha.Range, len(var_list) + 1, 3)
            var_table.Columns(1).SetWidth(4*28.35, 0)  # 1cm = 28.35磅
            var_table.Columns(2).SetWidth(4*28.35, 0)  # 1cm = 28.35磅
            var_table.Columns(3).SetWidth(6.5*28.35, 0)  # 1cm = 28.35磅
            var_table.Borders.Enable = True
            self._set_table_border(var_table, constants.wdLineWidth150pt, constants.wdLineWidth150pt,
                                   constants.wdLineWidth150pt, constants.wdLineWidth150pt)
            var_table.Cell(1, 1).Range.Text = '属性名称'
            var_table.Cell(1, 2).Range.Text = '数据类型'
            var_table.Cell(1, 3).Range.Text = '数据描述'
            for index, var in enumerate(var_list):
                var_table.Cell(index + 2, 1).Range.Text = var[1]
                var_table.Cell(index + 2, 2).Range.Text = var[0]
                var_table.Cell(index + 2, 3).Range.Text = var[2]
            # var_table.Range.Select()
            var_table.Range.Font.Name = '宋体'
            var_table.Range.Font.Name = 'Times New Roman'
            var_table.Rows(1).Range.Font.Name = '黑体'
            var_table.Select()
            self.word_app.Selection.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
            # 删除题注末尾的换行符
            var_table.Select()
            ref = self.word_app.Selection.Previous(constants.wdParagraph, 2)
            ref.Select()
            self.word_app.Selection.Find.Execute(FindText='^p', ReplaceWith=' ', Replace=constants.wdReplaceOne)
            # 设置表格题注的字体
            table_heading_pha.Range.Font.Size = 12
            table_heading_pha.Range.Font.Name = '黑体'
        else:
            var_contents_pha.Range.InsertBefore('无。')


    def _write_fun_list(self, fun_list, visiable):
        """
        写入类型的成员函数
        :param fun_list: 成员函数列表
        :param visiable: 成员变量列表中成员的可见性
        :return: 无
        """
        # 输出成员函数标题
        fun_pha = self.doc.Paragraphs.Add()
        fun_pha.Range.InsertBefore(visiable + '方法')
        # fun_pha.Range.Select()
        fun_pha.Style = self._get_title_(self.start_title + 1)
        if len(fun_list):
            for index, fun in enumerate(fun_list):
                # 函数名作为标题
                fun_name = fun[1].split(' ')[0]
                fun_heading_pha = self.doc.Paragraphs.Add()
                fun_heading_pha.Range.InsertBefore(fun_name + '方法')
                fun_heading_pha.Style = self._get_title_(self.start_title + 2)
                # 输出描述
                fun_contents_pha = self.doc.Paragraphs.Add()
                fun_contents_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
                fun_contents_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
                fun_contents_pha.Range.Font.Size = 12  # 小四
                fun_contents_pha.Range.Font.Name = 'Times New Roman'
                fun_contents_pha.Range.Font.NameFarEast = '宋体'
                fun_contents_pha.Range.InsertBefore(fun_name + '方法说明如表XX所示。')
                # 输出表题
                table_heading_pha = self.doc.Paragraphs.Add()
                table_heading_pha.LineSpacing = 1.5 * 12
                table_heading_pha.Alignment = constants.wdAlignParagraphCenter
                table_heading_pha.Range.InsertCaption("表", '', '', constants.wdCaptionPositionAbove)
                table_heading_pha.Range.InsertBefore('{0}方法'.format(fun_name))
                table_pha = self.doc.Paragraphs.Add()
                # 输出属性表格 4行，2列，分别为函数原型，函数描述，参数说明，返回值，流程图，外边框1.5磅
                fun_table = table_pha.Range.Tables.Add(table_pha.Range, 5, 2)
                fun_table.Borders.Enable = True
                self._set_table_border(fun_table, constants.wdLineWidth150pt, constants.wdLineWidth150pt,
                                       constants.wdLineWidth150pt, constants.wdLineWidth150pt)
                fun_table.Columns(1).SetWidth(3*28.35, 0)  # 1cm = 28.35磅
                fun_table.Columns(2).SetWidth(12*28.35, 0)  # 1cm = 28.35磅
                fun_table.Cell(1, 1).Range.Text = '函数原型'
                fun_table.Cell(2, 1).Range.Text = '函数描述'
                fun_table.Cell(3, 1).Range.Text = '参数说明'
                fun_table.Cell(4, 1).Range.Text = '返 回 值'
                fun_table.Cell(5, 1).Range.Text = '流 程 图'
                template_desc = fun[5] + '\n' if fun[5] else ''
                fun_table.Cell(1, 2).Range.Text = template_desc + ((fun[0] + ' ' + fun[1]) if len(fun[0]) else fun[1])  # 函数声明
                fun_table.Cell(2, 2).Range.Text = fun[2]  # 函数描述
                fun_table.Cell(3, 2).Range.Text = '\n'.join(fun[3]) if len(fun[3]) else '无'  # 参数说明
                fun_table.Cell(4, 2).Range.Text = fun[4] if len(fun[4]) else '无'  # 返回值说明
                fun_table.Cell(5, 2).Range.Text = '无'  # 流程图
                fun_table.Range.Font.Name = '宋体'
                fun_table.Range.Font.Name = 'Times New Roman'
                fun_table.Cell(1, 1).Range.Font.Name = '黑体'
                fun_table.Cell(2, 1).Range.Font.Name = '黑体'
                fun_table.Cell(3, 1).Range.Font.Name = '黑体'
                fun_table.Cell(4, 1).Range.Font.Name = '黑体'
                fun_table.Cell(5, 1).Range.Font.Name = '黑体'
                fun_table.Select()
                self.word_app.Selection.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
                # 删除题注末尾的换行符
                fun_table.Select()
                ref = self.word_app.Selection.Previous(constants.wdParagraph, 2)
                ref.Select()
                self.word_app.Selection.Find.Execute(FindText='^p', ReplaceWith=' ', Replace=constants.wdReplaceOne)
                # 设置表格题注字体
                table_heading_pha.Range.Font.Size = 12
                table_heading_pha.Range.Font.Name = '黑体'
        else:
            no_fun_pha = self.doc.Paragraphs.Add()
            no_fun_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
            no_fun_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
            no_fun_pha.Range.Font.Size = 12  # 小四
            no_fun_pha.Range.Font.Name = '宋体'
            no_fun_pha.Range.InsertBefore('无。')

    def _write_typedefs(self, typedef_list):
        """
        写入类型内部的重定义类型列表
        :param typedef_list:重定义类型列表
        :return: 无
        """
        # 输出类型重定义标题
        typedef_pha = self.doc.Paragraphs.Add()
        typedef_pha.Range.InsertBefore('类型重定义')
        # fun_pha.Range.Select()
        typedef_pha.Style = self._get_title_(self.start_title + 1)
        if len(typedef_list):
            # 输出描述
            typedef_contents_pha = self.doc.Paragraphs.Add()
            typedef_contents_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
            typedef_contents_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
            typedef_contents_pha.Range.Font.Size = 12  # 小四
            typedef_contents_pha.Range.Font.Name = 'Times New Roman'
            typedef_contents_pha.Range.Font.NameFarEast = '宋体'
            typedef_contents_pha.Range.InsertBefore('类型重定义如表XX所示。')
            # 输出标题
            table_heading_pha = self.doc.Paragraphs.Add()
            table_heading_pha.LineSpacing = 1.5 * 12
            table_heading_pha.Alignment = constants.wdAlignParagraphCenter
            table_heading_pha.Range.InsertCaption("表", '', '', constants.wdCaptionPositionAbove)
            table_heading_pha.Range.InsertBefore('数据类型重定义说明')
            table_pha = self.doc.Paragraphs.Add()
            # 输出属性表格 多行，2列，分别为重定义描述/重定义说明，外边框1.5磅
            typedef_table = table_pha.Range.Tables.Add(table_pha.Range, len(typedef_list) + 1, 2)
            typedef_table.Borders.Enable = True
            self._set_table_border(typedef_table, constants.wdLineWidth150pt, constants.wdLineWidth150pt,
                                   constants.wdLineWidth150pt, constants.wdLineWidth150pt)
            typedef_table.Columns(1).SetWidth(10 * 28.35, 0)  # 1cm = 28.35磅
            typedef_table.Columns(2).SetWidth(5 * 28.35, 0)  # 1cm = 28.35磅
            typedef_table.Cell(1, 1).Range.Text = '类型定义'
            typedef_table.Cell(1, 2).Range.Text = '类型描述'
            # 填充表格内容
            for index, type_item in enumerate(typedef_list):
                typedef_table.Cell(index + 2, 1).Range.Text = type_item[0]
                typedef_table.Cell(index + 2, 2).Range.Text = type_item[1] if len(type_item[1]) else '无'
            # 设置表格字体，中文，英文，表头
            typedef_table.Range.Font.Name = '宋体'
            typedef_table.Range.Font.Name = 'Times New Roman'
            typedef_table.Cell(1, 1).Range.Font.Name = '黑体'
            typedef_table.Cell(1, 2).Range.Font.Name = '黑体'
            typedef_table.Select()
            self.word_app.Selection.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
            # 删除题注末尾的换行符
            typedef_table.Select()
            ref = self.word_app.Selection.Previous(constants.wdParagraph, 2)
            ref.Select()
            self.word_app.Selection.Find.Execute(FindText='^p', ReplaceWith=' ', Replace=constants.wdReplaceOne)

            # 设置表格题注字体
            table_heading_pha.Range.Font.Size = 12
            table_heading_pha.Range.Font.Name = '黑体'
        else:
            no_def_pha = self.doc.Paragraphs.Add()
            no_def_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
            no_def_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
            no_def_pha.Range.Font.Size = 12  # 小四
            no_def_pha.Range.Font.Name = '宋体'
            no_def_pha.Range.InsertBefore('无。')

    def _write_enums(self, enum_list):
        """
        写入类型内部的重定义类型列表
        :param enum_list:枚举类型类型列表
        :return: 无
        """
        # 输出枚举标题
        enum_pha = self.doc.Paragraphs.Add()
        enum_pha.Range.InsertBefore('枚举值定义')
        enum_pha.Style = self._get_title_(self.start_title + 1)
        if len(enum_list):
            # 输出描述
            enum_contents_pha = self.doc.Paragraphs.Add()
            enum_contents_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
            enum_contents_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
            enum_contents_pha.Range.Font.Size = 12  # 小四
            enum_contents_pha.Range.Font.Name = 'Times New Roman'
            enum_contents_pha.Range.Font.NameFarEast = '宋体'
            enum_contents_pha.Range.InsertBefore('枚举值定义如表XX所示。')
            # 输出标题
            table_heading_pha = self.doc.Paragraphs.Add()
            table_heading_pha.LineSpacing = 1.5 * 12
            table_heading_pha.Alignment = constants.wdAlignParagraphCenter
            table_heading_pha.Range.InsertCaption("表", '', '', constants.wdCaptionPositionAbove)
            table_heading_pha.Range.InsertBefore('枚举值定义说明')
            table_pha = self.doc.Paragraphs.Add()
            # 输出属性表格 多行，2列，分别为枚举值/枚举值说明，外边框1.5磅
            enum_table = table_pha.Range.Tables.Add(table_pha.Range, len(enum_list) + 1, 2)
            enum_table.Borders.Enable = True
            self._set_table_border(enum_table, constants.wdLineWidth150pt, constants.wdLineWidth150pt,
                                   constants.wdLineWidth150pt, constants.wdLineWidth150pt)
            enum_table.Columns(1).SetWidth(5 * 28.35, 0)  # 1cm = 28.35磅
            enum_table.Columns(2).SetWidth(10 * 28.35, 0)  # 1cm = 28.35磅
            enum_table.Cell(1, 1).Range.Text = '枚举值'
            enum_table.Cell(1, 2).Range.Text = '说明'
            # 填充表格内容
            for index, type_item in enumerate(enum_list):
                enum_table.Cell(index + 2, 1).Range.Text = type_item[0]
                enum_table.Cell(index + 2, 2).Range.Text = type_item[1] if len(type_item[1]) else '无'
            # 设置表格字体，中文，英文，表头
            enum_table.Range.Font.Name = '宋体'
            enum_table.Range.Font.Name = 'Times New Roman'
            enum_table.Cell(1, 1).Range.Font.Name = '黑体'
            enum_table.Cell(1, 2).Range.Font.Name = '黑体'
            enum_table.Select()
            self.word_app.Selection.Tables(1).Rows.Alignment = constants.wdAlignRowCenter
            # 删除题注末尾的换行符
            enum_table.Select()
            ref = self.word_app.Selection.Previous(constants.wdParagraph, 2)
            ref.Select()
            self.word_app.Selection.Find.Execute(FindText='^p', ReplaceWith=' ', Replace=constants.wdReplaceOne)
            # 设置表格题注字体
            table_heading_pha.Range.Font.Size = 12
            table_heading_pha.Range.Font.Name = '黑体'
        else:
            no_def_pha = self.doc.Paragraphs.Add()
            no_def_pha.CharacterUnitFirstLineIndent = 2  # 首行缩进2字符
            no_def_pha.LineSpacing = 1.5 * 12  # 设置行距1.5
            no_def_pha.Range.Font.Size = 12  # 小四
            no_def_pha.Range.Font.Name = '宋体'
            no_def_pha.Range.InsertBefore('无。')


    def write(self, data_type):
        """
        将data_type表示的数据类型信息写入配置文件
        :param data_type: 数据类型信息
        :return: 无
        """
        self._write_type_title(data_type.name)
        self._write_type_desc(data_type.desc+'。')
        self._write_typedefs(data_type.typedef_list)
        self._write_enums(data_type.enum_list)
        self._write_var_list(data_type.name, data_type.public_var_list, 'Public')
        self._write_var_list(data_type.name, data_type.protected_var_list, 'Protected')
        self._write_var_list(data_type.name, data_type.private_var_list, 'Private')
        self._write_fun_list(data_type.public_fun_list, 'Public')
        self._write_fun_list(data_type.protected_fun_list, 'Protected')
        self._write_fun_list(data_type.private_fun_list, 'Private')

    # def open_doc(self, filename=''):
    #     d = self.word_app.Documents.Open('d:\\test.docx')
    #
    #     #  先替换
    #     for i in range(1, d.Tables.Count + 1):
    #         d.Tables(i).Select()
    #         r_table_title = self.word_app.Selection.Previous(constants.wdParagraph, 1)  # 表格标题
    #         r_table_title.Select()
    #         self.word_app.Selection.Find.Execute(FindText='表XX', ReplaceWith='', Replace=constants.wdReplaceOne)
    #         r_table_title.InsertCaption("表", '', '', constants.wdCaptionPositionAbove)
    #         d.Tables(i).Select()
    #         r_table_title = self.word_app.Selection.Previous(constants.wdParagraph, 2)  # 表格标题
    #         r_table_title.Select()
    #         self.word_app.Selection.Find.Execute(FindText='^p', ReplaceWith='', Replace=constants.wdReplaceOne)
    #         d.Tables(i).Select()
    #         r_table_title = self.word_app.Selection.Previous(constants.wdParagraph, 1)  # 表格标题
    #         r_table_title.Select()
    #         self.word_app.Selection.Font.Name = '黑体'
    #         self.word_app.Selection.Font.Size = 12

    def save(self):
        """
        保存并关闭文件
        :return: 无
        """
        self._fix_table()
        self.doc.SaveAs(self.doc_name)
        self.doc.Close()



