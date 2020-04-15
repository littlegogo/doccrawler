
class DataType:
    """
    用于表示类或结构体类型的成员信息
    """
    def __init__(self, name, desc, varinfo, funinfo, typedefs, enums):
        """
        构造数据类型描述信息
        :param name: 数据类型的名称，类名或结构体名称
        :param desc: 数据类型的描述信息，类或结构体的描述信息
        :param varinfo: 描述类型成员变量信息的列表，列表元素依次为描述public，protected, private 成员的列表
        :param funinfo: 描述类型成员函数信息的列表，列表元素依次为描述public，protected, private 方法的列表
        :param typedefs: typedef 重定义类型列表
        :param enums: 类型内的枚举变量列表
        """
        self.name = name
        self.desc = desc
        self.public_var_list = varinfo[0]
        self.protected_var_list = varinfo[1]
        self.private_var_list = varinfo[2]
        self.public_fun_list = funinfo[0]
        self.protected_fun_list = funinfo[1]
        self.private_fun_list = funinfo[2]
        self.typedef_list = typedefs
        self.enum_list = enums

    def __str__(self):
        """
        输出自身表示的数据类型的信息
        :return: 无
        """
        print('类型名称：', self.name)
        print('类型描述：', self.desc)
        print('公有属性：')
        print('  ', self.public_var_list)
        print('保护属性：')
        print('  ', self.protected_var_list)
        print('私有属性：')
        print('  ', self.private_var_list)
        print('public 方法：')
        print('  ', self.public_fun_list)
        print('protected 方法：')
        print('  ', self.protected_fun_list)
        print('private 方法：')
        print('  ', self.private_fun_list)
        print('类型重定义：')
        print('  ', self.typedef_list)
        print('枚举类型定义：')
        print('  ', self.enum_list)
        return ''
