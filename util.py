import os
import os.path

def get_file_list(dir):
    """
    获取目录下的所有文件，不包含子目录中的文件
    :param dir: 需要遍历的目录
    :return: dir目录中的文件列表
    """
    filelist = []
    for parent, dirnames, filenames in os.walk(dir):
        for filename in filenames:
            filelist.append(os.path.join(parent, filename))
    return filelist


def get_file(dir, *args):
    """
    获取目录下的所有文件，不包含子目录中的文件
    :param args: 需要遍历的目录
    :return: dir中的目录列表
    """
    filelist = []
    for parent, dirnames, filenames in os.walk(dir):
        for filename in filenames:
            if args:
                if os.path.splitext(filename)[1] in args:
                    filelist.append(os.path.join(parent, filename))
            else:
                filelist.append(os.path.join(parent, filename))
    return filelist


if __name__ == '__main__':
    print(get_file('.', '.py'))
