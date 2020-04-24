import os


def create_folder(path):
    path = path.strip()
    path = path.rstrip('\\')
    if not os.path.exists(path):
        os.makedirs(path)
        return path


def main():
    create_folder(r'e:\报表\晨报小时报模板\use\日报')


if __name__ == '__main__':
    main()
