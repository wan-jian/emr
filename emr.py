# coding=utf-8
# designed by Wan Jian

from application import app
import sys


def main():
    try:
        app.do_processes()
    except Exception as e:
        sys.stderr.write(str(e) + '\n')
        sys.exit(11)


if __name__ == '__main__':
    main()
