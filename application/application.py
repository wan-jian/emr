# coding=utf-8

import sys
import json
import getopt
from .dataproc.process1_1 import process1_1
from .dataproc.process1_2 import process1_2


class Application:
    def __init__(self):
        self.project = {}
        self.usage = 'Usage: tj.py [-h] [project file]'

        try:
            opts, args = getopt.getopt(sys.argv[1:], 'h', ['help'])
        except getopt.GetoptError:
            sys.stderr.write(self.usage + '\n')
            sys.exit(1)

        for opt, arg in opts:
            if opt in ('-h', '--help'):
                print(self.usage)
                sys.exit(0)

        if len(args) == 0:
            print("Warning: Not refer to a project file, use 'default.json' in working directory instead")
            self.project['file_name'] = 'default.json'
        elif len(args) > 1:
            sys.stderr.printf(self.usage)
            sys.exit(1)
        else:
            self.project['file_name'] = args[0]

        self.parser_project()

    def parser_project(self):
        try:
            with open(self.project['file_name'], 'r') as file:
                self.project = json.load(file)
        except KeyError as e:
            s = "Invalid key found in project file: '" + e.args[0] + "'\n"
            sys.stderr.write(s)
            sys.exit(2)
        except Exception as e:
            sys.stderr.write(str(e) + '\n')
            sys.exit(2)

    def do_processes(self):
        for process in self.project:
            if process['enabled'].upper() != 'YES':
                continue

            print("\n>> {}: {}".format(process['process_name'], process['comment']))
            if process['process_name'] == 'process1_1':
                process1_1(process)
            elif process['process_name'] == 'process1_2':
                process1_2(process)
