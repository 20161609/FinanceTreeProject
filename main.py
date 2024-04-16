from lib_branch import *
from lib_shell import *

def __main__():
    try:
        json_tree = load_tree()
        root = build_tree_from_json(json_tree)
        shell = Shell(root)
        while True:
            try:
                command = input(shell.prompt)
                if len(command):
                    if command in quit_commands:
                        break
                    else:
                        shell.fetch(command)
            except:
                pass
    except Exception as e:
        print(e)


if __name__ == '__main__':
    __main__()
