import json
from collections import deque
import fileDescripter

Files = []  # List for storing all expenditure datas


# Class for handling nodes in a tree structure.
class Branch:
    def __init__(self, cur_path: str, parent = None):
        self.children = {}  # child nodes
        self.parent = parent  # parent node(if root, NULL)
        self.cur_path = cur_path  # path to locate the current node.
        self.name = self.cur_path.split('/')[-1]  # current Node's name


# Find branch based on cur_path and return the branch
def search_branch(branch: Branch, path: str):
    for branchName in path.split('/'):
        if branchName in branch.children:
            branch = branch.children[branchName]
        else:  # Not valid path
            return None
    return branch


# Initiate Tree based on local json file.
def make_tree():
    # Root node of Tree
    root = Branch("HOME", None)

    file_path = fileDescripter.get_resource_path("directory.json")
    with open(file_path, 'r', encoding='utf-8') as file:
        dir_dict = json.load(file)

        # DFS search and make Tree
        stack = deque([(dir_dict["HOME"], root, "HOME")])
        while stack:
            dict_node, node, path = stack.pop()
            for child in dict_node:
                node.children[child] = Branch(path + '/' + child, node)
                stack.append((dict_node[child], node.children[child], path + '/' + child))
    return root


# Add branch into current Branch
def add_branch(branch: Branch, add_node: str):
    try:
        # 1. Add branch
        child_path = branch.cur_path + '/' + add_node
        parent = branch
        branch.children[add_node] = Branch(child_path, parent)

        # 2. Update file json
        file_path = fileDescripter.get_resource_path("directory.json")
        data = None

        new_data = {}
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        node = data
        for node_name in branch.cur_path.split('/'):
            node = node[node_name]
        node[add_node] = {}

        with open(file_path, 'w') as file:
            json.dump(data, file)
    except Exception as e:
        pass


# Delete branch from current Branch
def delete_branch(branch: Branch, del_key: str):
    try:
        # 1. Add branch
        branch.children.pop(del_key)

        # 2. Update file json
        file_path = fileDescripter.get_resource_path("directory.json")
        data = None
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        node = data
        for node_name in branch.cur_path.split('/'):
            node = node[node_name]
        node.pop(del_key)
        with open(file_path, 'w') as file:
            json.dump(data, file)
    except Exception as e:
        print(e)
