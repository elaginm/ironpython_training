import clr
import sys
import os.path

project_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
sys.path.append(os.path.join(project_dir, "TestStack.White.0.13.3\\lib\\net40\\"))
sys.path.append(os.path.join(project_dir, "Castle.Core.3.3.0\\lib\\net40-client\\"))
clr.AddReferenceByName('TestStack.White')

from TestStack.White import Application
from TestStack.White.InputDevices import Keyboard
from TestStack.White.WindowsAPI import KeyboardInput
from TestStack.White.UIItems.Finders import *

clr.AddReferenceByName('UIAutomationTypes, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *


class GroupHelper:
    def __init__(self, app):
        self.app = app

    def open_group_editor(self, main_window):
        main_window.Get(SearchCriteria.ByAutomationId("groupButton")).Click()
        modal = main_window.ModalWindow("Group editor")

        return modal

    def close_group_editor(self, modal):
        modal.Get(SearchCriteria.ByAutomationId("uxCloseAddressButton")).Click()

    group_cache = None

    def get_group_list(self, main_window):
        if self.group_cache is None:
            modal = self.open_group_editor(main_window)
            self.group_cache = []
            tree = modal.Get(SearchCriteria.ByAutomationId("uxAddressTreeView"))
            root = tree.Nodes[0]
            self.group_cache = [node.Text for node in root.Nodes]
            self.close_group_editor(modal)
        return list(self.group_cache)

    def add_new_group(self, main_window, group):
        modal = self.open_group_editor(main_window)
        modal.Get(SearchCriteria.ByAutomationId("uxNewAddressButton")).Click()
        modal.Get(SearchCriteria.ByControlType(ControlType.Edit)).Enter(group.name)
        Keyboard.Instance.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN)
        self.close_group_editor(modal)
        self.group_cache = None

    def delete_group_by_index(self, main_window, index):
        group = self.open_group_editor(main_window)
        tree = group.Get(SearchCriteria.ByAutomationId("uxAddressTreeView"))
        root = tree.Nodes[0]
        root.Nodes[index].Select()
        group.Get(SearchCriteria.ByAutomationId("uxDeleteAddressButton")).Click()
        delete_group = group.ModalWindow("Delete group")
        delete_group.Get(SearchCriteria.ByAutomationId("uxOKAddressButton")).Click()
        self.close_group_editor(group)
        self.group_cache = None