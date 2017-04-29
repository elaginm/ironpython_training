import clr
import sys
sys.path.append("C:\\Project\\ironpython_training\\TestStack.White.0.13.3\\lib\\net40\\")
sys.path.append("C:\\Project\\ironpython_training\\Castle.Core.3.3.0\\lib\\net40-client\\")
clr.AddReferenceByName('TestStack.White')

from TestStack.White import Application
from TestStack.White.InputDevices import Keyboard
from TestStack.White.WindowsAPI import KeyboardInput
from TestStack.White.UIItems.Finders import *

clr.AddReferenceByName('UIAutomationTypes, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
from System.Windows.Automation import *
from fixture.group import GroupHelper

class Application1:

    def __init__(self):
        self.group = GroupHelper(self)
        self.application = Application.Launch("c:\\Project\\AddressBook\\AddressBook.exe")
        self.main_window = self.application.GetWindow("Free Address Book")

    def close_app(self):
        self.main_window.Get(SearchCriteria.ByAutomationId("uxExitAddressButton")).Click()
