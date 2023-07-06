# Archimedes Main Control by Blake Blount

from archimedes_gui import ArchimedesGUI

from config import Config

#from pattern_work_order_gui import PatternWorkOrderForm
#import pattern_work_order_manager as PWOM

#from pattern_receipt_gui import PatternReceiptForm
#import pattern_receipt_manager as PRM

from shop_copy_gui import ShopCopyForm
import shop_copy_manager as SCM

class ArchimedesManager:
    def __init__(self):
        self.archimedes_gui = ArchimedesGUI(self)
        self.config = Config()

#    def initialize_pattern_work_order_tool(self):
#        pattern_work_order_manager = PWOM.PatternWorkOrderManager(self.archimedes_gui.root, PatternWorkOrderForm, self.config)

#    def initialize_pattern_receipt_tool(self):
#        pattern_receipt_manager = PRM.PatternReceiptManager(self.archimedes_gui.root, PatternReceiptForm, self.config)

    def initialize_shop_copy_tool(self, event=None):
        shop_copy_manager = SCM.ShopCopyManager(self.archimedes_gui.root, ShopCopyForm, self.config)

#    def initialize_tasks_tool(self):
#        self.archimedes_gui.open_tasks_form()


if __name__ == "__main__":
    app = ArchimedesManager()
    app.archimedes_gui.root.mainloop()
