# -*- coding: utf-8 -*-
"""
Created on Sun Feb 16 12:36:10 2020

@author: Cesar
"""

import wx
###############################################################################

class CustomNumValidator(wx.Validator):
    """ Validator for entering custom low and high limits """
    print('creado')
    def __init__(self):
        super(CustomNumValidator, self).__init__()


    # --------------------------------------------------------------------------
    def Clone(self):
        """ """
        return CustomNumValidator()

    # --------------------------------------------------------------------------
    def Validate(self, win):
        """ """
        textCtrl = self.GetWindow()
        text = textCtrl.GetValue()
        print('pasando')
        if text.isdigit():
            return True
        else:
            wx.MessageBox("Please enter numbers only", "Invalid Input",
            wx.OK | wx.ICON_ERROR)
        return False

    # --------------------------------------------------------------------------
    def TransferToWindow(self):
        return True

    # --------------------------------------------------------------------------
    def TransferFromWindow(self):
        return True


class CustomNumbers(wx.Frame):
    """ Dialog for choosing custom numbers """
    def __init__(self, *args, **kwargs):
        super(CustomNumbers, self).__init__(*args, **kwargs)

        wx.Frame.__init__(self, None, wx.ID_ANY, "Software Legal")
        
        self.SetBackgroundColour("WHITE")
        self.widget_dict = {}

        self.initUI()
        self.SetSizerAndFit(self.main_sizer)
        
        self.InitDialog()
        self.Layout()
        self.Refresh()

    # --------------------------------------------------------------------------
    def initUI(self):
        """ """
        self.createSizer()
        self.createText()
        self.createInputBox()
        self.createButton()
        self.addSizerContent()

    # --------------------------------------------------------------------------
    def createSizer(self):
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

    # --------------------------------------------------------------------------
    def createText(self):
        """ """
        
        low_num_text = wx.StaticText(self, -1, "Low Number")
        high_num_text = wx.StaticText(self, -1, "High Number")

        self.widget_dict["low_num_text"] = low_num_text
        self.widget_dict["high_num_text"] = high_num_text

    # --------------------------------------------------------------------------
    def createInputBox(self):
        """ """
        low_input = wx.TextCtrl(self, validator=CustomNumValidator())
        high_input = wx.TextCtrl(self, validator=CustomNumValidator())
        self.widget_dict["low_input"] = low_input
        self.widget_dict["high_input"] = high_input

    # --------------------------------------------------------------------------
    def createButton(self):
        """ """
        ok_btn = wx.Button(self, wx.ID_OK, "Enter")
        cancel_btn = wx.Button(self, wx.ID_CANCEL, "Cancel")
        
        ok_btn.Bind(wx.EVT_BUTTON, self.onbutton)

        self.widget_dict["ok_btn"] = ok_btn
        self.widget_dict["cancel_btn"] = cancel_btn

    def onbutton(self,event):
        print('si')
        self.InitDialog()
        CustomNumValidator().Validate
        
    
    # --------------------------------------------------------------------------
    def addSizerContent(self):
        """ """
        top_sizer = wx.BoxSizer()
        top_sizer.Add(self.widget_dict["low_num_text"], 3, wx.ALL, 10)
        top_sizer.Add(self.widget_dict["low_input"], 7, wx.ALL ^ wx.RIGHT, 10)

        btm_sizer = wx.BoxSizer()
        btm_sizer.Add(self.widget_dict["high_num_text"], 3, wx.ALL, 10)
        btm_sizer.Add(self.widget_dict["high_input"], 7, wx.ALL, 10)

        btn_sizer = wx.BoxSizer()
        btn_sizer.Add(self.widget_dict["ok_btn"], 0, wx.CENTER | wx.ALL, 10)
        btn_sizer.Add(self.widget_dict["cancel_btn"], 0,
                      wx.CENTER | wx.ALL, 10)

        self.main_sizer.Add(top_sizer)
        self.main_sizer.Add(btm_sizer)
        self.main_sizer.Add(btn_sizer, 0, wx.CENTER | wx.ALL, 10)

    # --------------------------------------------------------------------------
    def getValues(self):
        """ """


###############################################################################


class MyApp(wx.App):
    def OnInit(self):
        self.frame= CustomNumbers()
        self.frame.Show()
        return True  

# Run the program     
app=MyApp()
app.MainLoop()
del app