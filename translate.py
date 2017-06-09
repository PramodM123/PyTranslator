
import xlrd
import re

inp="./EventProperties_5123.xls"
out="./TestExample.cs"

class trans:
    xls=None
    DEBUG=True
    row=0
    sheet=None
    applicationName="ABSuite_Test on (local)\ABSuite_Test"
    out=None

    def __init__(self, init, end, excelFile, cSharpFile):
        self.xls=xlrd.open_workbook(excelFile,on_demand=True)

        if cSharpFile:
            self.out=open(cSharpFile,"w")
        if init:
            self.writeToFile(init)
        if end:
            self.end=end

    def writeToFile(self, data):
        if self.out:
            self.out.write(data)
            self.out.write("\n")

    def sheet_list(self):
        return self.xls.sheet_names()

    def start(self, sheet=None):
        sheets=[]
        if sheet != None:
            sheets+=[sheet]
        else:
            sheets+=self.xls.sheet_names()

        for s in sheets:
            self.sheet=self.xls.sheet_by_name(s)
            data=""
            #try:
            while not "Test execution complete" in str(self.sheet.cell_value(self.row,2)):
            #for data in sheet.col(2):
                #self.row+=1;
                #print data
                flag=self.parse_data()
                #if not flag:
                #   print self.sheet.cell_value(self.row,2),self.sheet.cell_value(self.row,3)
                    #print data
            #except:
            self.debug("{} parsing problem in {} sheet".format(data,s))
            self.xls.unload_sheet(s)

    def parse_data(self):
        obj,action=self.getNextData()

        flag=False
        #p=re.match(r"(?P<type>.*):(?P<val>.*)", d )
        #if p.group('type') == 'text':

        #data=p.group('val')[2:-1]

        #Kill Developer
        if re.match("KillDeveloper",obj):
            print "devFun.DeleteModelData();"
            self.writeToFile("devFun.DeleteModelData();")
            flag=True

        #devFun.Connect_To_ExistingModel();
        if re.match("Open_ExistingProject", obj):
            print "devFun.Connect_To_ExistingModel();"
            self.writeToFile("b=devFun.Connect_To_ExistingModel();")
            self.writeToFile('if (b == 1)')
            self.writeToFile('      resultobj.logfilewrite('+'Msg - Connected to project'+');')
            self.writeToFile('else')
            self.writeToFile('      throw new Exception("Connected to project not happened'+');')
            flag=True

        #select_projectnode
        if re.match("select_projectnode", obj):
            print 'devFun.SelectClassViewItem(@"{}");'.format(self.applicationName)
            self.writeToFile('devFun.SelectClassViewItem(@"{}");'.format(self.applicationName))
            flag=True
##            d1,e1=self.getNextData()
##            p=re.match("window;Application=(?P<application>.*) Caption='(?P<caption>.*)'",str(d1))
##            if p:
##                self.applicationName=p.group('caption')
##                print 'devFun.SelectClassViewItem(@"ABSuite_Test on (local)\{}");'.format(p.group('caption'))
##                flag=True

##        if re.match("textbox;Name=nameTextBox", obj):
##            p=re.match("Settext::(?P<segmentName>.*)", action)
##            if p:
##                print 'Winop.AddElemntDev(@"{}", "Segment", "{}");'.format(self.applicationName,p.group('segmentName'))
##                self.writeToFile('Winop.AddElemntDev(@"{}", "Segment", "{}");'.format(self.applicationName,p.group('segmentName')))
##                flag=True

        if re.match("guiobject;VsClassViewTypesPane GUIObject", obj):
            p=re.match("textselect::(?P<eventName>.*)", action)
            if p:
                print 'devFun.SelectClassViewItem("{}");'.format(p.group('eventName'))
                self.writeToFile('devFun.SelectClassViewItem("{}");'.format(p.group('eventName')))
                flag=True

        if re.match("window;ABSuite_Test - Microsoft Visual Studio Window", obj):
            p=re.match("type::{(?P<key>.*)}", action)
            if p:
                print 'Keyboard.SendKeys("{}");'.format('{'+p.group('key')+'}')
                self.writeToFile('Keyboard.SendKeys("{}");'.format('{'+p.group('key')+'}'))
                flag=True

        if re.match("Propertygrid;Parent.Caption=PropertyGrid", obj):
            p=re.match("select::(?P<item>.*)",action)
            if p:
                print 'Winro = Winop.MemItem("{}");'.format(p.group('item'))
                self.writeToFile('Winro = Winop.MemItem("{}");'.format(p.group('item')))
                self.writeToFile('if ((Winro.Value).Equals(" "))')
                self.writeToFile('      resultobj.logfilewrite("Success - " + Winro.Name + " exists and value is " + Winro.Value);')
                self.writeToFile(' else')
                self.writeToFile('      resultobj.logfilewrite("ERRMSG - " + Winro.Name + " exists and value " + Winro.Value + " is incorrect");')
                flag=True
                
        if re.match("listview;Name=templateListView", obj):
            p=re.match("Select::(?P<selection>.*)", action)
            if p:
                d1,e1=self.getNextData()
                if re.match("textbox;Name=nameTextBox",d1):
                    p2=re.match("Settext::(?P<segmentName>.*)",e1)
                    print 'Winop.AddElemntDev(@"{}", "{}", "{}");'.format(self.applicationName, p.group('selection'),p2.group('segmentName'))
                    self.writeToFile('Winop.AddElemntDev(@"{}", "{}", "{}");'.format(self.applicationName, p.group('selection'),p2.group('segmentName')))
                    flag=True

        if not flag:
            print obj,"->",action
            self.writeToFile("//" + obj + " -> " + action)

        return flag

    def getNextData(self):
        obj=str(self.sheet.cell_value(self.row,2))
        action=str(self.sheet.cell_value(self.row,3))

        #print "getNextData -> " , obj, action
        p=re.match(r"(?P<type>.*):(?P<val>.*)", obj )
        if p:
            obj=p.group('val')

        #p=re.match(r"(?P<type>.*):(?P<val>.*)", action )
        #if p:
        #   action=p.group('val')

        self.row+=1
        return [str(obj),str(action)]

    def close(self):
        if self.end:
            self.writeToFile(self.end)
        self.out.close()

    def debug(self,arg):
        if self.DEBUG:
            print arg



def main():

    init='''
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using messagebox = System.Windows.Forms.MessageBox;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using System.IO;
using System.Collections;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using MouseButtons = System.Windows.Forms.MouseButtons;
using Developer_funcs;
using ABSuiteAutomation.Developer;
using System.Threading;
using WINFORMFUNC;
using CommonFunction;
using System.Diagnostics;
using CLRAdminUtility;
using time = System.Timers;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;


namespace ABSuiteAutomation.Scripts.CLR.Developer
{
    /// <summary>
    ///
    /// </summary>
    [CodedUITest]
    public class Sample
    {
        public Sample()
        {
        }

        public static int setflag;
        public static int setTestflag;

        [TestMethod]
        public void Sample_CLR()
        {

            /* *************************************************************************************

          Test Objective: 
          * Scripter: Pramod Mekapothula
           * Date: 1-2-2015
          * *************************************************************************************
          */
            ResultLog resultObject = new ResultLog();
            resultObject.Resultlog_txt("Samplelog_CLR");
            resultObject.logfilewrite("Msg - Started Logging");

            AllDevFuncs devFun = new AllDevFuncs();
            WINFORM_FUNCTIONS Winop = new WINFORM_FUNCTIONS();
            WinRow Winro = new WinRow();
             WinRow Winro1 = new WinRow();
            WinTreeItem Wintr = new WinTreeItem();
            WinWindow Winwi = new WinWindow();
            int b;
            try
            {

    '''
    end='''
               resultObject.logfileclose();
                resultObject.FinalResult("Samplelog_CLR");


            }


            catch
            {
                resultObject.logfilewrite("ERRMSG - Script Failed");
                resultObject.logfileclose();
                resultObject.FinalResult("Samplelog_CLR");
                devFun.killDeveloper();
            }

        }
    }
}
    '''
      
    t=trans(init, end, inp,out)
    l=t.sheet_list()
    t.start(l[0])
    t.close()

if __name__ == '__main__':
    main()
