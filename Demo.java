using OpenQA.Selenium;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SeleniumCSharpMSTest.GeneralFunctions;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using OpenQA.Selenium.Interactions;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.Globalization;
using System.Linq.Expressions;
using SeleniumCSharpMSTest.ObjectRepositoryNew;
using NPOI.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Windows;
using NPOI.XSSF.UserModel;
using Syncfusion.XlsIO;
using Spire.Xls;
using SeleniumCSharpMSTest.Functions.LoginLogoutFunction;

namespace SeleniumCSharpMSTest.Functions
{
    public class CampaignBriefFunctions : Helper
    {

        public String currentStatus;
        WorkFlowObj WorkFlowObj = new WorkFlowObj();
        CampaignBriefObj CampaignbriefObj = new CampaignBriefObj();
        ClientObj ClientObj = new ClientObj();
        LoginPageObj LoginPageObj = new LoginPageObj();
        CommonObj CommonObj = new CommonObj();
      //  ContactPage ContactObj = new ContactPage();
        AdvanceSearchObj AdvanceSearchObj = new AdvanceSearchObj();
      //  AccountObj accountObj = new AccountObj();
        GenericFunctions generic = new GenericFunctions();
        CampaignObj CampaignObj = new CampaignObj();
        ProjectObj projectObj = new ProjectObj();

        string recordId;
        string acntName;
        string contactName;
        string AddressName;
        public int Elements(IWebDriver driver, By by)
        {
            int elements = driver.FindElements(by).Count;
            return elements;
        }

        public static Boolean CampaignBriefFailFlag;
        public Boolean ResetWorkFlowFailFlag()
        {
            return CampaignBriefFailFlag = false;
        }


        public Tuple<string> fillformCampaignBrief(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string rate, string day, string client, string product)
        {
            Random rnd = new Random();
            int rndint = rnd.Next(1000000);
            ThinkTime(3);
            string name = "CampaignBrief" + rndint;
            WaitUntil(driver, ClientObj.Name, 90); //fill name in new product
            Element(driver, ClientObj.Name).Click();
            ThinkTime(3);

            Element(driver, ClientObj.Name).SendKeys(name);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief name Entered", "Campaign Brief Name Entered");

            ThinkTime(2);
            WaitUntil(driver, ClientObj.ratefield, 60);
            Element(driver, ClientObj.ratefield).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.Rating(rate), 60);
            Element(driver, ClientObj.Rating(rate)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief  " + rate + " Rate Selected", "Campaign Brief  " + rate + " Rate Selected");

            
            ThinkTime(2);

            string due = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
            WaitUntil(driver, ClientObj.duedate, 30);
            Element(driver, ClientObj.duedate).Click();
            ThinkTime(2);
            Element(driver, ClientObj.duedate).Click();
            ThinkTime(2);
            Element(driver, ClientObj.duedate).SendKeys(Keys.Control + "a");
            ThinkTime(2);
            Element(driver, ClientObj.duedate).SendKeys(due);
            ThinkTime(3);

            WaitUntil(driver, ClientObj.Client, 60);
            Element(driver, ClientObj.Client).Click();
            Element(driver, ClientObj.Client).SendKeys(Keys.Enter);
            Element(driver, ClientObj.Client).SendKeys(client);
            ThinkTime(3);

            WaitUntil(driver, ClientObj.clientvalue(client), 60);
            Element(driver, ClientObj.clientvalue(client)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief  " + client + " client Selected", "Campaign Brief  " + client + " client Selected");

            ThinkTime(3);
            WaitUntil(driver, CampaignbriefObj.Product, 60);
            Element(driver, CampaignbriefObj.Product).Click();
            Element(driver, CampaignbriefObj.Product).SendKeys(Keys.Enter);
            Element(driver, CampaignbriefObj.Product).SendKeys(product);
            ThinkTime(3);

            WaitUntil(driver, CampaignbriefObj.Productvalue(product), 60);
            Element(driver, CampaignbriefObj.Productvalue(product)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief  " + product + " product Selected", "Campaign Brief  " + product + " product Selected");
            ThinkTime(3);

            return Tuple.Create(name);
        }

        public void Campaigndate(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {
            ThinkTime(2);
            string campstart = DateTime.Now.ToString("MM/dd/yyyy");
            WaitUntil(driver, CampaignObj.CMPstart, 30);
            Element(driver, CampaignObj.CMPstart).Click();
            ThinkTime(2);
            Element(driver, CampaignObj.CMPstart).Click();
            ThinkTime(2);
            Element(driver, CampaignObj.CMPstart).SendKeys(campstart);
            ThinkTime(2);
            string campend = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
            WaitUntil(driver, CampaignObj.CMPend, 30);
            Element(driver, CampaignObj.CMPend).Click();
            ThinkTime(2);
            Element(driver, CampaignObj.CMPend).Click();
            ThinkTime(2);
            Element(driver, CampaignObj.CMPend).SendKeys(Keys.Control + "a");
            ThinkTime(2);
            Element(driver, CampaignObj.CMPend).SendKeys(campend);
            ThinkTime(3);

            

            try
            {

                if (Element(driver, CampaignObj.sel_exwf).Displayed)
                {

                    ThinkTime(2);
                    MoveToElement(driver, CampaignObj.sel_exwf);
                    ThinkTime(2);

                    WaitUntil(driver, CampaignObj.Del_wf, 60);

                    ActionsClick(driver, CampaignObj.Del_wf);
                    ThinkTime(2);

                    ThinkTime(3);
                    WaitUntil(driver, ClientObj.workflowlookup, 60);
                    Element(driver, ClientObj.workflowlookup).Click();
                    Element(driver, ClientObj.workflowlookup).SendKeys(Keys.Enter);

                    ThinkTime(3);
                    WaitUntil(driver, ClientObj.lookupfirst, 60);
                    Element(driver, ClientObj.lookupfirst).Click();


                    ThinkTime(2);

                  ////  WaitUntil(driver, CampaignObj.wf_data, 60);
                  //  ActionsClick(driver, CampaignObj.wf_data);
                    //ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "info", "New wf selected", "New wf");

                }
                


            }
            catch
            {

                WaitUntil(driver, ClientObj.workflowlookup, 60);
                Element(driver, ClientObj.workflowlookup).Click();
                Element(driver, ClientObj.workflowlookup).SendKeys(Keys.Enter);

                ThinkTime(2);
                WaitUntil(driver, ClientObj.lookupfirst, 60);
                Element(driver, ClientObj.lookupfirst).Click();
                ThinkTime(2);

                AddLog(driver, testInReport, testName, testDataIteration, "info", "New wf is not selected", "New wf");

            }
            generic.ClickSaveAndClose(driver, testInReport, testName, testDataIteration);

            ThinkTime(2);

        }


        public void GoToUser(IWebDriver driver, ExtentTest extentTest, string testName, string testDataIteration, string member)
        {

            ThinkTime(3);
            WaitUntil(driver, CampaignbriefObj.morecommand, 60);
            Element(driver, CampaignbriefObj.morecommand).Click();
            

            ThinkTime(3);
            WaitUntil(driver, ClientObj.adduser, 60);
            Element(driver, ClientObj.adduser).Click();
            ThinkTime(2);

            WaitUntil(driver, ClientObj.lookup, 60);
            Element(driver, ClientObj.lookup).Click();
            Element(driver, ClientObj.lookup).SendKeys(Keys.Enter);
            Element(driver, ClientObj.lookup).SendKeys(member);
            ThinkTime(4);

            WaitUntil(driver, ClientObj.lookupval(member), 60);
            Element(driver, ClientObj.lookupval(member)).Click();

            ThinkTime(4);
            WaitUntil(driver, ClientObj.add, 60);
            Element(driver, ClientObj.add).Click();
            ThinkTime(5);

            AddLog(driver, extentTest, testName, testDataIteration, "Info", "Agency Contact Added", " Agency Contact Added");
        }

        public void RelatedCampaign(IWebDriver driver, ExtentTest extentTest, string testName, string testDataIteration)
        {

            ThinkTime(6);
            WaitUntil(driver, CampaignbriefObj.related, 60);
            Element(driver, CampaignbriefObj.related).Click();


            ThinkTime(3);
            WaitUntil(driver, CampaignbriefObj.campaigns, 60);
            Element(driver, CampaignbriefObj.campaigns).Click();


            ThinkTime(3);
            WaitUntil(driver, CampaignbriefObj.newcampaigns, 60);
            Element(driver, CampaignbriefObj.newcampaigns).Click();

        }


        public void Campaign(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string rate, string startday, string enddate, string client)
        {

            Random rnd = new Random();
            int rndint = rnd.Next(1000000);
            ThinkTime(4);
            string name = "CampaignBrief" + rndint;
            WaitUntil(driver, ClientObj.Name, 60); //fill name in new product
            Element(driver, ClientObj.Name).Click();
            Element(driver, ClientObj.Name).SendKeys(name);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief name Entered", "Campaign Brief Name Entered");

            ThinkTime(2);
            WaitUntil(driver, ClientObj.ratefield, 60);
            Element(driver, ClientObj.ratefield).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.Rating(rate), 60);
            Element(driver, ClientObj.Rating(rate)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief  " + rate + " Rate Selected", "Campaign Brief  " + rate + " Rate Selected");

            ThinkTime(2);
            WaitUntil(driver, ClientObj.sdate, 60);
            Element(driver, ClientObj.sdate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(startday), 60);
            JSClick(driver, ClientObj.datenew(startday));

            ThinkTime(2);
            WaitUntil(driver, ClientObj.edate, 60);
            Element(driver, ClientObj.edate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(enddate), 60);
            JSClick(driver, ClientObj.datenew(enddate));

            WaitUntil(driver, ClientObj.Client, 60);
            Element(driver, ClientObj.Client).Click();
            Element(driver, ClientObj.Client).SendKeys(Keys.Enter);
            Element(driver, ClientObj.Client).SendKeys(client);

            WaitUntil(driver, ClientObj.clientvalue(client), 60);
            Element(driver, ClientObj.clientvalue(client)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief  " + client + " client Selected", "Campaign Brief  " + client + " client Selected");

            ThinkTime(3);
            WaitUntil(driver, ClientObj.workflowlookup, 60);
            Element(driver, ClientObj.workflowlookup).Click();
            Element(driver, ClientObj.workflowlookup).SendKeys(Keys.Enter);

            ThinkTime(2);
            WaitUntil(driver, ClientObj.lookupfirst, 60);
            Element(driver, ClientObj.lookupfirst).Click();


        }


       

       

        public  Tuple<string>  FillAppointment(IWebDriver driver, ExtentTest extentTest, string testName, string testDataIteration, string date, string starttimme, string endtime)
        {
            Random rnd = new Random();
            int rndint = rnd.Next(1000000);
            ThinkTime(2);
            string subject = "Subject" + rndint;
            WaitUntil(driver, ClientObj.subject, 60); 
            Element(driver, ClientObj.subject).Click();
            ThinkTime(2);
            Element(driver, ClientObj.subject).SendKeys(Keys.Enter);

            Element(driver, ClientObj.subject).SendKeys(subject);

            ThinkTime(2);
            /*WaitUntil(driver, CampaignbriefObj.starttimedate, 60);
            generic.scrollToElement(driver, CampaignbriefObj.starttimedate);
            Element(driver, CampaignbriefObj.starttimedate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(date), 60);
            JSClick(driver, ClientObj.datenew(date));*/

            //string start = DateTime.Now.ToString("MM/dd/yyyy");
            //WaitUntil(driver, CampaignbriefObj.starttimedate, 30);
            //Element(driver, CampaignbriefObj.starttimedate).Click();
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.starttimedate).Click();
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.starttimedate).SendKeys(Keys.Control + "a");
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.starttimedate).SendKeys(start);
            //ThinkTime(3);

            //string end = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
            //WaitUntil(driver, CampaignbriefObj.endtimedate, 30);
            //Element(driver, CampaignbriefObj.endtimedate).Click();
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.endtimedate).Click();
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.endtimedate).SendKeys(Keys.Control + "a");
            //ThinkTime(2);
            //Element(driver, CampaignbriefObj.endtimedate).SendKeys(start);
            //ThinkTime(3);
            /*WaitUntil(driver, CampaignbriefObj.endtimedate, 60);
            Element(driver, CampaignbriefObj.endtimedate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(date), 60);
            JSClick(driver, ClientObj.datenew(date));

            ThinkTime(3);*/
            WaitUntil(driver, CampaignbriefObj.allday, 60);
            Element(driver, CampaignbriefObj.allday).Click();
            ThinkTime(3);


            return Tuple.Create(subject);
        }
        public void CloseAppointment(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string confirm)
        {



            ThinkTime(4);
            WaitUntil(driver, CampaignbriefObj.closeappointment, 30);
            Element(driver, CampaignbriefObj.closeappointment).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Pass", " button is clicked", "HeaderButton");



            ThinkTime(3);
            WaitUntil(driver, CampaignbriefObj.OK(confirm), 30);
            Element(driver, CampaignbriefObj.OK(confirm)).Click();

            ThinkTime(3);


        }

        public void Conflict(IWebDriver driver, ExtentTest extentTest, string testName, string testDataIteration)
        {  
            if (Element(driver, CampaignbriefObj.conflict).Displayed)
            {
                ThinkTime(3);
                WaitUntil(driver, CampaignbriefObj.ignore, 90);
                Element(driver, CampaignbriefObj.ignore).Click();
                ThinkTime(2);

            }

        }



        
        public Tuple<string> Campaign(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string rate, string startday, string enddate, string client, string wfname)
        {



            Random rnd = new Random();
            int rndint = rnd.Next(1000000);
            ThinkTime(2);
            string name = "CampaignBrief" + rndint;
            WaitUntil(driver, ClientObj.Name, 60); //fill name in new product
            Element(driver, ClientObj.Name).Click();
            ThinkTime(4);
            Element(driver, ClientObj.Name).SendKeys(Keys.Enter);

            Element(driver, ClientObj.Name).SendKeys(name);
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief name Entered", "Campaign Brief Name Entered");



            ThinkTime(3);
            WaitUntil(driver, ClientObj.ratefield, 60);
            Element(driver, ClientObj.ratefield).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.Rating(rate), 60);
            Element(driver, ClientObj.Rating(rate)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief " + rate + " Rate Selected", "Campaign Brief " + rate + " Rate Selected");



            ThinkTime(3);
            WaitUntil(driver, ClientObj.sdate, 60);
            Element(driver, ClientObj.sdate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(startday), 60);
            JSClick(driver, ClientObj.datenew(startday));



            ThinkTime(3);
            WaitUntil(driver, ClientObj.edate, 60);
            Element(driver, ClientObj.edate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(enddate), 60);
            JSClick(driver, ClientObj.datenew(enddate));



            WaitUntil(driver, ClientObj.Client, 60);
            Element(driver, ClientObj.Client).Click();
            Element(driver, ClientObj.Client).SendKeys(Keys.Enter);
            Element(driver, ClientObj.Client).SendKeys(client);



            WaitUntil(driver, ClientObj.clientvalue(client), 60);
            Element(driver, ClientObj.clientvalue(client)).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Info", "Campaign Brief " + client + " client Selected", "Campaign Brief " + client + " client Selected");



            try
            {

                if (Element(driver, CampaignObj.sel_exwf).Displayed)
                {

                    ThinkTime(2);
                    MoveToElement(driver, CampaignObj.sel_exwf);

                    WaitUntil(driver, CampaignObj.Del_wf, 60);

                    ActionsClick(driver, CampaignObj.Del_wf);
                    ThinkTime(2);

                    WaitUntil(driver, CampaignObj.WF_field, 60);
                    ActionsClick(driver, CampaignObj.WF_field);
                    ThinkTime(2);
                    ActionSendKeys(driver, CampaignObj.WF_field, wfname);

                    ThinkTime(2);

                    WaitUntil(driver, CampaignObj.wf_data, 60);
                    ActionsClick(driver, CampaignObj.wf_data);
                    ThinkTime(2);
                    AddLog(driver, testInReport, testName, testDataIteration, "info", "New wf selected", "New wf");

                }
               

            }
            catch
            {
                WaitUntil(driver, WorkFlowObj.wf_lookup, 60);
                ActionsClick(driver, WorkFlowObj.wf_lookup);
                ThinkTime(2);
                WaitUntil(driver, WorkFlowObj.wf_lookup, 60);
                ActionSendKeys(driver, WorkFlowObj.wf_lookup, wfname);
                ThinkTime(3);

                WaitUntil(driver, WorkFlowObj.select_wf, 60);
                ActionsClick(driver, WorkFlowObj.select_wf);
                ThinkTime(2);
               generic.ClickSaveAndClose(driver, testInReport, testName, testDataIteration);


                AddLog(driver, testInReport, testName, testDataIteration, "info", "New wf is not selected", "New wf");

                ThinkTime(5);

            }



            return Tuple.Create(name);



        }

        public void CheckInactiveState(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {
            ThinkTime(4);

            try
            {

                if (Element(driver, CommonObj.Activatebtn).Displayed)
                {
                    MoveToElement(driver, CommonObj.Activatebtn);
                    generic.Activate(driver, testInReport, testName, testDataIteration);

                }
            }
            catch
            {
                
            }

            ThinkTime(3);

        }

        public void CloseAppoinment(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration)
        {



            ThinkTime(4);
            WaitUntil(driver, CampaignbriefObj.closeappointment, 80);
            Element(driver, CampaignbriefObj.closeappointment).Click();
            AddLog(driver, testInReport, testName, testDataIteration, "Pass", " button is clicked", "HeaderButton");



            ThinkTime(5);
            WaitUntil(driver, CampaignbriefObj.closeapp, 80);
            Element(driver, CampaignbriefObj.closeapp).Click();

            ThinkTime(3);


        }

        
    public void Campaigndate(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string startday, string enddate, string workflow)
        {
            ThinkTime(4);
            WaitUntil(driver, ClientObj.sdate, 60);
            Element(driver, ClientObj.sdate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(startday), 60);
            JSClick(driver, ClientObj.datenew(startday));



            ThinkTime(3);
            WaitUntil(driver, ClientObj.edate, 60);
            Element(driver, ClientObj.edate).Click();
            ThinkTime(2);
            WaitUntil(driver, ClientObj.datenew(enddate), 60);
            JSClick(driver, ClientObj.datenew(enddate));
            try
            {
                ThinkTime(6);
                bool b = false;
                b = Element(driver, projectObj.workflowname).Displayed;
                if (b)
                { }
            }
            catch
            {
                ThinkTime(3);
                WaitUntil(driver, projectObj.Workflow, 60);
                Element(driver, projectObj.Workflow).Click();

                //Element(driver, projectObj.Workflow).SendKeys(Keys.Enter);
                //Element(driver, projectObj.Workflow).SendKeys(workflow); ThinkTime(3);
                //WaitUntil(driver, projectObj.Workflowvalue(workflow), 60);
                //Element(driver, projectObj.Workflowvalue(workflow)).Click();

                WaitUntil(driver, projectObj.wf_search, 60);
                Element(driver, projectObj.wf_search).Click();
                ThinkTime(2);

                WaitUntil(driver, projectObj.wf_sel, 60);
                Element(driver, projectObj.wf_sel).Click();
                ThinkTime(2);
            }
        }

        
    public void Getdetails(IWebDriver driver, ExtentTest testInReport, string testName, string testDataIteration, string details)
        {
            ThinkTime(3);
            generic.ScrollPage(driver, testInReport, testName, testDataIteration, "500");
            ThinkTime(5);
            WaitUntil(driver, CampaignbriefObj.briefdetails, 60);
            Element(driver, CampaignbriefObj.briefdetails).Click();
            ThinkTime(3);
            string content = Element(driver, CampaignbriefObj.briefdetails).Text; Assert.AreEqual(details, content); AddLog(driver, testInReport, testName, testDataIteration, "Pass", " The global client workflow " + content + " is selected", "Pass ");
        }
        public Tuple<string> Additionaldetail(IWebDriver driver, ExtentTest extentTest, string testName, string testDataIteration)
        {
            ThinkTime(3);
            string name = Element(driver, CampaignbriefObj.titlename).GetAttribute("title");
            string details = "Brief details for Campaign Brief: " + name;
            WaitUntil(driver, CampaignbriefObj.details, 60);
            Element(driver, CampaignbriefObj.details).Click();



            ThinkTime(3);
            string content = Element(driver, CampaignbriefObj.briefdetails).Text;



            if (content == "---")
            {
                ThinkTime(3);
                WaitUntil(driver, CampaignbriefObj.briefdetails, 60);
                Element(driver, CampaignbriefObj.briefdetails).Click();
                Element(driver, CampaignbriefObj.briefdetails).SendKeys(details);
            }
            else
            {



            }
            //}



            return Tuple.Create(details);
        }

    }


}


