﻿using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.DirectoryServices.AccountManagement;




namespace PrintServerGroups
{

     /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Global for list of users. Stored here because of the nested groups
        //TODO: See if there is a better way to store this information without using a global?
        private List<string> userList = new List<string>();
        #region Get Group List

        //GetAllGroups will query the domain's Active Directory and pull a list of all of the groups found on the server.
        //Groups will be sorted and added to the combobox by their display name
        public void getAllGroups()
        {
            List<string> groupNames = new List<string>();
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            GroupPrincipal group = new GroupPrincipal(ctx);

            PrincipalSearcher pSearcher = new PrincipalSearcher(group);

            //Find all groups and store Display name into List
            foreach (var found in pSearcher.FindAll())
            {
                GroupPrincipal foundGroup = found as GroupPrincipal;

                if (foundGroup != null)
                {
                    groupNames.Add(foundGroup.DisplayName);
                }
            }

            groupNames.Sort();

            foreach (string s in groupNames)
            {
                if (s != null)
                {
                    cmbxADGroups.Items.Add(s);
                }
            }
        }
        #endregion

        #region Get Group Members from Group
        //Pulls all users from the selected group in the combobox. If there are groups inside the groups, pull their members too.
        //Check if the list already has the user listed in case they are in both the main group and included groups.
        //Sort the list alphabetically. 
        public List<string> getUsers(string groupName)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, groupName);

            if (group != null)
            {
                foreach (Principal p in group.GetMembers())
                {
                    if(p.StructuralObjectClass == "group")
                    {
                        //Pulls members from nested groups
                        getUsers(p.DisplayName);
                    }
                    else
                    {
                        if (!userList.Contains(p.DisplayName))
                        {
                            userList.Add(p.DisplayName);
                        }
                    }
                }
            }

            return userList;
        }
        #endregion

        #region Main
        public MainWindow()
        {
            InitializeComponent();

            getAllGroups();

           
        }
        #endregion

        #region Group Selection Changed
        private void cmbxADGroups_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Clear out listbox and update the list with new group selection
            if (!listBoxGroupMembers.Items.IsEmpty)
            {
                listBoxGroupMembers.Items.Clear();
                userList.Clear();
            }

            List<string> memberList = new List<string>();
            memberList = getUsers(cmbxADGroups.SelectedItem.ToString());
            memberList.Sort();

            //Do not include "like" roles. Ex. Like Kim (Accountant)
            foreach (string s in memberList)
            {
                if (s != null && !s.Contains("Like"))
                {
                    listBoxGroupMembers.Items.Add(s);
                }
            }
            memberList.Clear();
        }
        #endregion  

        #region Print Document
        public void printDocument(StringBuilder Text)
        {
            //Create print dialog box, if yes print the list. If no, don't do anything. String passed in from Print Button Click method.
            PrintDialog pd = new PrintDialog();
            FlowDocument fd = new FlowDocument(new Paragraph(new Run(Text.ToString())));
            fd.Name = "GroupMembers";
            fd.PagePadding = new Thickness(40);

            IDocumentPaginatorSource idpSource = fd;

            if (pd.ShowDialog().GetValueOrDefault(false))
            {
                pd.PrintDocument(idpSource.DocumentPaginator, "Group Members");
            }
        }
        #endregion

        #region Print Button
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            //Create String from all items in list box. Pass the string to the print method.
            StringBuilder textString = new StringBuilder();
            textString.AppendLine("Users in Group " + cmbxADGroups.SelectedItem.ToString() + ":");
            textString.AppendLine("  ");
            foreach (string s in listBoxGroupMembers.Items)
            {
                if(s != null)
                {
                    textString.AppendLine(s);
                }
            }
            printDocument(textString);
        }
        #endregion  

        #region Export Button
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            //Create an instance of Microsoft Excel and open it.
            var exApp = new Microsoft.Office.Interop.Excel.Application();
            exApp.Visible = true;
            exApp.Workbooks.Add();

            //Neat formatting.
            var row = 3;
            exApp.Cells[1, "A"] = cmbxADGroups.SelectedItem.ToString() + " Group Members";
            exApp.Cells[3, "A"] = "First Name";
            exApp.Cells[3, "B"] = "Last Name";

            exApp.Range["A1", "I1"].Merge();
            exApp.Range["A1"].Font.Size = 16;
            exApp.Range["A1"].Font.Bold = true;

            exApp.Range["A3, B3"].Font.Bold = true;

            char delimiter = ' ';

            //Loop through each item in list box. Split names at the space. Add names to worksheet columns. Auto fit the columns.
            foreach(string s in listBoxGroupMembers.Items)
            {
                row++;
                if (s != null)
                {
                    string[] splitString = s.Split(delimiter);
                        exApp.Cells[row, "A"] = splitString[0];
                    if(splitString.Length == 2)
                        exApp.Cells[row, "B"] = splitString[1];
                }
            }

            exApp.Columns.AutoFit(); 
        }
        #endregion
    }
}