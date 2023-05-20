using SimioAPI;
using SimioAPI.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AutoCADCreateLinks
{
    internal class AutoCADCreateTables : IDesignAddIn, IDesignAddInGuiDetails
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Create Tables"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Create Tables."; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.CreateTables; }
        }

        /// <summary>
        /// Method called when the add-in is run.
        /// </summary>
        public void Execute(IDesignContext context)
        {
            // This example code places some new objects from the Standard Library into the active model of the project.
            if (context.ActiveModel != null)
            {
                MessageBox.Show("Create Table Not Implemented", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);         
            }
        }

        #endregion

        #region IDesignAddInGuiDetails Members

        public string CategoryName
        {
            get { return "Table Tools"; }
        }

        public string TabName
        {
            get { return "AutoCAD"; }
        }

        public string GroupName
        {
            get { return "Intergration"; }
        }

        #endregion
    }

    internal class AutoCADCreateLinks : IDesignAddIn, IDesignAddInGuiDetails
    {
        #region IDesignAddIn Members

        /// <summary>
        /// Property returning the name of the add-in. This name may contain any characters and is used as the display name for the add-in in the UI.
        /// </summary>
        public string Name
        {
            get { return "Create Links"; }
        }

        /// <summary>
        /// Property returning a short description of what the add-in does.
        /// </summary>
        public string Description
        {
            get { return "Create Links."; }
        }

        /// <summary>
        /// Property returning an icon to display for the add-in in the UI.
        /// </summary>
        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.CreateLinks; }
        }

        /// <summary>
        /// Method called when the add-in is run.
        /// </summary>
        public void Execute(IDesignContext context)
        {
            // This example code places some new objects from the Standard Library into the active model of the project.
            if (context.ActiveModel != null)
            {
                context.ActiveModel.BulkUpdate(model =>
                {
                    var table = context.ActiveModel.Tables["AutoCADExport"];
                    var dataTable = AutoCADCreateLinksUtils.ConvertTableToDataTable(table);

                    DataView dataView = new DataView(dataTable);//datatable to dataview
                    dataView.Sort = "Layer ASC, Segment ASC, Sequence ASC";//string that contains the column name  followed by "ASC" (ascending) or "DESC" (descending)
                    dataTable = dataView.ToTable();//push the chages back to the datatable;

                    // get tables
                    var nodesTable = context.ActiveModel.Tables["AutoCADNodes"];
                    var linksTable = context.ActiveModel.Tables["AutoCADLinks"];
                    var verticesTable = context.ActiveModel.Tables["AutoCADVertices"];

                    // clear tables
                    nodesTable.Rows.Clear();
                    linksTable.Rows.Clear();
                    verticesTable.Rows.Clear();
                   
                    string lastSegment = String.Empty;
                    string lastSequence = String.Empty;
                    string startNodeName = String.Empty;
                    string endNodeName = String.Empty;
                    string linkName = String.Empty;
                    int numberOfRows = dataTable.Rows.Count;
                    int rowIdx = 0;
                    var verticesList = new List<string[]>();
                    foreach (DataRow row in dataTable.Rows) 
                    {
                        if (lastSegment != row["Segment"].ToString())
                        {
                            linkName = row["Layer"].ToString() + row["Segment"].ToString();
                            linkName.Replace(" ", "_");
                            linkName.Replace(".", "_");
                            startNodeName = row["Layer"].ToString() + "_Start_" + row["Segment"].ToString() + "_" + row["Sequence"].ToString();
                            startNodeName.Replace(" ", "_");
                            startNodeName.Replace(".", "_");

                            // add start row
                            var nodeRow = nodesTable.Rows.Create();
                            nodeRow.Properties["Node"].Value = startNodeName;
                            nodeRow.Properties["XLoc"].Value = row["StartX"].ToString();
                            nodeRow.Properties["YLoc"].Value = row["StartZ"].ToString();
                            nodeRow.Properties["ZLoc"].Value = row["StartY"].ToString();
                        }
                        else
                        {
                            string[] startVertexArray = { linkName, row["Sequence"].ToString(), row["StartX"].ToString(), row["StartZ"].ToString(), row["StartY"].ToString() };
                            verticesList.Add(startVertexArray);
                        }
                        lastSegment = row["Segment"].ToString();

                        rowIdx++;
                        if (rowIdx == numberOfRows || row["Segment"].ToString() != dataTable.Rows[rowIdx]["Segment"].ToString())
                        {
                            endNodeName = row["Layer"].ToString().Replace(" ", "_") + "_End_" + row["Segment"].ToString() + "_" + row["Sequence"].ToString();
                            endNodeName.Replace(" ", "_");
                            endNodeName.Replace(".", "_");

                            // add end node
                            var nodeRow = nodesTable.Rows.Create();
                            nodeRow.Properties["Node"].Value = endNodeName;
                            nodeRow.Properties["XLoc"].Value = row["EndX"].ToString();
                            nodeRow.Properties["YLoc"].Value = row["EndZ"].ToString();
                            nodeRow.Properties["ZLoc"].Value = row["EndY"].ToString();

                            // add link
                            var linkRow = linksTable.Rows.Create();
                            linkRow.Properties["Link"].Value = linkName;
                            linkRow.Properties["StartingNode"].Value = startNodeName;
                            linkRow.Properties["EndingNode"].Value = endNodeName;

                            // add vertices
                            foreach (var array in verticesList)
                            {
                                var verticesRow = verticesTable.Rows.Create();
                                verticesRow.Properties["Link"].Value = array[0];
                                verticesRow.Properties["Sequence"].Value = array[1];
                                verticesRow.Properties["XLoc"].Value = array[2];
                                verticesRow.Properties["YLoc"].Value = array[3];
                                verticesRow.Properties["ZLoc"].Value = array[4];
                            }
                        }
                        else
                        {
                            string[] endVertexArray = { linkName, row["Sequence"].ToString(), row["EndX"].ToString(), row["EndZ"].ToString(), row["EndY"].ToString() };
                            verticesList.Add(endVertexArray);
                        }
                    }
                });

                MessageBox.Show("Create Links Completed", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            }
        }      
      
        #endregion

        #region IDesignAddInGuiDetails Members

        public string CategoryName
        {
            get { return "Table Tools"; }
        }

        public string TabName
        {
            get { return "AutoCAD"; }
        }

        public string GroupName
        {
            get { return "Intergration"; }
        }

        #endregion
    }

}
