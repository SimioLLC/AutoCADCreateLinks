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
                    bool sortXYDesc = false;
                    DialogResult dr = MessageBox.Show("Sort Table By XY Desc?", "Sort By XY Desc?", MessageBoxButtons.YesNoCancel);

                    if (dr == DialogResult.Cancel)
                    {
                        MessageBox.Show("Canceled by user.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    else if (dr == DialogResult.Yes)
                    {
                        sortXYDesc = true;
                    }

                    var table = context.ActiveModel.Tables["AutoCADExport"];
                    var dataTable = AutoCADCreateLinksUtils.ConvertTableToDataTable(table);

                    DataView dataView = new DataView(dataTable);//datatable to dataview
                    if (sortXYDesc) dataView.Sort = "Layer ASC, Segment ASC, Sequence ASC, StartX DESC, StartY DESC";//string that contains the column name  followed by "ASC" (ascending) or "DESC" (descending)
                    else dataView.Sort = "Layer ASC, Segment ASC, Sequence ASC";
                    dataTable = dataView.ToTable();//push the chages back to the datatable;

                    // get tables
                    var nodesTable = context.ActiveModel.Tables["AutoCADNodes"];
                    var linksTable = context.ActiveModel.Tables["AutoCADLinks"];
                    var verticesTable = context.ActiveModel.Tables["AutoCADVertices"];

                    // clear tables
                    nodesTable.Rows.Clear();
                    linksTable.Rows.Clear();
                    verticesTable.Rows.Clear();

                    string lastLayer = String.Empty;
                    string lastSegment = String.Empty;
                    string lastSequence = String.Empty;
                    string startNodeName = String.Empty;
                    string endNodeName = String.Empty;
                    string linkName = String.Empty;
                    int numberOfRows = dataTable.Rows.Count;
                    int rowIdx = 0;
                    var verticesList = new List<string[]>();
                    String startX, startY, startZ, endX, endY, endZ = String.Empty;
                    foreach (DataRow row in dataTable.Rows) 
                    {
                        if (row["Name"].ToString() == "Line")
                        {
                            if ((Convert.ToBoolean(row["ForceEndToStart"].ToString()) == false) && ((Convert.ToDouble(row["StartX"].ToString()) < Convert.ToDouble(row["EndX"].ToString())) ||
                            (Convert.ToDouble(row["StartX"].ToString()) == Convert.ToDouble(row["EndX"].ToString()) &&
                            Convert.ToDouble(row["StartY"].ToString()) > Convert.ToDouble(row["EndY"].ToString()))))
                            {
                                startX = row["StartX"].ToString();
                                startY = row["StartY"].ToString(); 
                                startZ = row["StartZ"].ToString();
                                endX = row["EndX"].ToString();
                                endY = row["EndY"].ToString();
                                endZ = row["EndZ"].ToString();
                            }
                            else
                            {
                                startX = row["EndX"].ToString();
                                startY = row["EndY"].ToString();
                                startZ = row["EndZ"].ToString();
                                endX = row["StartX"].ToString();
                                endY = row["StartY"].ToString();
                                endZ = row["StartZ"].ToString();
                            }

                            if (lastLayer != row["Layer"].ToString() || lastSegment != row["Segment"].ToString())
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
                                nodeRow.Properties["XLoc"].Value = startX;
                                nodeRow.Properties["YLoc"].Value = startZ;
                                nodeRow.Properties["ZLoc"].Value = startY;
                            }
                            else
                            {
                                string[] startVertexArray = { linkName, row["Sequence"].ToString(), startX, startZ, startY };
                                verticesList.Add(startVertexArray);
                            }
                            lastLayer = row["Layer"].ToString();
                            lastSegment = row["Segment"].ToString();

                            rowIdx++;
                            if (rowIdx == numberOfRows || row["Layer"].ToString() != dataTable.Rows[rowIdx]["Layer"].ToString() || row["Segment"].ToString() != dataTable.Rows[rowIdx]["Segment"].ToString())
                            {
                                endNodeName = row["Layer"].ToString().Replace(" ", "_") + "_End_" + row["Segment"].ToString() + "_" + row["Sequence"].ToString();
                                endNodeName.Replace(" ", "_");
                                endNodeName.Replace(".", "_");

                                // add end node
                                var nodeRow = nodesTable.Rows.Create();
                                nodeRow.Properties["Node"].Value = endNodeName;
                                nodeRow.Properties["XLoc"].Value = endX;
                                nodeRow.Properties["YLoc"].Value = endZ;
                                nodeRow.Properties["ZLoc"].Value = endY;

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
                                // clear vertices
                                verticesList.Clear();
                            }
                            else
                            {
                                string[] endVertexArray = { linkName, row["Sequence"].ToString(), endX, endZ, endY };
                                verticesList.Add(endVertexArray);
                            }
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
