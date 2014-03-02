using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using VBCollection = Microsoft.VisualBasic.Collection;
using ProjectBoardDefinitions;
using Microsoft.Office.Interop.Excel;

namespace WpfWindow
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class ListSelectionWindow : System.Windows.Window
    {
        /// <summary>
        /// in den ch.. Variablen stehen Informationen, wo das Chart erzeugt werden soll und um welchen Chart-Typ es sich handelt
        /// d.h. welche Methode aufgerufen werden soll 
        /// </summary>
        public double chTop;
        public double chLeft;
        public double chWidth;
        public double chHeight;
        public string chTyp;
        // kennung wird ggf an aufrufender Stelle gesetzt, so dass man bspweise bei Diagrammtyp(5) Meilenstein
        // unterscheiden kann zwischen Visualisieren und Summenchart
        public string kennung;
        public List<string> allItems { get; private set; }
        public VBCollection selectedItems { get; private set; }
        public Boolean isChecked { get; private set; }

        public ListSelectionWindow(VBCollection items, string title)
        {
            selectedItems = new VBCollection();
            allItems = new List<string>();

            if (title == "Phasen visualisieren")
            {
                this.Top = Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.top];
                this.Left = Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.left];
            }
            else if (title == "Rollen auswählen" || title == "Kostenarten auswählen")
            {
                this.Top = Module1.frmCoord[(int)Module1.PTfrm.listSelR, (int)Module1.PTpinfo.top];
                this.Left = Module1.frmCoord[(int)Module1.PTfrm.listSelR, (int)Module1.PTpinfo.left];
            }
            else
            {
                this.Top = Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.top];
                this.Left = Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.left];
            }

            InitializeComponent();

            populate(items);
            this.Title = title;
            this.title_label.Content = title;
        }

        public ListSelectionWindow(VBCollection items, string title, string checkbox) : this(items, title)
        {
            this.checkbox1.Content = checkbox;
            this.checkbox1.Visibility = Visibility.Visible;
        }

        private void populate(VBCollection items)
        {
            foreach (string item in items)
            {
                this.allItems.Add(item);
            }
        }

        private void submit(object sender, RoutedEventArgs e)
        {
            Module1.appInstance.EnableEvents = false;
            Module1.enableOnUpdate = false;

            if (this.listbox.SelectedItems.Count > 0)
            {

                 if (this.chTyp == Module1.DiagrammTypen[5])
                    if (this.kennung == "sum") 
                    {

                        if (this.checkbox1.IsChecked == false)
                        {
                            VBCollection myCollection = new VBCollection();
                            foreach (string name in this.listbox.SelectedItems)
                            {

                                myCollection.Add(name, name);

                            }
                            object repObj = null;
                            awinDiagrams.awinCreateprcCollectionDiagram(ref myCollection, ref repObj, chTop, chLeft,
                                                                           chWidth, chHeight, false, chTyp, false);
                        }
                        else
                        {
                            foreach (string name in this.listbox.SelectedItems)
                            {
                                VBCollection myCollection = new VBCollection();
                                // es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                                // wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....

                                myCollection.Add(name, name);
                                object repObj = null;
                                
                                awinDiagrams.awinCreateprcCollectionDiagram(ref myCollection, ref repObj, chTop, chLeft,
                                                                               chWidth, chHeight, false, chTyp, false);


                                chTop = chTop + chHeight + 2;

                            }
                        }


                        
                        Module1.appInstance.EnableEvents = true;
                        Module1.enableOnUpdate = true;

                       
                        this.isChecked = (bool)this.checkbox1.IsChecked;

                        // this.selectedItems.Clear(); dieses Kommando führt dazu, dass myCollection in der DiagramList wieder auf Null ist; 
                        // weil in awinCreateprcCollectionDiagramm per ref übergeben

                        selectedItems = new VBCollection();
                        this.listbox.SelectedItems.Clear();


                    }
                    else
                    {
                        int farbID = 4;
                        SortedList<string, string> nameList = new SortedList<string, string>();


                        if (this.checkbox1.IsChecked == true)
                        {
                            Module1.awinDeleteMilestoneShapes(1);
                        }

                        foreach (string name in this.listbox.SelectedItems)
                        {
                            nameList.Add(name, name);
                        }
                        
                        Projekte.awinZeichneMilestones(nameList, farbID, false);


                        Module1.appInstance.EnableEvents = true;
                        Module1.enableOnUpdate = true;

                        //awinCreateprcCollectionDiagram(this.selectedItems, repObj, chtop, chleft, chwidth, chheight, False, chtyp, False)
                        //this.DialogResult = true;

                        this.isChecked = (bool)this.checkbox1.IsChecked;

                        // this.selectedItems.Clear(); dieses Kommando führt dazu, dass myCollection in der DiagramList wieder auf Null ist; 
                        // weil in awinCreateprcCollectionDiagramm per ref übergeben

                        selectedItems = new VBCollection();
                        nameList.Clear();
                        this.listbox.SelectedItems.Clear();
                    }
                else if (this.chTyp == Module1.DiagrammTypen[6])
                {
                    
                    VBCollection myCollection = new VBCollection();
                    clsProjekt hproj = Module1.selectedProjekte.get_getProject(1);

                    foreach (string name in this.listbox.SelectedItems)
                    {
                        myCollection.Add(name, name);
                    }

                    // hier wird die Aktion durchgeführt 
                    object repObj = null;
                    Projekte.createMsTrendAnalysisOfProject(ref hproj, ref repObj, ref myCollection, this.chTop, this.chLeft, this.chHeight, this.chWidth);

                    //clsProjekt hproj = Module1.selectedProjekte.get_getProject(1);

                    Module1.appInstance.EnableEvents = true;
                    Module1.enableOnUpdate = true;

                    myCollection.Clear();
                    this.listbox.SelectedItems.Clear();


                }
                else if (this.Title == "Phasen visualisieren")
                {
                    int farbID = 4;
                    VBCollection myCollection = new VBCollection();

                    if (this.checkbox1.IsChecked == true)
                    {
                        Module1.awinDeleteMilestoneShapes(3);
                    }
                        
                    foreach (string name in this.listbox.SelectedItems)
                    {
                        myCollection.Add(name, name);
                    }

                    Projekte.awinZeichnePhasen(myCollection, farbID, false);

                    Module1.appInstance.EnableEvents = true;
                    Module1.enableOnUpdate = true;

                    myCollection.Clear();
                    this.listbox.SelectedItems.Clear();
                  

                }
                else
                {
                    Module1.awinLoescheChartsAtPosition(chLeft);

                    if (this.checkbox1.IsChecked == false)
                    {
                        VBCollection myCollection = new VBCollection();
                        foreach (string name in this.listbox.SelectedItems)
                        {
                            
                            myCollection.Add(name, name);                                               

                        }
                        object repObj = null;
                        awinDiagrams.awinCreateprcCollectionDiagram(ref myCollection, ref repObj, chTop, chLeft,
                                                                       chWidth, chHeight, false, chTyp, false);
                    }
                    else
                    {
                        foreach (string name in this.listbox.SelectedItems)
                        {
                            VBCollection myCollection = new VBCollection();
                            // es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                            // wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....

                            myCollection.Add(name, name);
                            object repObj = null;
                            awinDiagrams.awinCreateprcCollectionDiagram(ref myCollection, ref repObj, chTop, chLeft,
                                                                           chWidth, chHeight, false, chTyp, false);


                            chTop = chTop + chHeight + 2;

                        }
                    }
                    

                    //VBCollection myCollection = this.selectedItems;


                    //clsDiagramme tmpVar = Module1.DiagramList;

                    Module1.appInstance.EnableEvents = true;
                    Module1.enableOnUpdate = true;

                    //awinCreateprcCollectionDiagram(this.selectedItems, repObj, chtop, chleft, chwidth, chheight, False, chtyp, False)
                    //this.DialogResult = true;

                    this.isChecked = (bool)this.checkbox1.IsChecked;

                    // this.selectedItems.Clear(); dieses Kommando führt dazu, dass myCollection in der DiagramList wieder auf Null ist; 
                    // weil in awinCreateprcCollectionDiagramm per ref übergeben

                    selectedItems = new VBCollection();
                    this.listbox.SelectedItems.Clear();


                }


            }
            else
            {
                MessageBox.Show("bitte mind. 1 Element auswählen");
            }


            
            
        }

        private void isclosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.Title == "Phasen visualisieren")
            {
                Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.top] = this.Top;
                Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.left] = this.Left;
            }
            else if (this.Title == "Rollen auswählen" || this.Title == "Kostenarten auswählen")
            {
                Module1.frmCoord[(int)Module1.PTfrm.listSelR, (int)Module1.PTpinfo.top] = this.Top;
                Module1.frmCoord[(int)Module1.PTfrm.listSelR, (int)Module1.PTpinfo.left] = this.Left;
            }
            else
            {
                Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.top] = this.Top;
                Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.left] = this.Left;
            }
            
                        
        }

        //private void isloaded(object sender, RoutedEventArgs e)
        //{
        //    if (this.Title == "Phasen visualisieren")
        //    {
        //        this.Top = Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.top];
        //        this.Left = Module1.frmCoord[(int)Module1.PTfrm.listselP, (int)Module1.PTpinfo.left];
        //    }
        //    else if (this.Title == "Rollen auswählen" || this.Title == "Kostenarten auswählen") 
        //    {
        //        this.Top = Module1.frmCoord[(int)Module1.PTfrm.listSelR , (int)Module1.PTpinfo.top];
        //        this.Left = Module1.frmCoord[(int)Module1.PTfrm.listSelR, (int)Module1.PTpinfo.left];
        //    }
        //    else
        //    {
        //        this.Top = Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.top];
        //        this.Left = Module1.frmCoord[(int)Module1.PTfrm.listSelM, (int)Module1.PTpinfo.left];
        //    }
            
            
        //}
    }
}
