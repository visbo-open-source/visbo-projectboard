using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Windows.Media.Animation;
using ScottLogic.Shapes;
using ScottLogic.Controls.PieChart;
using ScottLogic;
using System.Windows.Markup;
using ProjectBoardDefinitions;

namespace WPFPieChart
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    /// 
    
    public partial class PieChartWindow : Window
    {
        public bool isFuture;
        public bool formerSetting = Module1.awinSettings.mppShowAmpel;
       
        public Dictionary<string, clsWPFPieValues> piechartInput;

        private ObservableCollection<AssetClass> classes;
        public List<Brush> MyBrushes { get; private set; }
        public String selectedItemName
        {
            get
            {
                CollectionView collectionView = (CollectionView)CollectionViewSource.GetDefaultView(this.DataContext);
                return ((AssetClass)collectionView.CurrentItem).Class;
            }
        }

        //showWPFDiagramm (nameofChart, piechartInput, top, left, width, height, selectedItem) 
        public PieChartWindow(Dictionary<string, clsWPFPieValues> piechartInput)
        {
            
            
            //this.keepSymbols.IsChecked;


            System.Drawing.Color tmp;
            // create our test dataset and bind it
            MyBrushes = new List<Brush>();
            List<AssetClass> assetClasses = new List<AssetClass>();
            foreach (KeyValuePair<string, clsWPFPieValues> kvp in piechartInput)
            {
                assetClasses.Add(new AssetClass() {Class = kvp.Key, Variabel = kvp.Value.value});
                tmp = System.Drawing.ColorTranslator.FromOle((int)kvp.Value.color);
                MyBrushes.Add(new SolidColorBrush(Color.FromArgb(tmp.A, tmp.R, tmp.G, tmp.B)));
            }
            this.Resources.Add("mycolors", MyBrushes.ToArray());

            this.Top = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.top];
            this.Left = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.left];
            InitializeComponent();
            Module1.awinSettings.mppShowAmpel = true;
            // wird hier zwar gesetzt, aber komischerweise nicht berücksichtigt 
            //this.Top = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.top];
            //this.Left = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.left];

            classes = new ObservableCollection<AssetClass>(assetClasses);
            this.DataContext = classes;

           


        }

        private void isClosing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
            Window parent = Window.GetWindow(this);
            CheckBox checkbox = (CheckBox)parent.FindName("keepSymbols");
            Module1.awinSettings.mppShowAmpel = formerSetting;
           
            
            Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.top] = this.Top;
            Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.left] = this.Left;


            if ((bool)checkbox.IsChecked)
            {
                // Symbole löschen
                Module1.awinDeleteProjectChildShapes(1);
            }
           
           
        }

        private void isloaded(object sender, RoutedEventArgs e)
        {
            //this.Top = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.top];
            //this.Left = Module1.frmCoord[(int)Module1.PTfrm.ziele, (int)Module1.PTpinfo.left];

        }

    }
}
