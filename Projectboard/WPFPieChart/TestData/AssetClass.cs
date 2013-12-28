using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace WPFPieChart
{
    public class AssetClass : INotifyPropertyChanged
    {

        private String myClass;

        public String Class
        {
            get { return myClass; }
            set {
                myClass = value;
                RaisePropertyChangeEvent("Class");
            }
        }

        private double variabel;

        public double Variabel
        {
            get { return variabel; }
            set {
                variabel = value;
                RaisePropertyChangeEvent("Variabel");
            }
        }



        public static List<AssetClass> ConstructTestData()
        {
            List<AssetClass> assetClasses = new List<AssetClass>();

            assetClasses.Add(new AssetClass(){Class="Cash", Variabel=1.56});
            assetClasses.Add(new AssetClass(){Class="Bonds", Variabel=2.92});
            assetClasses.Add(new AssetClass(){Class="Real Estate", Variabel=13.24});
            assetClasses.Add(new AssetClass(){Class="Foreign Currency", Variabel=16.44});
            assetClasses.Add(new AssetClass(){Class="Stocks; Domestic", Variabel=27.57});
            assetClasses.Add(new AssetClass(){Class="Stocks; Foreign", Variabel=50.03});
            
            return assetClasses;
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChangeEvent(String propertyName)
        {
            if (PropertyChanged!=null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            
        }

        #endregion
    }
}
