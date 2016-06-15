
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ANDPI_Sales_Analysis
{
    class Property
    {
        private int salesPrice;
        private int assmtValue;
        private string parcelNo;
        private string neighborhood;
        private string salesDate;
        private double ratio;
        private bool isTarget;

        public Property(string parcel, string dateSold, string neighbor, int salesPrice, int assmtV, bool isSpecial )
        {
            this.salesPrice = salesPrice;
            salesDate = dateSold;
            assmtValue = assmtV;
            parcelNo = parcel;
            neighborhood = neighbor;
            
            isTarget = isSpecial;
        }
        public string getNeighborhood()
        {
            return neighborhood;
        }
        public string getYear()
        {
            return salesDate;
        }
        public double getRatio()
        {
            //corrects for forclorsures
            if (salesPrice < 10000)
                ratio = 0.0;
            else
                ratio = ((1.0) * assmtValue) / salesPrice;
            return ratio;
        }
        public double getRawRatio()
        {
            ratio = ((1.0) * assmtValue) / salesPrice;
            return ratio;
        }
        public bool isAND()
        {
            return isTarget;
        }
        public int getSalePrice()
        {
            return salesPrice;
        }

    }
}
