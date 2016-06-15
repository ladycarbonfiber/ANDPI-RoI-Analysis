using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ANDPI_Sales_Analysis
{
    class Neighborhood
    {
        private List<Property> PList;
        private List<Property> rList; //no ANDPI properties
        private string neighborhood;
        public bool ignore;//super hacky way to omit stuff
        public Neighborhood(string name)
        {
            PList = new List<Property>();
            rList = new List<Property>();
            neighborhood = name;
            ignore = false;

        }
        public void addProperty(Property p)
        {
            PList.Add(p);
            if(!p.isAND())
            {
                rList.Add(p);
            }
        }
        public double getAggRatio(string year, bool includeANDPI)//unuesd at the present needs work.
        {
            double averageSale = getAverageSalePrice(year, includeANDPI);
            double thirdWeight = 1.15;
            double twothirdsWeight = 1.35;
            double minusthirdWeight = .85;
            double minus2thirdWeight = .7;
            double propertyRatios = 0.0;
            int count = 0;
            foreach(Property p in PList)
            {
                if(p.isAND() && !includeANDPI)
                {}//if we are excluding passover any of the inclusive set
                else
                {
                    if (p.getYear().Contains(year))
                    {
                        double temp = 0.0;
                        if (p.getRatio() > 1.66 * averageSale)
                        {
                            temp = p.getRatio() * twothirdsWeight;
                        }
                        else if (p.getRatio() > 1.33 * averageSale)
                        {
                            temp = p.getRatio() * thirdWeight;
                        }
                        else if (p.getRatio() < .66 * averageSale)
                        {
                            temp = p.getRatio() * minusthirdWeight;
                        }
                        else if (p.getRatio() < .33 * averageSale)
                        {
                            temp = p.getRatio() * minus2thirdWeight;
                        }
                        else
                        {
                            temp = p.getRatio();
                        }
                        propertyRatios += temp;
                        count++;
                    }
                }
                
            }
            double aRatio =0.0;
            if(propertyRatios > 0)
                 aRatio = propertyRatios / count;
            if(aRatio >1)
            { }
            return aRatio;

        }
        public double getAverageSalePrice(string year, bool includeANDPI)
        {
            double sales = 0.0;
            int count = 0;
            foreach (Property p in PList)
            {
                if (p.isAND() && !includeANDPI)
                { }//if we are excluding passover any of the inclusive set
                else
                {
                    if (p.getYear().Contains(year))
                    {
                        if (p.getSalePrice() > 10000)
                        {
                            sales += p.getSalePrice();
                            count++;
                        }
                    }
                }

            }
            double aSales = 0.0;
            if (sales > 0)
                aSales = (sales * 1.0) / count;
           
            return aSales;
        }
        public void printANDP(string y)
        {
            System.Console.WriteLine(y);
            foreach (Property p in PList)
            {
                
                if(p.getYear().Contains(y) && p.isAND())
                {
                    System.Console.WriteLine(p.getSalePrice());
                    System.Console.ReadLine();
                }
                
            }
        }
        public int getTotalSales(string year, bool includeANDPI)
        {
            int total = 0;
            if(includeANDPI)
            {
                foreach (Property p in PList)
                {
                    if(p.getYear().Contains(year))
                        total += p.getSalePrice();
                }
            }
            else
            {
                foreach(Property p in rList)
                {
                    if(p.getYear().Contains(year))
                        total += p.getSalePrice();
                }
            }
            return total;
        }
        public double getPercantageBelowAverage(string year, bool includeANDPI)
        {
            double percent = 0.0;
            double aSales = getAverageSalePrice(year, includeANDPI);
            int count = 0;
            int tCount = 0;
            List<Property> revisedList = new List<Property>();
            if (includeANDPI)
                revisedList = PList;
            else
            {
                revisedList = rList;

            }
            foreach(Property p in revisedList)
            {
                if (p.getYear().Contains(year))
                {
                    tCount++;
                    if ((p.getSalePrice() < aSales))
                        count++;
                }
            }
            if (tCount == 0)
                percent = 0;
            else
            {
                percent = (count * (1.0) / tCount);
            }
            return percent;
        }
        public  int getSalesVolume(string year, bool includeANDPI)
        {
            int count = 0;
            if(includeANDPI)
            {
                foreach(Property p in PList)
                {
                    if(p.getYear().Contains(year))
                    {
                        count++;
                    }
                }
            }
            else
            {
                foreach(Property p in rList)
                {
                    if(p.getYear().Contains(year))
                    {
                        count++;
                    }
                }
            }


            return count;
        }
        public double getPercentANDPI(string year, bool includeANDPI)
        {
            double percent = 0.0;
            int count = 0;
            int tCount = 0;
            if (!includeANDPI)
                return percent;
            
            else
            {
                foreach (Property p in PList)
                {
                    if (p.getYear().Contains(year))
                    {

                        tCount++;
                        if (p.isAND())
                            count++;
                    }

                }
            }
            if (tCount == 0)
                return 0.0;
            percent = (count *1.0)/ tCount;
            return percent;
        }
        public int getANDPISalesCount(string year)
        {
            int homeCount = 0;
            foreach (Property p in PList)
            {
                if (p.getYear().Contains(year) && p.isAND())
                {
                    homeCount++;
                }
            }
            return homeCount;

        }

        public bool hasANPI(string year)
        {
            return (getANDPISalesCount(year) > 0);
        }
        public string getName()
        {
            return neighborhood;
        }
        
    }
}
