using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ANDPI_Sales_Analysis
{
    class ExcelManager
    {
        private string file;// = {"2009.xls", "2010.xls", "2011.xls","2012.xls", "2013.xls"}; //known
        private string[] parcels;
        private string[] date;
        private string[] years;
        private Dictionary<string, int> neighborhoodParcelCounts;
        private List<Property> andpProp;
        public ExcelManager(string inputItems, string inputPfile)
        {
            years =new string[] { "2008", "2009", "2010", "2011", "2012","2013" };//included years edit as need be
            neighborhoodParcelCounts = new Dictionary<string, int>();
            generateNeighborhoodCounts("parcels.txt");
            file = inputPfile;
            string ANDPIProperties = inputItems; //"input.txt";
            List<string> ANDPI = new List<string>();
            andpProp = new List<Property>();
            StreamReader reader = new StreamReader(ANDPIProperties);
            
            try
            {
                while (reader.Peek() != -1)
                {

                    ANDPI.Add(reader.ReadLine());
                }

            }
            catch
            {
                //TODO crash gracefully
            }
            finally
            {
                reader.Close();
            }
            parcels = new string[ANDPI.Count];
            date = new string[ANDPI.Count];
            int count = 0;
            string[] temp;
            foreach (string s in ANDPI)
            {
                //string format is known so this is reasonable
                temp = s.Split('\t');
                parcels[count] = temp[0];
                date[count] = temp[1];
                count++;
            }
 
        }
        private List<Neighborhood> getNeighborhoods()
        {
            List<Neighborhood>NList = new List<Neighborhood>();
            List<Property> allP = generateProperty();
            string loc = "";
            foreach(Property p in allP)
            {
                loc = p.getNeighborhood();
                if(NList.Count == 0)
                {
                    NList.Add(new Neighborhood(loc));
                }
                bool added = false;
                foreach(Neighborhood n in NList)
                {
                    if(String.Equals(loc, n.getName()))
                    {
                        n.addProperty(p);
                        added = true;
                    }
                }
                if(!added)
                {
                    NList.Add(new Neighborhood(loc));
                    NList.Last<Neighborhood>().addProperty(p);
                }

            }
            return NList;


        }
        private List<Property> generateProperty()
        {
           

            List<Property> pList = new List<Property>();
            StreamReader reader = new StreamReader(file);
            string inString = "";
            string [] temp;
            bool isANDPI;
            while(reader.Peek() != -1)
            {
                isANDPI = false;
                inString = reader.ReadLine();
                temp = inString.Split('\t');
                for(int i = 0; i <parcels.Length; i++)
                {
                    if(temp[0].Contains(parcels[i]) && temp[1].Contains(date[i]))
                    {
                        isANDPI = true;
                    }
                }
                Property input = new Property(temp[0], temp[1], temp[2], Convert.ToInt32(temp[3]), Convert.ToInt32(temp[4]), isANDPI);
                //terrible behavior just to quicking grab these
                if(isANDPI)
                    andpProp.Add(input);
                //comment out here
                pList.Add(input);
            }
            //also terrible behavior
            using(StreamWriter writer = new StreamWriter("ANDPI.txt"))
            {
                foreach(Property p in andpProp)
                {
                    writer.WriteLine(p.getYear() + "\t" + p.getSalePrice()); 
                }
            }

            return pList;
            
        }
        public void generateExcel( bool ANDPI, string fileName)
        {
            string[] metrics = { "Average Weighted Sales Ratio", "Average Sales Price", "Percent Sales Below Average", "Sales Volume", "Percent Sales ANDPI" }; //included metrics
            int firstSet = 2; //row location of first set of names
            List<Neighborhood> Nlist = getNeighborhoods();
            Application excelApp = new Application();
            excelApp.Visible = true;
            _Workbook workbook = (_Workbook)(excelApp.Workbooks.Add(Type.Missing));
            _Worksheet worksheet = (_Worksheet)workbook.ActiveSheet;
            int count = 2;
            for (int metricCount = 1; metricCount < metrics.Length+1; metricCount++)//excel cells are 1 indexed don't ask me why
            {
                worksheet.Cells[metricCount + (((years.Length+2) * (metricCount - 1))), 1] = metrics[metricCount - 1];
                worksheet.Cells[metricCount + 1+ ((years.Length+2) * (metricCount - 1)), 1] = "Year/NeighborhoodCode";
                for(int  yearCount = 0; yearCount <years.Length; yearCount++ )
                {
                    worksheet.Cells[metricCount + 2 + ((years.Length+2) * (metricCount-1)) + yearCount, 1] = years[yearCount];
                }
                
            }
            foreach(Neighborhood n in Nlist)
            {
                
                int metricCount = 0;
                worksheet.Cells[firstSet + ((years.Length+3) *metricCount), count] = n.getName();
                for (int yearCount = 0; yearCount < years.Length; yearCount++)
                {
                    worksheet.Cells[(firstSet + 1) + metricCount + (metricCount * (years.Length + 2)) + yearCount, count] = n.getAggRatio(years[yearCount], ANDPI);
                }

                metricCount++;
                worksheet.Cells[firstSet + ((years.Length + 3) * metricCount), count] = n.getName();
                for (int yearCount = 0; yearCount < years.Length; yearCount++)
                {
                    
                    worksheet.Cells[(firstSet + 1) + metricCount + (metricCount * (years.Length+2)) + yearCount, count] = n.getAverageSalePrice(years[yearCount], ANDPI);
                }
                metricCount++;
                worksheet.Cells[firstSet + ((years.Length + 3) * metricCount), count] = n.getName();
                for (int yearCount = 0; yearCount < years.Length; yearCount++)
                {
                    worksheet.Cells[(firstSet + 1) + metricCount + (metricCount * (years.Length + 2)) + yearCount, count] = n.getPercantageBelowAverage(years[yearCount], ANDPI);
                }
                metricCount++;
                worksheet.Cells[firstSet + ((years.Length + 3) * metricCount), count] = n.getName();
                for (int yearCount = 0; yearCount < years.Length; yearCount++)
                {
                    worksheet.Cells[(firstSet + 1) + metricCount + (metricCount * (years.Length + 2)) + yearCount, count] = n.getSalesVolume(years[yearCount], ANDPI);
                }
                metricCount++;
                worksheet.Cells[firstSet + ((years.Length + 3) * metricCount), count] = n.getName();
                for (int yearCount = 0; yearCount < years.Length; yearCount++)
                {
                    worksheet.Cells[(firstSet + 1) + metricCount + (metricCount * (years.Length + 2)) + yearCount, count] = n.getPercentANDPI(years[yearCount], ANDPI);
                }
               
                count++;
            }
            worksheet.SaveAs(fileName);

        }
        public void generateBulkExcel(string fileName)
        {
           // double millageRate = 9.9*.1;
            List<Neighborhood> Nlist = getNeighborhoods();
            Application excelApp = new Application();
            excelApp.Visible = true;
            _Workbook workbook = (_Workbook)(excelApp.Workbooks.Add(Type.Missing));
            _Worksheet worksheet = (_Worksheet)workbook.ActiveSheet;
            //Headers
            string[] headers = { "Years", "Total Qualified Sales", "Total ANDPI Sales", "Total Neighborhoods with ANDPI Involvement", "Avgerage Market Value of Homes Sold: ANDPI Included", "Average Market Value of Homes Sold: ANDPI Excluded", "Weighted Average Differential Per Home", "Total Homes in Neighborhood with ANDPI", "Projected Value Increase Across All Homes" };
                               //,"Tax Revenue Increase on Account of ANDPI  Involvement"};
            int headerStart = 1;
            foreach (string s in headers)
            {
                worksheet.Cells[1, headerStart] = s;
                headerStart++;
            }
            //bookeeping
            int yearStart = 2;
            int yearSep = 3;
            foreach (string y in years)
            {

                int ANDPISales = 0;
                int ANDPIDS = 0;
                int effectedCount = 0;
                double averageTotalValueANDPI = 0.0;
                double sumWith = 0;
                double sumWithDS = 0;
                double averageTotalValueNoANDPI = 0.0;
                double sumWithout = 0;
                double sumWithoutDS = 0;
                double averageValueDifferential = 0.0;

                double maximumProjectedValueIncrease = 0.0;
                double taxIncrease = 0.0;
                double weight = 0.0;
                List<double> Totals = new List<double>();
                List<double> Avg = new List<double>();

                //get total homes
                int totalHomesinANDPI = 0;
                int totalHomes = 0;
                int totalQualifiedSales = 0;
                int totalQualifiedSalesWithout = 0;
                int totalHomesDS = 0;
                // int totalQSDS = 0;
                // int totalQSWithoutDS = 0;
                foreach (Neighborhood n in Nlist)
                {
                    totalHomes += getNeighborhoodCount(n.getName());
                    // totalQualifiedSales += n.getSalesVolume(y, true);
                    //totalQualifiedSalesWithout += n.getSalesVolume(y, false);
                    if (n.hasANPI(y))
                    {
                        totalQualifiedSales += n.getSalesVolume(y, true);
                        totalQualifiedSalesWithout += n.getSalesVolume(y, false);
                        totalHomesinANDPI += getNeighborhoodCount(n.getName());
                        effectedCount++;
                    }
                    //if (n.getName().Contains("1016")) specific neighborhood
                    //{
                    //   totalHomesDS = getNeighborhoodCount
                    //      (n.getName());
                    ///   totalQSDS = n.getSalesVolume(y, true);
                    //   totalQSWithoutDS = n.getSalesVolume(y, false);
                    // }

                }
                foreach (Neighborhood n in Nlist)
                {
                    if (n.hasANPI(y))
                    {
                        ANDPISales += n.getANDPISalesCount(y);
                        sumWith += n.getTotalSales(y, true);// n.getAverageSalePrice(y, true);// * (n.getSalesVolume(y, true));// / (totalQualifiedSales * 1.0));
                        sumWithout += n.getTotalSales(y, false); // n.getAverageSalePrice(y, false);// * (n.getSalesVolume(y, false));// / (totalQualifiedSalesWithout * 1.0));
                        // debug statment n.printANDP(y);
                    }
                }
                //double weighingFactorA = (totalQualifiedSalesWithout / (1.25 * totalQualifiedSales));
                //double weighingFactorB = 1-weighingFactorA;
                //double test2 = Math.Sqrt(Math.Pow(totalQualifiedSales,2)+Math.Pow(totalQualifiedSalesWithout,2));
                if (totalQualifiedSales != 0)
                {
                    averageTotalValueANDPI = sumWith / totalQualifiedSales;
                    /// (weighingFactorB * totalQualifiedSalesWithout + weighingFactorA * totalQualifiedSales);//(1.0*totalQualifiedSales);
                    averageTotalValueNoANDPI = sumWithout / (totalQualifiedSalesWithout);
                }
                else
                {
                    averageTotalValueNoANDPI = 0;
                    averageTotalValueANDPI = 0;//sumWithout / totalHomesinANDPI / (weighingFactorA * totalQualifiedSalesWithout + weighingFactorB * totalQualifiedSales);// (1.0 * totalQualifiedSalesWithout);
                }
                if (false)
                { 
                System.Console.WriteLine(ANDPISales);
                    System.Console.WriteLine(sumWith);
                    System.Console.WriteLine(sumWithout);
                    System.Console.WriteLine(averageTotalValueANDPI);
                    System.Console.WriteLine(averageTotalValueNoANDPI);
                    System.Console.Read();
            }

                averageValueDifferential = averageTotalValueANDPI - averageTotalValueNoANDPI;
               
                maximumProjectedValueIncrease = averageValueDifferential * totalHomesinANDPI;
    
                //taxIncrease = maximumProjectedValueIncrease * .4 * millageRate;
                //bookkeeping

                worksheet.Cells[yearStart, 1] = y;

                worksheet.Cells[yearStart, 2] = totalQualifiedSales;

                worksheet.Cells[yearStart, 3] = ANDPISales;

                worksheet.Cells[yearStart, 4] = effectedCount;
                worksheet.Cells[yearStart, 5] = averageTotalValueANDPI;

                worksheet.Cells[yearStart, 6] = averageTotalValueNoANDPI;

                worksheet.Cells[yearStart, 7] = averageValueDifferential;// (totalQualifiedSales / ((10.0) * ANDPISales)) * averageValueDifferential;

                worksheet.Cells[yearStart, 8] = totalHomesinANDPI;
             
                worksheet.Cells[yearStart, 9] = maximumProjectedValueIncrease;
               
                //worksheet.Cells[yearStart, 10] = taxIncrease; //will add back in once those values are better understood

                yearStart += yearSep;
            }
            worksheet.Cells[yearStart + 1, 1] = "Totals";
            worksheet.Cells[yearStart + 1, 2] = "=Sum(B2,B5,B8,B11,B14,C17)";
            worksheet.Cells[yearStart + 1, 3] = "=Sum(C2,C5,C8,C11,C14,C17)";
            worksheet.Cells[yearStart + 1, 8] = "=Sum(H2,H5,H8,H11,H14,H17)";
            worksheet.Cells[yearStart + 1, 9] = "=Sum(I2,I5,I8,I11,I14,I17)";

            //worksheet.SaveAs(fileName);
        }
        private void generateNeighborhoodCounts(string Nfile)
        {
            StreamReader reader = new StreamReader(Nfile);
            string inString = "";
            string[] temp;
            while(reader.Peek() != -1)
            {
                inString = reader.ReadLine();
                temp = inString.Split('\t');
                neighborhoodParcelCounts.Add(temp[0], Convert.ToInt32(temp[1]));
            }
            reader.Close();
        }
       private int getNeighborhoodCount(string name)
        {
           int outV;
           neighborhoodParcelCounts.TryGetValue(name, out outV);
           return outV;

        }

    }
}
