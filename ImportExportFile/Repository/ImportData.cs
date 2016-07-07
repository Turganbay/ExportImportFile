using ImportExportFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ImportExportFile.Repository
{
    public class ImportData
    {
        Repository repo;
        public ImportData() 
        {
            repo = new Repository();
        }

        public void Import(string fileLocation) 
        {

                    string excelConnectionString = string.Empty;

                    excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation
                          + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";


                    DataTable dt = new DataTable();
                    DataSet ds = new DataSet();

                    using (OleDbConnection excelConnection = new OleDbConnection(excelConnectionString))
                    {
                        if (excelConnection.State != ConnectionState.Open)
                            excelConnection.Open();

                        dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


                        String[] excelSheets = new String[dt.Rows.Count];
                        int t = 0;
                        //excel data saves in temp file here

                        foreach (DataRow row in dt.Rows)
                        {
                            excelSheets[t] = row["TABLE_NAME"].ToString();
                            t++;
                        }

                        string query = string.Format("SELECT * FROM[{0}]", excelSheets[0]);

                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection))
                        {
                            dataAdapter.Fill(ds);
                        }
                    }

                    List<Region> listRegions = new List<Region>();
                    List<Product> listProducts = new List<Product>();
                    List<Company> listCompany = new List<Company>();
                    List<Quantity> listQuantity = new List<Quantity>();


                    DateTime reportDate = new DateTime();

                    string containsRegion = "область";
                    string productHead = "Регион/Область";
                    string endOfFile = "Итого";
                    int headerRowIndex;

                    int RegionRowIndex = -1;
                    string eachRegionName = "";
                    int RegionID = 0;

                    int CompanyRowIndex = -1;
                    int CompanyID = 0;

                    Dictionary<int, int> ProductsID = new Dictionary<int, int>();

                    string nextNumber = "";

                    bool afterHead = false;

                    int headLine = -1;

                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string endLine = ds.Tables[0].Rows[i][0].ToString();

                        if (endLine.Contains(endOfFile))
                        {
                            break;
                        }

                        for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                        {
                            string number = ds.Tables[0].Rows[i][j].ToString();

                            if (j == 0)
                            {
                                nextNumber = ds.Tables[0].Rows[i][1].ToString();
                            }

                            double convNum = 0;
                            bool isNumber = Double.TryParse(number, out convNum);

                            if (number.Contains(containsRegion))
                            {
                                Region r = new Region();
                                r.name = number;
                                listRegions.Add(r);

                                RegionRowIndex = i;
                                eachRegionName = number;
                                RegionID = repo.getRegionID(number);

                            }
                            else if (number.Contains(productHead))
                            {
                                headLine = i;
                                afterHead = true;

                            }
                            else if (!isNumber && j == 0 && !String.IsNullOrEmpty(number) && !String.IsNullOrEmpty(nextNumber) && afterHead == true)
                            {
                                Company c = new Company();
                                c.name = number;
                                listCompany.Add(c);

                                CompanyID = repo.getCompanyID(number);
                            }
                            else if (!isNumber && !String.IsNullOrEmpty(number) && afterHead == false)
                            {
                                Regex r = new Regex(@"\d{2}.\d{2}.\d{4}");
                                Match m = r.Match(number);
                                if (m.Success)
                                {
                                    reportDate = Convert.ToDateTime(m.ToString());
                                }

                            }


                            if (isNumber && afterHead == true && RegionID > 0 && CompanyID > 0)
                            {
                                Quantity q = new Quantity();
                                q.region_id = RegionID;
                                q.company_id = CompanyID;
                                q.product_id = ProductsID[j];
                                q.quantity = Convert.ToDouble(number);
                                q.create_time = reportDate;
                                listQuantity.Add(q);

                            }


                            if (headLine > -1 && j > 0 && !String.IsNullOrEmpty(number))
                            {
                                Product p = new Product();
                                p.name = number;
                                listProducts.Add(p);
                                ProductsID[j] = repo.getProductID(number);

                            }



                        }

                        if (headLine > -1)
                        {
                            headerRowIndex = headLine;
                            headLine = -1;
                        }


                    }


                     //   repo.InsertRegions(listRegions);
                    //    repo.InsertProducts(listProducts);
                    //    repo.InsertCompany(listCompany);
                     //   repo.InsertQuantity(listQuantity);  

        }



    }
}