using System;
using System.Collections.Generic;

public class DriverModel
{
    public string DriverID { get; set; }
    public DateTime PayPeriod { get; set; }
    public char Qualify { get; set; }
    public int PaperWork { get; set; }
    public int Expiration { get; set; }
    public int Complaint { get; set; }
    public DateTime UpdateTime { get; set; }

    public override string ToString()
    {
        return $"{DriverID} | {PayPeriod} | {Qualify} | {PaperWork} | {Expiration} | {Complaint} | {UpdateTime}";
    }

    public static List<DriverModel> Convert(IList<IList<object>> values)
    {
        List<DriverModel> result = new List<DriverModel>();
        foreach (var item in values)
        {
            result.Add(new DriverModel
            {
                DriverID = item[0].ToString(),
                PayPeriod = DateTime.Parse(item[1].ToString()),
                Qualify = item[2].ToString()[0],
                PaperWork = int.Parse(item[3].ToString()),
                Expiration = int.Parse(item[4].ToString()),
                Complaint = int.Parse(item[5].ToString()),
                UpdateTime = DateTime.Parse(item[6].ToString())
            });
        }

        return result;
    }
}