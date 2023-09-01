namespace Svero.HelloExcel
{
    public class BankTransaction
    {
        public long Id { get; set; }

        public DateTime PlanDate { get; set; }

        public DateTime? ActualDate { get; set; }

        public string? Description { get; set; }

        public decimal Value { get; set; }

        public override string ToString()
        {
            return $"BankTransaction(Id: {Id}, PlanDate: {PlanDate.ToShortDateString()}, " + 
                $"ActualDate: {ActualDate?.ToShortDateString()}, Description: {Description}, Value: {Value})";
        }
    }
}